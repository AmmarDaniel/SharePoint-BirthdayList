import {
  SPHttpClient,
  SPHttpClientResponse,
  MSGraphClientV3,
} from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IBirthdayItem } from "../components/IBirthdayListNewState";
import { ResponseType } from "@microsoft/microsoft-graph-client";

/** REST response envelope */
interface ISpListItemsResponse<TItem> {
  value: TItem[];
}

/** Person field (expanded) minimal shape */
interface IPersonValue {
  Id?: number;
  Title?: string; // display name
  EMail?: string; // user email/UPN in most tenants
  Name?: string; // claims login: e.g. "i:0#.f|membership|user@tenant.com"
}

// Base item shape returned by REST; person column is dynamic key
type IItem = {
  Id: number;
  BirthDate: string; // ISO string
} & Record<string, unknown>;

export default class BirthdayServiceNew {
  private _photoCache = new Map<string, string>(); // key: upn/login -> objectUrl
  private _graphClient?: MSGraphClientV3;

  constructor(private context: WebPartContext) {}

  private async getGraph(): Promise<MSGraphClientV3> {
    if (!this._graphClient) {
      this._graphClient = await this.context.msGraphClientFactory.getClient(
        "3"
      );
    }
    return this._graphClient;
  }

  // Clean up any created object URLs
  public dispose(): void {
    this._photoCache.forEach((url) => URL.revokeObjectURL(url));
    this._photoCache.clear();
  }

  public async getCurrentMonthBirthdays(
    listTitle: string,
    showAge: boolean,
    personFieldInternalName: string
  ): Promise<IBirthdayItem[]> {
    const now = new Date();
    const currentMonth = now.getMonth(); // 0-11
    const today = now.getDate();
    const currentYear = now.getFullYear();

    // --- SharePoint list query ---
    const safeListTitle = listTitle.replace(/'/g, "''");

    // Valid sub-properties for Person fields in list items:
    // Id, Title, EMail, Name (claims). Do NOT select Picture/UserPrincipalName here.
    const selectFields = [
      "Id",
      "BirthDate",
      `${personFieldInternalName}/Id`,
      `${personFieldInternalName}/Title`,
      `${personFieldInternalName}/EMail`,
      `${personFieldInternalName}/Name`,
    ];

    const endpoint =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('${safeListTitle}')/items` +
      `?$select=${selectFields.join(",")}` +
      `&$expand=${personFieldInternalName}` +
      `&$top=5000`;

    const res: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );
    if (!res.ok) {
      throw new Error(
        `Failed to load items. HTTP ${res.status}. ${await res.text()}`
      );
    }

    const json = (await res.json()) as ISpListItemsResponse<
      Record<string, unknown>
    >;
    const values = (json?.value ?? []) as IItem[];

    const monthsShort = [
      "Jan",
      "Feb",
      "Mar",
      "Apr",
      "May",
      "Jun",
      "Jul",
      "Aug",
      "Sep",
      "Oct",
      "Nov",
      "Dec",
    ];

    // --- Graph client ---
    const graph = await this.getGraph();

    const items: IBirthdayItem[] = await Promise.all(
      values
        .filter(
          (v) => typeof v.BirthDate === "string" && v.BirthDate.length > 0
        )
        .map(async (v) => {
          const dob = new Date(v.BirthDate);
          const person = v[personFieldInternalName] as IPersonValue | undefined;

          // Derive identifier for Graph
          // Prefer EMail (works in most tenants). Fallback to claims login stripped of prefix.
          const claims = person?.Name;
          const loginFromClaims = claims
            ? claims.replace(/^.*\|/, "")
            : undefined;
          const upnOrLogin =
            (person?.EMail || loginFromClaims || "").trim() || undefined;

          // Defaults from list in case Graph fails
          let displayName = person?.Title ?? "";
          let department: string | undefined;
          let photoUrl: string | undefined;

          // Fetch profile data
          if (upnOrLogin) {
            // 1) Name & department from Graph
            try {
              const user = await graph
                .api(`/users/${encodeURIComponent(upnOrLogin)}`)
                .select("displayName,department,id,mail,userPrincipalName")
                .get();
              displayName = user?.displayName || displayName;
              department = user?.department;
            } catch {
              // tolerate read issues (guest, missing scope, etc.)
            }

            // 2) Profile photo (48x48) — cached
            const cached = this._photoCache.get(upnOrLogin);
            if (cached) {
              photoUrl = cached;
            } else {
              try {
                const blob: Blob = await graph
                  .api(
                    `/users/${encodeURIComponent(
                      upnOrLogin
                    )}/photos/48x48/$value`
                  )
                  .responseType(ResponseType.BLOB)
                  .get(); // returns binary
                photoUrl = URL.createObjectURL(blob);
                this._photoCache.set(upnOrLogin, photoUrl);
              } catch {
                // user might not have a photo — fall back to initials in UI
              }
            }
          }

          const itemMonth = dob.getMonth();
          const itemDay = dob.getDate();
          const isThisMonth = itemMonth === currentMonth;
          const isToday = isThisMonth && itemDay === today;
          const isPast = isThisMonth && itemDay < today;

          const ageThisYear =
            showAge && !Number.isNaN(dob.getTime())
              ? currentYear - dob.getFullYear()
              : undefined;

          const item: IBirthdayItem = {
            id: v.Id,
            name: displayName || "(Unknown)",
            birthDate: dob,
            day: itemDay,
            month: itemMonth,
            monthShort: monthsShort[itemMonth],
            department,
            photoUrl,
            isToday,
            isPast,
            ageThisYear,
            upn: upnOrLogin,
          };
          return item;
        })
    );

    // Keep only current month; push past items to bottom; sort by day asc
    return items
      .filter((i) => i.month === currentMonth)
      .sort((a, b) => {
        if (a.isPast !== b.isPast) return a.isPast ? 1 : -1;
        return a.day - b.day;
      });
  }
}
