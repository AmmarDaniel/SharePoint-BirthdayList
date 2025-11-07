import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IBirthdayItem } from "./components/IBirthdayListState";

// REST response envelope
interface ISpListItemsResponse<TItem> {
  value: TItem[];
}

// Hyperlink/Picture field shape
interface IHyperlinkFieldValue {
  Url: string;
  Description?: string;
}

// Modern Image (Thumbnail) field shape often returned by REST
export interface IImageFieldValue {
  type?: string;
  fileName?: string;
  serverUrl?: string;
  serverRelativeUrl?: string;
  id?: string;
  width?: number;
  height?: number;
}

// Base item shape; image column will be added dynamically via indexer
interface IBirthdayListItemBase {
  Id: number;
  Title: string;
  BirthDate: string; // REST returns string (ISO)
  Department?: string;
  PhotoUrl?: string | IHyperlinkFieldValue; // legacy/fallback column (optional)
}

// Because the image field internal name is dynamic, allow unknown keys
type IBirthdayListItem = IBirthdayListItemBase & Record<string, unknown>;

export default class BirthdayService {
  constructor(private context: WebPartContext) {}

  private escapeListTitleForUrl(listTitle: string): string {
    return listTitle.replace(/'/g, "''");
  }

  private getOrigin(): string {
    return new URL(this.context.pageContext.web.absoluteUrl).origin;
  }

  // Try to parse Image column value into a known shape (handles JSON string or object)
  private parseImageFieldValue(input: unknown): IImageFieldValue | undefined {
    if (!input) return undefined;

    // If SharePoint returned stringified JSON
    if (typeof input === "string") {
      try {
        const parsed = JSON.parse(input) as unknown;
        if (parsed && typeof parsed === "object") {
          const p = parsed as Partial<IImageFieldValue>;
          if (p.serverRelativeUrl || p.serverUrl) return p as IImageFieldValue;
        }
      } catch {
        // fall through
      }
      return undefined;
    }

    // If SharePoint returned an object
    if (typeof input === "object") {
      const obj = input as Partial<IImageFieldValue>;
      if (obj.serverRelativeUrl || obj.serverUrl) {
        return obj as IImageFieldValue;
      }
    }

    return undefined;
  }

  /*
    Normalize to a usable absolute image URL from:
     - modern Image (Thumbnail) field (object or stringified)
     - Hyperlink/Picture field
   */
  private extractImageUrl(imageFieldValue: unknown): string | undefined {
    // Modern Image column
    const img = this.parseImageFieldValue(imageFieldValue);
    if (img) {
      if (img.serverRelativeUrl) {
        return `${this.getOrigin()}${img.serverRelativeUrl}`;
      }
      if (img.serverUrl && img.serverRelativeUrl) {
        return `${img.serverUrl}${img.serverRelativeUrl}`;
      }
    }

    // Hyperlink/Picture fallback
    if (typeof imageFieldValue === "object" && imageFieldValue !== null) {
      const link = imageFieldValue as Partial<IHyperlinkFieldValue>;
      if (typeof link.Url === "string" && link.Url.length > 0) {
        return link.Url;
      }
    }
    if (typeof imageFieldValue === "string" && imageFieldValue.length > 0) {
      return imageFieldValue;
    }

    return undefined;
  }

  public async getCurrentMonthBirthdays(
    listTitle: string,
    _maxItemsViewport: number,
    showAge: boolean,
    imageFieldInternalName: string
  ): Promise<IBirthdayItem[]> {
    const now = new Date();
    const currentMonth = now.getMonth(); // 0-11
    const today = now.getDate();
    const currentYear = now.getFullYear();

    const safeListTitle = this.escapeListTitleForUrl(listTitle);
    const selectFields = [
      "Id",
      "Title",
      "BirthDate",
      "Department",
      "PhotoUrl", // legacy/fallback
      imageFieldInternalName, // modern Image column (dynamic)
    ];

    const endpoint =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('${safeListTitle}')/items` +
      `?$select=${selectFields.join(",")}` +
      `&$top=5000`;

    const res: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Failed to load items. HTTP ${res.status}. ${text}`);
    }

    const json = (await res.json()) as ISpListItemsResponse<
      Record<string, unknown>
    >;
    const values: IBirthdayListItem[] = (json?.value ??
      []) as IBirthdayListItem[];

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

    const items: IBirthdayItem[] = values
      .filter((v) => typeof v.BirthDate === "string" && v.BirthDate.length > 0)
      .map((v) => {
        const dob = new Date(v.BirthDate);
        const itemMonth = dob.getMonth();
        const itemDay = dob.getDate();

        const isThisMonth = itemMonth === currentMonth;
        const isToday = isThisMonth && itemDay === today;
        const isPast = isThisMonth && itemDay < today;

        const imageRaw = v[imageFieldInternalName];
        let photoUrl: string | undefined = this.extractImageUrl(imageRaw);

        if (!photoUrl && typeof v.PhotoUrl !== "undefined") {
          const link = v.PhotoUrl;
          photoUrl = typeof link === "string" ? link : link?.Url;
        }

        const ageThisYear =
          showAge && !Number.isNaN(dob.getTime())
            ? currentYear - dob.getFullYear()
            : undefined;

        const item: IBirthdayItem = {
          id: v.Id,
          name: v.Title,
          birthDate: dob,
          day: itemDay,
          month: itemMonth,
          monthShort: monthsShort[itemMonth],
          department: v.Department,
          photoUrl,
          isToday,
          isPast,
          ageThisYear,
        };

        return item;
      })
      .filter((i) => i.month === currentMonth)
      .sort((a, b) => {
        if (a.isPast !== b.isPast) return a.isPast ? 1 : -1; // past to bottom
        return a.day - b.day;
      });

    return items;
  }
}
