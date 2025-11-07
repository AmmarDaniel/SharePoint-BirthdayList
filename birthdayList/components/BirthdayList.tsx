import * as React from "react";
import styles from "./BirthdayList.module.scss";
import { IBirthdayListProps } from "./IBirthdayListProps";
import { IBirthdayListState } from "./IBirthdayListState";
import BirthdayService from "../BirthdayService";
import BirthdayItem from "./BirthdayItem";

const monthsLong = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

type CSSVars = React.CSSProperties & { ["--rows"]?: string };

const BirthdayList: React.FC<IBirthdayListProps> = (props) => {
  const [state, setState] = React.useState<IBirthdayListState>({
    loading: true,
    items: [],
    currentMonthName: monthsLong[new Date().getMonth()],
    currentYear: new Date().getFullYear(),
  });

  const serviceRef = React.useRef<BirthdayService | null>(null);
  if (!serviceRef.current) {
    serviceRef.current = new BirthdayService(props.context);
  }

  const load = React.useCallback(async (): Promise<void> => {
    setState((prev) => ({ ...prev, loading: true, error: undefined }));
    try {
      const items = await serviceRef.current!.getCurrentMonthBirthdays(
        props.listTitle,
        props.maxItems,
        props.showAge,
        props.imageFieldInternalName
      );
      setState((prev) => ({ ...prev, loading: false, items }));
    } catch (e: unknown) {
      const message =
        e instanceof Error ? e.message : "Failed to load birthdays";
      setState((prev) => ({ ...prev, loading: false, error: message }));
    }
  }, [
    props.listTitle,
    props.maxItems,
    props.showAge,
    props.imageFieldInternalName,
  ]);

  React.useEffect(() => {
    load().catch(() => {});
  }, [load]);

  const viewportStyle: CSSVars = { ["--rows"]: String(props.maxItems) };

  return (
    <div className={styles.wrapper}>
      <div className={styles.wrapperHeader}>
        <div className={styles.logo}>
          <img src={props.headerIconUrl} width={40} height={40} alt="" />
        </div>
        <div className={styles.caption}>
          <p>
            This Month&apos;s Birthdays
            <span>
              {state.currentMonthName} {state.currentYear}
            </span>
          </p>
        </div>
      </div>

      <div className={styles.wrapperContent}>
        {state.loading && (
          <div className={styles.loading} role="status" aria-live="polite">
            Loading birthdays…
          </div>
        )}
        {!state.loading && state.error && (
          <div className={styles.error} role="alert">
            {state.error}
          </div>
        )}
        {!state.loading && !state.error && state.items.length === 0 && (
          <div className={styles.empty}>
            No birthdays found for {state.currentMonthName}.<br />
            Check the list: <strong>{props.listTitle}</strong>.
          </div>
        )}

        {!state.loading && !state.error && state.items.length > 0 && (
          <div className={styles.listViewport} style={viewportStyle}>
            <ul className={styles.list} role="list">
              {state.items.map((item) => (
                <BirthdayItem
                  key={item.id}
                  item={item}
                  showDepartment={props.showDepartment}
                  showAge={props.showAge}
                />
              ))}
            </ul>
          </div>
        )}
      </div>
    </div>
  );
};

export default BirthdayList;
