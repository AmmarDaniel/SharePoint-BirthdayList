import * as React from "react";
import styles from "./BirthdayListNew.module.scss";
import { IBirthdayItem } from "./IBirthdayListNewState";

export interface IBirthdayItemProps {
  item: IBirthdayItem;
  showDepartment: boolean;
  showAge: boolean;
}

const BirthdayItem: React.FC<IBirthdayItemProps> = ({
  item,
  showDepartment,
  showAge,
}) => {
  const initials = item.name
    ? item.name
        .split(" ")
        .filter((s) => s.length > 0)
        .map((p) => p[0])
        .slice(0, 2)
        .join("")
        .toUpperCase()
    : "?";

  const liClass = [
    styles.card,
    item.isToday ? styles.cardToday : "",
    item.isPast ? styles.cardPast : "",
  ]
    .filter(Boolean)
    .join(" ")
    .trim();

  return (
    <li
      className={liClass}
      aria-label={`${item.name} birthday on ${item.monthShort} ${item.day}`}
    >
      <div className={styles.graphic} aria-hidden="true">
        {item.photoUrl ? (
          <img src={item.photoUrl} alt={`${item.name} photo`} />
        ) : (
          <div className={styles.initials} title={item.name}>
            {initials}
          </div>
        )}
      </div>
      <div className={styles.nameBlock}>
        <span className={styles.header}>{item.name}</span>
        {showDepartment && item.department && (
          <span className={styles.meta}>{item.department}</span>
        )}
      </div>
      <div className={styles.dateBlock} aria-hidden="true">
        <span className={styles.stat}>{item.day}</span>
        <span className={styles.sub}>{item.monthShort}</span>
        {showAge && typeof item.ageThisYear === "number" && (
          <span className={styles.age}>({item.ageThisYear})</span>
        )}
      </div>
    </li>
  );
};
export default BirthdayItem;
