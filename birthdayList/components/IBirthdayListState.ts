export interface IBirthdayItem {
  id: number;
  name: string;
  birthDate: Date;
  day: number;
  month: number; // 0-11
  monthShort: string;
  department?: string;
  photoUrl?: string;
  isToday: boolean;
  isPast: boolean;
  ageThisYear?: number;
}

export interface IBirthdayListState {
  loading: boolean;
  error?: string;
  items: IBirthdayItem[];
  currentMonthName: string;
  currentYear: number;
}
