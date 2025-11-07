export interface IBirthdayItem {
  id: number;
  name: string;
  birthDate: Date;
  day: number;
  month: number; // 0-11
  monthShort: string; // 'Jan'..'Dec'
  department?: string;
  photoUrl?: string;
  isToday: boolean;
  isPast: boolean;
  ageThisYear?: number;
  upn?: string; // for caching photos
}

export interface IBirthdayListNewState {
  loading: boolean;
  error?: string;
  items: IBirthdayItem[];
  currentMonthName: string;
  currentYear: number;
}
