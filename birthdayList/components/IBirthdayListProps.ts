import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBirthdayListProps {
  context: WebPartContext;
  listTitle: string;
  showDepartment: boolean;
  showAge: boolean;
  maxItems: number;
  headerIconUrl: string;
  imageFieldInternalName: string;
}
