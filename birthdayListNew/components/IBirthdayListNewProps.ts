import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBirthdayListNewProps {
  context: WebPartContext;
  listTitle: string;
  personFieldInternalName: string;
  showDepartment: boolean;
  showAge: boolean;
  maxItems: number;
  headerIconUrl: string;
}
