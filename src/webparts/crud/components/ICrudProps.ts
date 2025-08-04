import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICrudProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  ListName:string;
  siteurl:string;
 context:WebPartContext;
}
