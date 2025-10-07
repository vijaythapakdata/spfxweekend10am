import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IGetAttachmentProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  siteurl:string;
ListName:string;
context:WebPartContext;
}
