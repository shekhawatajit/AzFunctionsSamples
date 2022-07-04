import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IItemCreatorProps {
  context: WebPartContext;
  hasTeamsContext: boolean;
  ListTitle: string;
  ClientID: string;
  apiUrl: string;
  redirectUrl: string;
  ProvisionTemplate: string;
  SiteType: string;
}
export interface IItemCreatorState {
  Title: string;
  Description: string;
  OnwersIds: number[];
  OnwersSPNs: string[];
  MembersIds: number[];
  VisitorsIds: number[];
  Submitted: boolean;
  InProgess: boolean;
  ErrorMessage: string;
}
