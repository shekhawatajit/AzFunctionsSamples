import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IItemCreatorProps {
  context: WebPartContext;
  hasTeamsContext: boolean;
  ListTitle: string;
  ClientID: string;
  apiUrl: string;
  redirectUrl: string;
}
export interface IItemCreatorState {
  Title: string;
  Description: string;
  Onwers: number[];
  Members: number[];
  Visitors: number[];
  Submitted: boolean;
}
