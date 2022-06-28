import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IItemCreatorProps {
  context: WebPartContext;
  hasTeamsContext: boolean;
}
export interface IItemCreatorState {
  DataItems: any[];
}