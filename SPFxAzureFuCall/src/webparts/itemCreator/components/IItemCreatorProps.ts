import { AadHttpClientFactory } from '@microsoft/sp-http';
export interface IItemCreatorProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string; 
  aadFactory: AadHttpClientFactory;
}
