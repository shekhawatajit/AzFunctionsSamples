declare interface IItemCreatorWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ListTitleFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  ClientIDFieldLabel: string;
  apiUrlFieldLabel: string;
  redirectUrlFieldLabel: string;
}

declare module 'ItemCreatorWebPartStrings' {
  const strings: IItemCreatorWebPartStrings;
  export = strings;
}
