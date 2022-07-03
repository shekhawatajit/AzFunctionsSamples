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
  ProvisionTemplateFieldLabel: string;
  SiteTypeLabel: string;
  GroupWithTeams: string;
  GroupWithTeamsLabel: string;
  GroupWithoutTeams: string;
  GroupWithoutTeamsLabel: string;
}

declare module 'ItemCreatorWebPartStrings' {
  const strings: IItemCreatorWebPartStrings;
  export = strings;
}
