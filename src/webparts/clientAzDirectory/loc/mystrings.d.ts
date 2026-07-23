declare interface IClientAzDirectoryWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ListTitleFieldLabel: string;
  ListTitleFieldDescription: string;
  NameFieldLabel: string;
  NameFieldDescription: string;
  SecondaryFieldLabel: string;
  SecondaryFieldDescription: string;
  LinkFieldLabel: string;
  LinkFieldDescription: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'ClientAzDirectoryWebPartStrings' {
  const strings: IClientAzDirectoryWebPartStrings;
  export = strings;
}
