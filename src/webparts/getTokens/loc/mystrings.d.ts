declare interface IGetTokensWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  URLFieldLabel: string;
  SubscriptionFieldLabel:string;
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

declare module 'GetTokensWebPartStrings' {
  const strings: IGetTokensWebPartStrings;
  export = strings;
}
