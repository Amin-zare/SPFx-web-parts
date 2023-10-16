declare interface ITemplateWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ToggleFieldLabel: string;
  MultiLineFieldLabel: string;
  RatingFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  ListTitleFieldLabel: string;
  ListNameFieldLabel: string;
  ItemNameFieldLabel: string;
}

declare module 'TemplateWebPartStrings' {
  const strings: ITemplateWebPartStrings;
  export = strings;
}
