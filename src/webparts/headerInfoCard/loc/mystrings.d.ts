declare interface IHeaderInfoCardWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DataGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  headerInfoCardTitleFieldLabel: string;
  headerInfoCardColorFieldLabel: string;
  headerInfoCardIconFieldLabel: string;
  headerInfoCardIconFieldDescription: string;
  headerInfoCardIconDarkToggleFieldLabel: string;
  headerInfoCardIconDarkToggleOnText: string;
  headerInfoCardIconDarkToggleOffText: string;
  listFieldLabel: string;
  dataFilterFieldLabel: string;
  dataFilterFieldPlaceholder: string;
  dataFilterFieldDescription: string;
}

declare module 'HeaderInfoCardWebPartStrings' {
  const strings: IHeaderInfoCardWebPartStrings;
  export = strings;
}
