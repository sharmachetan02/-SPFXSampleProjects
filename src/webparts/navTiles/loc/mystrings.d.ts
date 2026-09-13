declare interface INavTilesWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  NavTilesTitle: string;
}

declare module 'NavTilesWebPartStrings' {
  const strings: INavTilesWebPartStrings;
  export = strings;
}
