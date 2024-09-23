declare interface IGraphConnectorWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  graphApiLabel: string;
  graphFilterLabel: string;
  graphSelectLabel: string;
  graphExpandLabel: string;
  graphVersionLabel: string;
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

declare module 'GraphConnectorWebPartStrings' {
  const strings: IGraphConnectorWebPartStrings;
  export = strings;
}
