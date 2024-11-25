declare interface ISharePointConnectorWebPartStrings {
  PropertyPaneDescription: string;
  DataSource: {
    GroupNameLabel: string;
    DataSourceDescriptionText: string;
    SourceSelectorLabel: string;
    DynamicDataLabel: string;
  };
  SharePointApi: { 
    BasicGroupName: string;
    MainDescriptionText: string;
    ApiLabel: string;
    FilterLabel: string;
    SelectLabel: string;
    ExpandLabel: string;
    VersionLabel: string;
  };
  SharePointConnector: {
    ShowSharePointResultsLabel: string;
    ShowDynamicDataLabel: string;
  };
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

declare module 'SharePointConnectorWebPartStrings' {
  const strings: ISharePointConnectorWebPartStrings;
  export = strings;
}
