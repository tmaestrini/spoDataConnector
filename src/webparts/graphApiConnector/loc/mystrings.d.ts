declare interface IGraphConnectorWebPartStrings {
  PropertyPaneDescription: string;
  DataSource: {
    GroupNameLabel: string;
    DataSourceDescriptionText: string;
    SourceSelectorLabel: string;
    DynamicDataLabel: string;
    ShowDynamicDataLabel: string;
    ApiSelectorLabel: string;
    MainDescriptionText: string;
  };
  GraphAPI: { 
    BasicGroupName: string;
    ApiLabel: string;
    FilterLabel: string;
    SelectLabel: string;
    ExpandLabel: string;
    VersionLabel: string;
  };
  GraphConnector: {
    ShowGraphResultsLabel: string;
    referencePropertyInfo?: string;
  };
  SharePointAPI: {
    BasicGroupName: string;
    MainDescriptionText: string;
    ApiLabel: string;
    FilterLabel: string;
    SelectLabel: string;
    ExpandLabel: string;
    VersionLabel: string;
  };
  SharePointConnector: {
    ShowSPOResultsLabel: string;
    referencePropertyInfo?: string;
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

declare module 'GraphConnectorWebPartStrings' {
  const strings: IGraphConnectorWebPartStrings;
  export = strings;
}
