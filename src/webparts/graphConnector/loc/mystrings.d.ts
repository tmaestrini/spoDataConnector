declare interface IGraphConnectorWebPartStrings {
  PropertyPaneDescription: string;
  DataSource: {
    GroupNameLabel: string;
    DataSourceDescriptionText: string;
    SourceSelectorLabel: string;
    DynamicDataLabel: string;
  },
  GraphAPI: { 
    BasicGroupName: string;
    MainDescriptionText: string;
    ApiLabel: string;
    FilterLabel: string;
    SelectLabel: string;
    ExpandLabel: string;
    VersionLabel: string;
  }
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
