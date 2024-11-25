define([], function () {
  return {
    PropertyPaneDescription: "Display content from SharePoint API and make it available for other webparts.",
    DataSource: {
      GroupNameLabel: "Data Source",
      DataSourceDescriptionText: `You can optionally connect to data sources (page environment or other webparts on this page that provide 
      data source functionality) by selecting 'Dynamic data (Internal data source)' in the dropdown field 'Type of content'.`,
      SourceSelectorLabel: "Type of content",
      DynamicDataLabel: "Dynamic data (Internal data source)",
    },
    SharePointApi: {
      BasicGroupName: "SharePoint API request",
      MainDescriptionText: `Ingest the desired attribute from the result of the selected data source in the API field.
                  Use the curly brackets {{...}} as a placeholder to insert the value.`,
      ApiLabel: "API endpoint",
      FilterLabel: "Filter ($filter attribute)",
      SelectLabel: "Select ($select attribute)",
      VersionLabel: "SharePoint version",
      ExpandLabel: "Expand ($expand attribute)",
    },
    SharePointConnector: {
      ShowSharePointResultsLabel: "Show result(s) from SharePoint",
      ShowDynamicDataLabel: "Show dynamic data"
    },
    AppLocalEnvironmentSharePoint: "The app is running on your local environment as SharePoint web part",
    AppLocalEnvironmentTeams: "The app is running on your local environment as Microsoft Teams app",
    AppLocalEnvironmentOffice: "The app is running on your local environment in office.com",
    AppLocalEnvironmentOutlook: "The app is running on your local environment in Outlook",
    AppSharePointEnvironment: "The app is running on SharePoint page",
    AppTeamsTabEnvironment: "The app is running in Microsoft Teams",
    AppOfficeEnvironment: "The app is running in office.com",
    AppOutlookEnvironment: "The app is running in Outlook",
    UnknownEnvironment: "The app is running in an unknown environment"
  }
});