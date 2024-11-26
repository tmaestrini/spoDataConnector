import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  DynamicDataSharedDepth,
  type IPropertyPaneConfiguration,
  IPropertyPaneGroup,
  PropertyPaneDropdown,
  PropertyPaneDynamicField,
  PropertyPaneDynamicFieldSet,
  PropertyPaneLabel,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';
import { BaseClientSideWebPart, IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';
import { DynamicProperty, IReadonlyTheme } from '@microsoft/sp-component-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as strings from 'GraphConnectorWebPartStrings';
import GraphConnector from './components/GraphConnector';
import { IGraphConnectorProps } from './components/IGraphConnectorProps';
import { GraphError, GraphResult } from './models/types';


export interface IGraphConnectorWebPartProps {
  sourceSelector: 'none' | 'dynamicData';
  dataSource?: DynamicProperty<undefined>;
  apiSelector: 'graphApi' | 'sharePointApi';

  graph: {
    api: string;
    version: 'v1.0' | 'beta';
    filter?: string;
    select?: string;
    expand?: string;
  };

  sharePoint: {
    api: string;
    version: 'v1.0' | 'v2.0';
    filter?: string;
    select?: string;
    expand?: string;
  }
}

export default class GraphConnectorWebPart extends BaseClientSideWebPart<IGraphConnectorWebPartProps> implements IDynamicDataCallables {

  private graphClient: MSGraphClientV3;
  private graphData: GraphResult;
  private dataSourceValues: undefined;

  public render(): void {
    this.tryFetchDataSourceValues();

    const element: React.ReactElement<IGraphConnectorProps> = React.createElement(
      GraphConnector,
      {
        dataFromDynamicSource: this.dataSourceValues,
        api: this.properties.graph?.api,
        version: this.properties.graph?.version,
        filter: this.properties.graph?.filter,
        select: this.properties.graph?.select,
        expand: this.properties.graph?.expand,
        graphClient: this.graphClient,

        onGraphDataResult: (data: GraphResult | GraphError) => {
          if (data.type === 'result') {
            delete (data as { type?: string }).type; // delete type property for better readability
            this.graphData = data as GraphResult;
          } else {
            console.error(data);
          }

          this.context.dynamicDataSourceManager.notifyPropertyChanged('graphData');
        },
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private tryFetchDataSourceValues(): void {
    if (this.properties.sourceSelector === 'dynamicData') {
      this.dataSourceValues = this.properties.dataSource?.tryGetValue();
    } else {
      this.dataSourceValues = undefined;
    }
  }

  protected async onInit(): Promise<void> {
    this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    this.context.dynamicDataSourceManager.initializeSource(this);
    return Promise.resolve();
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      { id: 'graphData', title: 'Graph Data result' },
    ]
  }

  public getPropertyValue(propertyId: string): GraphResult {
    if (propertyId === 'graphData') return this.graphData;
    throw new Error(`property '${propertyId}' not found`);
  }

  // dynamic data method
  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'dataSource': {
        dynamicPropertyType: 'object'
      },
    };
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this.domElement.style.setProperty('--displayMode', this.displayMode === DisplayMode.Read ? 'none' : 'inherit');
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string): void {
    if (propertyPath === 'sourceSelector') {
      this.tryFetchDataSourceValues()
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.DataSource.GroupNameLabel,
              groupFields: [
                PropertyPaneLabel('dataSourceDescriptionLabel', {
                  text: strings.DataSource.DataSourceDescriptionText,
                }),
                PropertyPaneDropdown('sourceSelector', {
                  label: strings.DataSource.SourceSelectorLabel,
                  options: [
                    { key: 'none', text: 'Static content (None)' },
                    { key: 'dynamicData', text: 'Dynamic data (Internal data source)' },
                  ],
                }),
                ...(this.properties.sourceSelector === 'dynamicData' ? [PropertyPaneDynamicFieldSet({
                  label: "",
                  fields: [
                    PropertyPaneDynamicField('dataSource', {
                      label: strings.DataSource.DynamicDataLabel,
                      propertyValueDepth: DynamicDataSharedDepth.Property,
                      sourcesLabel: "Available data sources",
                    }),
                  ],
                })] : []),
                PropertyPaneDropdown('apiSelector', {
                  label: strings.DataSource.ApiSelectorLabel,
                  options: [
                    { key: 'graphApi', text: 'Call a Microsoft Graph API' },
                    { key: 'sharePointApi', text: 'Call a SharePoint API' },
                  ],
                }),

              ],
            },
            ...(this.properties.apiSelector === 'graphApi' ? [this.graphPropertyPaneGroup] : []),
            ...(this.properties.apiSelector === 'sharePointApi' ? [this.sharePointPropertyPaneGroup] : []),
          ],
        },
      ],
    };
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  private get graphPropertyPaneGroup(): IPropertyPaneGroup {
    return {
      groupName: strings.GraphAPI.BasicGroupName,
      groupFields: [
        ...(this.properties.sourceSelector === 'dynamicData' ? [PropertyPaneLabel('dataSourceSelectedLabel', {
          text: strings.DataSource.MainDescriptionText,
        })] : []),
        ...(this.properties.sourceSelector === 'dynamicData' ? [PropertyPaneLabel('dataSourceSelectedLabel', {
          text: `Example: {{siteTitle}} or {{value}}.`,
        })] : []),
        PropertyPaneDropdown('graph.version', {
          label: strings.GraphAPI.VersionLabel,
          options: [
            { key: 'v1.0', text: 'v1.0' },
            { key: 'beta', text: 'beta' },
          ],
        }),
        PropertyPaneTextField('graph.api', {
          label: `${strings.GraphAPI.ApiLabel} ${this.dataSourceValues ? '👇 use dynamic data' : ''}`,
          placeholder: '/me, /me/manager, /me/joinedTeams, /users',
          description: `https://graph.microsoft.com${this.properties.graph?.api}`,
          multiline: true,

        }),
        PropertyPaneTextField('graph.filter', {
          label: `${strings.GraphAPI.FilterLabel} ${this.dataSourceValues ? '👇 use dynamic data' : ''}`,
          placeholder: `emailAddress eq 'john@contoso.com'`,
          multiline: true,
        }),
        PropertyPaneTextField('graph.select', {
          label: strings.GraphAPI.SelectLabel,
          placeholder: 'givenName,surname'
        }),
        PropertyPaneTextField('graph.expand', {
          label: strings.GraphAPI.ExpandLabel,
          placeholder: 'members',
        }),
      ],
    };
  }

  private get sharePointPropertyPaneGroup(): IPropertyPaneGroup {
    return {
      groupName: strings.SharePointAPI.BasicGroupName,
      groupFields: [
        ...(this.properties.sourceSelector === 'dynamicData' ? [PropertyPaneLabel('dataSourceSelectedLabel', {
          text: strings.DataSource.MainDescriptionText,
        })] : []),
        ...(this.properties.sourceSelector === 'dynamicData' ? [PropertyPaneLabel('dataSourceSelectedLabel', {
          text: `Example: {{siteTitle}} or {{value}}.`,
        })] : []),
        PropertyPaneDropdown('sharePoint.version', {
          label: strings.SharePointAPI.VersionLabel,
          options: [
            { key: 'v1.0', text: 'v1.0' },
            { key: 'v2.0', text: 'v2.0' },
          ],
        }),
        PropertyPaneTextField('sharePoint.api', {
          label: `${strings.SharePointAPI.ApiLabel} ${this.dataSourceValues ? '👇 use dynamic data' : ''}`,
          placeholder: `/site, /web, /lists, /lists/getbytitle('listname')`,
          description: `https://{site}.sharepoint.com/_api${this.properties.sharePoint?.api}`,
          multiline: true,

        }),
        PropertyPaneTextField('sharePoint.filter', {
          label: `${strings.SharePointAPI.FilterLabel} ${this.dataSourceValues ? '👇 use dynamic data' : ''}`,
          placeholder: `substringof('Alfred', Title) eq true`,
          multiline: true,
          description: `See reference: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/use-odata-query-operations-in-sharepoint-rest-requests#select-items-to-return`,
        }),
        PropertyPaneTextField('sharePoint.select', {
          label: strings.SharePointAPI.SelectLabel,
          placeholder: 'Title,Products/Name'
        }),
        PropertyPaneTextField('sharePoint.expand', {
          label: strings.SharePointAPI.ExpandLabel,
          placeholder: 'Products/Name',
        }),
      ],
    };
  }
}