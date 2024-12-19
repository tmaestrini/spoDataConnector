import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  DynamicDataSharedDepth,
  type IPropertyPaneConfiguration,
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
import { ApiSelector, GraphError, GraphResult, IRequestResult, IRequestResultType, SharePointError, SharePointResult } from './models/types';
import { ApiConnectorFactory } from './ApiConnectorFactory';
import PropertyPaneGroup from './PropertyPaneGroup';


export interface IGraphConnectorWebPartProps {
  sourceSelector: 'none' | 'dynamicData';
  dataSource?: DynamicProperty<undefined>;
  apiSelector: ApiSelector;

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

export default class ApiConnectorWebpart extends BaseClientSideWebPart<IGraphConnectorWebPartProps> implements IDynamicDataCallables {

  private graphClient: MSGraphClientV3;
  private connectorData: IRequestResult;
  private dataSourceValues: undefined;

  public render(): void {
    this.tryFetchDataSourceValues();

    const element: React.ReactElement = ApiConnectorFactory.createConnector(this.properties?.apiSelector, {
      properties: this.properties,
      dataSourceValues: this.dataSourceValues,
      graphClient: this.graphClient,
      sharePointClient: this.context.spHttpClient,

      onResponseResult: (data: GraphResult | SharePointResult) => {
        // console.log('Data result', data);
        if (data.type === IRequestResultType.Graph) {
          this.connectorData = (data as GraphResult).result;
        } else if (data.type === IRequestResultType.SharePoint) {
          this.connectorData = (data as SharePointResult).result;
        }

        delete this.connectorData.type; // delete type property for better readability
        this.context.dynamicDataSourceManager.notifyPropertyChanged('connectorData');
      },
      onResponseError: (data: GraphError | SharePointError) => {
        console.log('Data error', data);
        if (this.connectorData.result) delete this.connectorData.result;
        this.context.dynamicDataSourceManager.notifyPropertyChanged('connectorData');
      },
    });

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
    // Set default values
    if (!this.properties.sharePoint.api) this.properties.sharePoint.api = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`;

    // Initialize necessary services
    this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    this.context.dynamicDataSourceManager.initializeSource(this);
    return Promise.resolve();
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      { id: 'connectorData', title: `Data result (${this.context.instanceId})` },
    ]
  }

  public getPropertyValue(propertyId: string): IRequestResult {
    if (propertyId === 'connectorData') return this.connectorData;
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

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const configGroup = new PropertyPaneGroup(this.properties, this.dataSourceValues);

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
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
                    { key: ApiSelector.Graph, text: 'Call a Microsoft Graph API' },
                    { key: ApiSelector.SharePoint, text: 'Call a SharePoint API' },
                  ],
                }),
              ],
            },
            ...(this.properties.apiSelector === ApiSelector.Graph ? [configGroup.graphPropertyPaneGroup] : []),
            ...(this.properties.apiSelector === ApiSelector.SharePoint ? [configGroup.sharePointPropertyPaneGroup] : []),
          ],
        },
        {
          header: {
            description: "Information / Reference",
          },
          groups: [
            {
              groupName: "Webpart Infos",
              groupFields: [
                PropertyPaneLabel('webpartId', {
                  text: 'Webpart Id',
                }),
                PropertyPaneTextField('webPartIdValue', {
                  value: this.context.instanceId,
                  description: 'Id of this webpart to use in dynamic data source',
                  disabled: true,
                }),
                PropertyPaneLabel('webpartVersion', {
                  text: 'Version',
                }),
                PropertyPaneTextField('webpartVersionValue', {
                  value: this.context.manifest.version,
                  description: 'Current version of this webpart',
                  disabled: true,
                }),
              ]
            }
          ],
        }
      ],
    };
  }
}