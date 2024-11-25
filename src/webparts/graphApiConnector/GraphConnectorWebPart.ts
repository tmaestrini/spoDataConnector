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
  PropertyPaneTextField
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

  api: string;
  version: 'v1.0' | 'beta';
  filter?: string;
  select?: string;
  expand?: string;
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
        api: this.properties.api,
        version: this.properties.version,
        filter: this.properties.filter,
        select: this.properties.select,
        expand: this.properties.expand,
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
              ],
            },
            {
              groupName: strings.GraphAPI.BasicGroupName,
              groupFields: [
                ...(this.properties.sourceSelector === 'dynamicData' ? [PropertyPaneLabel('dataSourceSelectedLabel', {
                  text: strings.GraphAPI.MainDescriptionText,
                })] : []),
                ...(this.properties.sourceSelector === 'dynamicData' ? [PropertyPaneLabel('dataSourceSelectedLabel', {
                  text: `Example: {{siteTitle}} or {{value}}.`,
                })] : []),
                PropertyPaneDropdown('version', {
                  label: strings.GraphAPI.VersionLabel,
                  options: [
                    { key: 'v1.0', text: 'v1.0' },
                    { key: 'beta', text: 'beta' },
                  ],
                }),
                PropertyPaneTextField('api', {
                  label: `${strings.GraphAPI.ApiLabel} ${this.dataSourceValues ? '👇 use dynamic data' : ''}`,
                  placeholder: '/me, /me/manager, /me/joinedTeams, /users',
                  description: `https://graph.microsoft.com${this.properties.api}`,
                  multiline: true,

                }),
                PropertyPaneTextField('filter', {
                  label: `${strings.GraphAPI.FilterLabel} ${this.dataSourceValues ? '👇 use dynamic data' : ''}`,
                  placeholder: `emailAddress eq 'john@contoso.com'`,
                  multiline: true,
                }),
                PropertyPaneTextField('select', {
                  label: strings.GraphAPI.SelectLabel,
                  placeholder: 'givenName,surname'
                }),
                PropertyPaneTextField('expand', {
                  label: strings.GraphAPI.ExpandLabel,
                  placeholder: 'members',
                }),
              ],
            },
          ],
        },
      ],
    };
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

}
