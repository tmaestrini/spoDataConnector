import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import * as strings from 'GraphConnectorWebPartStrings';
import GraphConnector from './components/GraphConnector';
import { IGraphConnectorProps } from './components/IGraphConnectorProps';
import { GraphError, GraphResult } from './models/types';

export interface IGraphConnectorWebPartProps {
  api: string;
  version: 'v1.0' | 'beta';
  filter?: string;
  select?: string;
  expand?: string;
}

export default class GraphConnectorWebPart extends BaseClientSideWebPart<IGraphConnectorWebPartProps> implements IDynamicDataCallables {

  private graphClient: MSGraphClientV3;
  private graphData: GraphResult;

  public render(): void {
    const element: React.ReactElement<IGraphConnectorProps> = React.createElement(
      GraphConnector,
      {
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('api', {
                  label: strings.graphApiLabel,
                  placeholder: '/me, /me/manager, /me/joinedTeams, /users'
                }),
                PropertyPaneDropdown('version', {
                  label: strings.graphVersionLabel,
                  options: [
                    { key: 'v1.0', text: 'v1.0' },
                    { key: 'beta', text: 'beta' },
                  ],
                }),
                PropertyPaneTextField('filter', {
                  label: strings.graphFilterLabel,
                  placeholder: `emailAddress eq 'jon@contoso.com'`
                }),
                PropertyPaneTextField('select', {
                  label: strings.graphSelectLabel,
                  placeholder: 'givenName,surname'
                }),
                PropertyPaneTextField('expand', {
                  label: strings.graphExpandLabel,
                  placeholder: 'members',
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

}
