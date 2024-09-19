import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import * as strings from 'GraphConnectorWebPartStrings';
import GraphConnector from './components/GraphConnector';
import { IGraphConnectorProps } from './components/IGraphConnectorProps';

export interface IGraphConnectorWebPartProps {
  api: string;
  version: 'v1.0' | 'beta';
}

export default class GraphConnectorWebPart extends BaseClientSideWebPart<IGraphConnectorWebPartProps> {

  private graphClient: MSGraphClientV3;

  public render(): void {
    const element: React.ReactElement<IGraphConnectorProps> = React.createElement(
      GraphConnector,
      {
        api: this.properties.api,
        version: this.properties.version,
        graphClient: this.graphClient,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this.graphClient = await this.context.msGraphClientFactory.getClient('3');
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
                  label: strings.graphApiLabel
                }),
                PropertyPaneDropdown('version', {
                  label: strings.graphVersionLabel,
                  options: [
                    { key: 'v1.0', text: 'v1.0' },
                    { key: 'beta', text: 'beta' },
                  ],
                }),

              ]
            }
          ]
        }
      ]
    };
  }
}
