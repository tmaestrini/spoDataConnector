import { IPropertyPaneGroup, PropertyPaneDropdown, PropertyPaneLabel, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import * as strings from 'GraphConnectorWebPartStrings';
import { IGraphConnectorWebPartProps } from "./ApiConnectorWebPart";
import { AuthSelector } from "./models/types";

export default class PropertyPaneGroup {

  private properties: IGraphConnectorWebPartProps;
  private dataSourceValues: undefined;

  constructor(properties: IGraphConnectorWebPartProps, dataSourceValues: undefined) {
    this.properties = properties;
    this.dataSourceValues = dataSourceValues;
  }

  public get graphPropertyPaneGroup(): IPropertyPaneGroup {
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

  public get sharePointPropertyPaneGroup(): IPropertyPaneGroup {
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
          placeholder: `https://{site}.sharepoint.com/_api/site, https://{site}.sharepoint.com/_api/lists/getbytitle('listname')`,
          description: `${this.properties.sharePoint?.api}`,
          multiline: true,
          rows: 4,
        }),
        PropertyPaneTextField('sharePoint.filter', {
          label: `${strings.SharePointAPI.FilterLabel} ${this.dataSourceValues ? '👇 use dynamic data' : ''}`,
          placeholder: `Title eq 'Alfred'`,
          multiline: true,
          description: `Also see reference: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/use-odata-query-operations-in-sharepoint-rest-requests#select-items-to-return`,
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

  public get entraIdAuthPropertyPaneGroups(): IPropertyPaneGroup[] {
    return [
      {
        groupName: "Authentication",
        groupFields: [
          PropertyPaneDropdown('authSelector', {
            label: 'Authentication Mode',
            options: [
              { key: AuthSelector.EntraIdApp, text: 'Entra ID app registration' },
              { key: AuthSelector.SPFx, text: 'built-in SFFx Authentication' },
            ],
          }),
          ...(this.properties.authSelector === AuthSelector.SPFx ? [PropertyPaneLabel('authSelectorWarning', {
            text: `⚠️ Warning: Built-in SFFx Authentication is not recommended for production use. This could lead to serious security vulnerabilities.`,
          })] : [])
        ]
      },
      ...(this.properties.authSelector === AuthSelector.EntraIdApp ? [{
        groupName: `Setup`,
        groupFields: [
          PropertyPaneTextField('entraId.appId', {
            label: `Entra ID App registration`,
            placeholder: `9fedeb5c-3ad0-4b7c-9f29-c70c52a4420b`,
            description: `Insert the client Id of your Entra ID app registration`,
          }),
          PropertyPaneTextField('entraId.scopes', {
            label: `API Scopes`,
            placeholder: ``,
            multiline: true,
            description: `Should match the permission scopes defined in your Entra ID app registration`,
          }),
        ],
      }] : [])
    ];
  }
}