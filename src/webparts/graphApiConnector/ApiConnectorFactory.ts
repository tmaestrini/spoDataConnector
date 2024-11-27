import * as React from "react";
import GraphConnector from './components/GraphConnector';
import { IGraphConnectorProps } from "./components/IGraphConnectorProps";
import { ApiSelector, GraphResult, GraphError, SharePointResult, SharePointError } from "./models/types";
import { IGraphConnectorWebPartProps } from "./GraphConnectorWebPart";
import { MSGraphClientV3, SPHttpClient } from '@microsoft/sp-http';
import SharePointConnector from "./components/SharePointConnector";
import { ISharePointConnectorProps } from "./components/ISharePointConnectorProps";


interface IApiConnectorFactory {
  properties: IGraphConnectorWebPartProps;
  dataSourceValues: undefined;
  graphClient: MSGraphClientV3;
  sharePointClient: SPHttpClient;
  onDataResult: (data: GraphResult | SharePointResult) => void;
  onDataError: (data: GraphError | SharePointError) => void;
}

// create a new factory class to create a new instance of React.ReactElement
export class ApiConnectorFactory {

  public static createConnector(apiSelector: ApiSelector,
    data: IApiConnectorFactory): React.ReactElement<IGraphConnectorProps | ISharePointConnectorProps> {
    const { properties, dataSourceValues, graphClient, sharePointClient, onDataResult, onDataError } = data;

    function createGraphConnectorElement(): React.ReactElement<IGraphConnectorProps> {
      const element: React.ReactElement<IGraphConnectorProps> = React.createElement<IGraphConnectorProps>(
        GraphConnector,
        {
          dataFromDynamicSource: dataSourceValues,
          api: properties.graph?.api,
          version: properties.graph?.version,
          filter: properties.graph?.filter,
          select: properties.graph?.select,
          expand: properties.graph?.expand,
          graphClient: graphClient,

          onGraphDataResult: onDataResult,
          onGraphDataError: onDataError,
        }
      );
      return element;
    }

    function createSharePointConnectorElement(): React.ReactElement<ISharePointConnectorProps> {
      const element: React.ReactElement<ISharePointConnectorProps> = React.createElement<ISharePointConnectorProps>(
        SharePointConnector,
        {
          dataFromDynamicSource: dataSourceValues,
          api: properties?.sharePoint.api,
          version: properties.sharePoint?.version,
          filter: properties.sharePoint?.filter,
          select: properties.sharePoint?.select,
          expand: properties.sharePoint?.expand,
          sharePointClient: sharePointClient,

          onSharePointDataResult: onDataResult,
          onSharePointDataError: onDataError,
        }
      );
      return element;
    }

    let element: React.ReactElement | null = null;
    switch (apiSelector) {
      case ApiSelector.Graph:
        element = createGraphConnectorElement();
        break;
      case ApiSelector.SharePoint:
        element = createSharePointConnectorElement();
        break;
      default:
        throw new Error('Invalid API selector');
    }
    return element;
  }
}