import { MSGraphClientV3 } from '@microsoft/sp-http';
import { GraphError, GraphResult } from '../models/types';

export interface IGraphConnectorProps<T = never> {
  dataFromDynamicSource?: T;

  // Graph API request
  api: string;
  version: 'beta' | 'v1.0';
  filter?: string;
  select?: string;
  expand?: string;
  graphClient: MSGraphClientV3;

  // Callbacks
  onResponseResult?(data: GraphResult): void;
  onResponseError?(error: GraphError): void;
}
