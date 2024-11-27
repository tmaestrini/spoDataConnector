import { SPHttpClient } from '@microsoft/sp-http';
import { SharePointError, SharePointResult } from '../models/types';

export interface ISharePointConnectorProps<T = never> {
  dataFromDynamicSource?: T;

  // Graph API request
  api: string;
  version: 'v1.0' | 'v2.0';
  filter?: string;
  select?: string;
  expand?: string;
  sharePointClient: SPHttpClient;

  // Callbacks
  onSharePointDataResult?(data: SharePointResult): void;
  onSharePointDataError?(error: SharePointError): void;
}
