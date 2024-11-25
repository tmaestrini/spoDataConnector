import { MSGraphClientV3 } from '@microsoft/sp-http';
import { SharePointError, SharePointResult } from '../models/types';

export interface ISharePointConnectorProps<T = never> {
  dataFromDynamicSource?: T;

  // Graph API request
  api: string;
  version: 'beta' | 'v1.0';
  filter?: string;
  select?: string;
  expand?: string;
  sharePointClient: MSGraphClientV3;

  // Callbacks
  onSharePointDataResult?(data: SharePointResult | SharePointError): void;
}
