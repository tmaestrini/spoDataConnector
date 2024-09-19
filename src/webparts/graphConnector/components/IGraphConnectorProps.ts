import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IGraphConnectorProps {
  api: string;
  version: 'beta' | 'v1.0';
  graphClient: MSGraphClientV3;
}
