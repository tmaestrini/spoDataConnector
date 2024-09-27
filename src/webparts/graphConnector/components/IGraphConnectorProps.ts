import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IGraphConnectorProps {
  api: string;
  version: 'beta' | 'v1.0';
  filter?: string;
  select?: string;
  expand?: string;
  graphClient: MSGraphClientV3;

  onGraphData?(data: any): void;
}
