import { MSGraphClientV3 } from '@microsoft/sp-http';
import { GraphError, GraphResult } from '../models/types';

export interface IGraphConnectorProps {
  api: string;
  version: 'beta' | 'v1.0';
  filter?: string;
  select?: string;
  expand?: string;
  graphClient: MSGraphClientV3;

  onGraphDataResult?(data: GraphResult | GraphError): void;
}
