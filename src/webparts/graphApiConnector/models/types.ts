export enum IRequestResultType {
  Graph = 'Graph',
  SharePoint = 'SharePoint'
}

export interface IRequestResult {
  type: IRequestResultType;
}

export type GraphResult<T = never> = IRequestResult & {
  type: IRequestResultType.Graph,
  value: T,
}

export interface GraphError {
  type: IRequestResultType.Graph,
  statusCode: number,
  code: string,
  requestId: string,
  date: string,
  body: string,
}

export type SharePointResult<T = never> = IRequestResult & {
  type: IRequestResultType.SharePoint,
  value: T,
}

export interface SharePointError {
  type: IRequestResultType.SharePoint,
  statusCode: number,
  code: string,
  requestId: string,
  date: string,
  body: string,
}

export enum ApiSelector {
  Graph = 'graphApi',
  SharePoint = 'sharePointApi'
}

