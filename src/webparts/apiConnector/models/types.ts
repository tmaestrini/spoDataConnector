export enum IRequestResultType {
  Graph = 'Graph',
  SharePoint = 'SharePoint'
}

export interface IRequestResult {
  type?: IRequestResultType;
  result: never;
}

export type GraphResult<T = never> = IRequestResult & {
  type: IRequestResultType.Graph,
  result: T,
}

export type GraphError = IRequestResult & {
  type: IRequestResultType.Graph,
  statusCode: number,
  code: string,
  requestId: string,
  date: string,
  body: string,
}

export type SharePointResult<T = never> = IRequestResult & {
  type: IRequestResultType.SharePoint,
  result: T,
}

export type SharePointError = IRequestResult & {
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

export enum AuthSelector {
  EntraIdApp = 'EntraIdApp',
  SPFx = 'SPFx'
}

