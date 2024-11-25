export type SharePointResult<T = never> = {
  type: 'result',
  value: T,
}

export type SharePointError = {
  type: 'error',
  statusCode: number,
  code: string,
  requestId: string,
  date: string,
  body: string,
}