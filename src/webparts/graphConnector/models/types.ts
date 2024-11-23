export type GraphResult<T = never> = {
  type: 'result',
  value: T,
}

export type GraphError = {
  type: 'error',
  statusCode: number,
  code: string,
  requestId: string,
  date: string,
  body: string,
}