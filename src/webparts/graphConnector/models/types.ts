export type GraphResult = {
  type: 'result',
  value: any,
}

export type GraphError = {
  type: 'error',
  statusCode: number,
  code: string,
  requestId: string,
  date: string,
  body: string,
}