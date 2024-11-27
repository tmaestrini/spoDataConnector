import * as React from 'react';
import styles from './GraphConnector.module.scss';
import type { IGraphConnectorProps } from './IGraphConnectorProps';
import { GraphError, GraphResult } from '../models/types';
import * as Handlebars from 'handlebars';
import { Icon, MessageBar, MessageBarType, Stack } from '@fluentui/react';
import * as strings from 'GraphConnectorWebPartStrings';
import RequestResults from '../../../common/components/RequestResults';

const GraphConnector: React.FunctionComponent<IGraphConnectorProps> = (props) => {
  const [graphData, setGraphData] = React.useState<GraphResult>({} as GraphResult);
  const [apiError, setApiError] = React.useState<GraphError | undefined>(undefined);
  const [apiCall, setApiCall] = React.useState<string>();

  React.useEffect(() => {
    setApiError(undefined);
    loadDataFromGraph()
      .catch((e) => {
        console.error(e);
        setApiError(e.message);
      });
  }, [props]);

  async function loadDataFromGraph(): Promise<void> {
    function tryIngestDynamicData(template: string): string {
      if (!props.dataFromDynamicSource) return template;
      return Handlebars.compile(template)(props.dataFromDynamicSource);
    }

    const path = tryIngestDynamicData(props.api ?? 'me');

    let graphQuery = props.graphClient.api(path);
    if (props.version) graphQuery = graphQuery.version(props.version);
    if (props.select) graphQuery = graphQuery.select(props.select);
    if (props.expand) graphQuery = graphQuery.expand(props.expand);
    if (props.filter) graphQuery = graphQuery.filter(encodeURIComponent(tryIngestDynamicData(props.filter)));

    graphQuery.header('ConsistencyLevel', 'eventual');

    try {
      setApiCall(`${props.version}${path}`);
      const data = await graphQuery.get();

      setGraphData({ value: { ...data } } as GraphResult);
      if (props.onGraphDataResult) props.onGraphDataResult({ value: { ...data } } as GraphResult);
    } catch (error) {
      setGraphData({} as GraphResult);
      setApiError({ ...error} as GraphError);
      if (props.onGraphDataError) props.onGraphDataError({ ...error} as GraphError);
    }
  }

  return (
    <div className={styles.graphConnector}>
      <h2><Icon iconName="PlugConnected" /> Microsoft Graph API Connection</h2>
      <div>Graph api call: {apiCall && <code>{apiCall}</code>}</div>
      {apiError &&
              <Stack tokens={{ childrenGap: 1 }} style={{ margin: '1rem 0' }}>
              <MessageBar messageBarType={MessageBarType.error}>
                <div>Error in api call: <br /><span className={styles.error}>{apiError.body}</span></div>
              </MessageBar>
            </Stack>}
      {!apiError && <>
        <RequestResults data={graphData as GraphResult}
          dataFromDynamicSource={props.dataFromDynamicSource}
          labels={{ apiRequestResults: strings.GraphConnector.ShowGraphResultsLabel, dynamicDataResults: strings.DataSource.ShowDynamicDataLabel }} />
      </>}
    </div>
  );
}

export default GraphConnector;