import * as React from 'react';
import styles from './GraphConnector.module.scss';
import type { IGraphConnectorProps } from './IGraphConnectorProps';
import { Icon, Popup, PrimaryButton } from '@fluentui/react';
import { GraphError, GraphResult } from '../models/types';
import { prettyPrintJson } from 'pretty-print-json';

const GraphConnector: React.FunctionComponent<IGraphConnectorProps> = (props) => {
  const [graphData, setGraphData] = React.useState<GraphResult>({} as GraphResult);
  const [apiError, setApiError] = React.useState<GraphError | undefined>(undefined);
  const [apiCall, setApiCall] = React.useState<string>();
  const [isPopupVisible, setIsPopupVisible] = React.useState(false);

  React.useEffect(() => {
    setApiError(undefined);
    loadDataFromGraph()
      .catch((e) => {
        console.error(e);
        setApiError(e.message);
      });
  }, [props]);

  async function loadDataFromGraph(): Promise<void> {
    const path = props.api ?? 'me';
    let graphQuery = props.graphClient.api(path);
    if (props.version) graphQuery = graphQuery.version(props.version);
    if (props.select) graphQuery = graphQuery.select(props.select);
    if (props.expand) graphQuery = graphQuery.expand(props.expand);
    if (props.filter) graphQuery = graphQuery.filter(encodeURIComponent(props.filter));

    try {
      setApiCall(`${props.version}${props.api}`);
      const data = await graphQuery.get();

      setGraphData({ type: 'result', value: { ...data } } as GraphResult);
      if (props.onGraphDataResult) props.onGraphDataResult({ type: 'result', value: { ...data } } as GraphResult);
    } catch (error) {
      setGraphData({} as GraphResult);
      setApiError({ ...error, type: 'error' } as GraphError);
      if (props.onGraphDataResult) props.onGraphDataResult({ ...error, type: 'error' } as GraphError);
    }
  }

  return (
    <div className={styles.graphConnector}>
      <h2><Icon iconName="PlugConnected" /> Graph Connector</h2>
      <div>Graph api call: {apiCall && <code>{apiCall}</code>}</div>
      {apiError && <div className={styles.error}>Error in api call: <br/>{apiError.body}</div>}

      {graphData.type === 'result' && <>
        <div style={{ marginBottom: '1em' }}>
          👉 <code>{JSON.stringify(graphData.value['@odata.count'])}</code> records found.
          See <code>value</code> property in connected webparts for results.
        </div>
        <PrimaryButton onClick={() => setIsPopupVisible(!isPopupVisible)} text={`${isPopupVisible ? 'Hide results' : 'Show results'}`} />
        {isPopupVisible && (
          <Popup>
            <pre dangerouslySetInnerHTML={{ __html: prettyPrintJson.toHtml(graphData.value) }} />
          </Popup>
        )}
      </>}
    </div>
  );
}

export default GraphConnector;