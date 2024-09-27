import * as React from 'react';
import styles from './GraphConnector.module.scss';
import type { IGraphConnectorProps } from './IGraphConnectorProps';
import { Icon } from '@fluentui/react';

const GraphConnector: React.FunctionComponent<IGraphConnectorProps> = (props) => {
  const [graphData, setGraphData] = React.useState<any>({});
  const [apiError, setApiError] = React.useState<string>();
  const [apiCall, setApiCall] = React.useState<string>();

  React.useEffect(() => {
    setApiError('');
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
      setGraphData(data);
      if (props.onGraphData) props.onGraphData(data);
    } catch (error) {
      setApiError(error.message);
    }
  }

  return (
    <div className={styles.graphConnector}>
      <h2><Icon iconName="PlugConnected" /> Graph Connector</h2>
      <div>Graph api call: {apiCall && <code>{apiCall}</code>}</div>
      {apiError && <div className={styles.error}>{apiError}</div>}

      {graphData && <div>👉 <code>{JSON.stringify(graphData['@odata.count'])}</code> records found. 
      See <code>value</code> property for results.</div>}
    </div>
  );
}

export default GraphConnector;