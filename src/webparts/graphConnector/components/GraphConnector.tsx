import * as React from 'react';
import styles from './GraphConnector.module.scss';
import type { IGraphConnectorProps } from './IGraphConnectorProps';
import { Icon } from '@fluentui/react';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

const GraphConnector: React.FunctionComponent<IGraphConnectorProps> = (props) => {
  const [graphData, setGraphData] = React.useState<MicrosoftGraph.User>({} as MicrosoftGraph.User);
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

  async function loadDataFromGraph() {
    const path = props.api ?? '/me';
    let graphQuery = props.graphClient.api(path);
    if (props.version) graphQuery = graphQuery.version(props.version);

    setApiCall(`${props.version}${props.api}`);
    const data = await graphQuery.get();
    console.log(data);
    setGraphData(data);
  }

  return (
    <div className={styles.graphConnector}>
      <h2><Icon iconName="PlugConnected" /> Graph Connector</h2>
      <div>Graph api call: {apiCall && <code>{apiCall}</code>}</div>
      {apiError && <div className={styles.error}>{apiError}</div>}

      {graphData && <div>Graph<pre>{JSON.stringify(graphData)}</pre></div>}
    </div>
  );
}

export default GraphConnector;