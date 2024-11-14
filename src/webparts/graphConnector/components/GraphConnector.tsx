import * as React from 'react';
import styles from './GraphConnector.module.scss';
import type { IGraphConnectorProps } from './IGraphConnectorProps';
import { GraphError, GraphResult } from '../models/types';
import { prettyPrintJson } from 'pretty-print-json';
import * as Handlebars from 'handlebars';
import { Icon, Popup, PrimaryButton } from 'office-ui-fabric-react';

const GraphConnector: React.FunctionComponent<IGraphConnectorProps> = (props) => {
  const [graphData, setGraphData] = React.useState<GraphResult>({} as GraphResult);
  const [apiError, setApiError] = React.useState<GraphError | undefined>(undefined);
  const [apiCall, setApiCall] = React.useState<string>();
  const [isGraphDataPopupVisible, setGraphDataPopupVisible] = React.useState(false);
  const [isDynamicDataPopupVisible, setDynamicDataPopupVisible] = React.useState(false);

  React.useEffect(() => {
    setApiError(undefined);
    loadDataFromGraph()
      .catch((e) => {
        console.error(e);
        setApiError(e.message);
      });
  }, [props]);

  async function loadDataFromGraph(): Promise<void> {
    function tryIngestData(template: string): string {
      if (!props.dataFromDynamicSource) return template;
      return Handlebars.compile(template)(props.dataFromDynamicSource);
    }

    const path = tryIngestData(props.api ?? 'me');
    
    let graphQuery = props.graphClient.api(path);
    if (props.version) graphQuery = graphQuery.version(props.version);
    if (props.select) graphQuery = graphQuery.select(props.select);
    if (props.expand) graphQuery = graphQuery.expand(props.expand);
    if (props.filter) graphQuery = graphQuery.filter(encodeURIComponent(tryIngestData(props.filter)));

    try {
      setApiCall(`${props.version}${path}`);
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
      {apiError && <div className={styles.error}>Error in api call: <br />{apiError.body}</div>}

      {graphData.type === 'result' && <>
        <div style={{ marginBottom: '1em' }}>
          👉 <code>{JSON.stringify(graphData.value['@odata.count'])}</code> records found.
          See <code>value</code> property in connected webparts for results.
        </div>

        <PrimaryButton onClick={() => setGraphDataPopupVisible(!isGraphDataPopupVisible)} text={`${isGraphDataPopupVisible ? 'Hide result(s) from Graph' : 'Show result(s) from Graph'}`} />
        {isGraphDataPopupVisible && (
          <Popup>
            <pre dangerouslySetInnerHTML={{ __html: prettyPrintJson.toHtml(graphData.value) }} />
          </Popup>
        )}
        {props.dataFromDynamicSource && (
          <div>
            <PrimaryButton onClick={() => setDynamicDataPopupVisible(!isDynamicDataPopupVisible)} text={`${isDynamicDataPopupVisible ? 'Hide dynamic data' : 'Show dynamic data'}`} />
            {isDynamicDataPopupVisible &&
              <Popup>
                <pre dangerouslySetInnerHTML={{ __html: prettyPrintJson.toHtml(props.dataFromDynamicSource) }} />
              </Popup>
            }
          </div>
        )}
      </>}
    </div>
  );
}

export default GraphConnector;