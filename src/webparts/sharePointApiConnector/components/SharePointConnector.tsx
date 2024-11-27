import * as React from 'react';
import styles from './SharePointConnector.module.scss';
import type { ISharePointConnectorProps } from './ISharePointConnectorProps';
import { SharePointError, SharePointResult } from '../models/types';
import * as Handlebars from 'handlebars';
import { Icon } from '@fluentui/react';

const SharePointConnector: React.FunctionComponent<ISharePointConnectorProps> = (props) => {
  const [_, setSharePointData] = React.useState<SharePointResult>({} as SharePointResult);
  const [apiError, setApiError] = React.useState<SharePointError | undefined>(undefined);
  const [apiCall, setApiCall] = React.useState<string>();

  React.useEffect(() => {
    setApiError(undefined);
    loadDataFromSharePoint()
      .catch((e) => {
        console.error(e);
        setApiError(e.message);
      });
  }, [props]);

  async function loadDataFromSharePoint(): Promise<void> {
    function tryIngestDynamicData(template: string): string {
      if (!props.dataFromDynamicSource) return template;
      return Handlebars.compile(template)(props.dataFromDynamicSource);
    }

    const path = tryIngestDynamicData(props.api ?? 'me');

    let sharePointQuery = props.sharePointClient.api(path);
    if (props.version) sharePointQuery = sharePointQuery.version(props.version);
    if (props.select) sharePointQuery = sharePointQuery.select(props.select);
    if (props.expand) sharePointQuery = sharePointQuery.expand(props.expand);
    if (props.filter) sharePointQuery = sharePointQuery.filter(encodeURIComponent(tryIngestDynamicData(props.filter)));

    sharePointQuery.header('ConsistencyLevel', 'eventual');

    try {
      setApiCall(`${props.version}${path}`);
      const data = await sharePointQuery.get();

      setSharePointData({ type: 'result', value: { ...data } } as SharePointResult);
      if (props.onSharePointDataResult) props.onSharePointDataResult({ type: 'result', value: { ...data } } as SharePointResult);
    } catch (error) {
      setSharePointData({} as SharePointResult);
      setApiError({ ...error, type: 'error' } as SharePointError);
      if (props.onSharePointDataResult) props.onSharePointDataResult({ ...error, type: 'error' } as SharePointError);
    }
  }

  return (
    <div className={styles.sharePointConnector}>
      <h2><Icon iconName="PlugConnected" /> SharePoint API Connection</h2>
      <div>Graph api call: {apiCall && <code>{apiCall}</code>}</div>
      {apiError && <div className={styles.error}>Error in api call: <br />{apiError.body}</div>}
    </div>
  );
}

export default SharePointConnector;