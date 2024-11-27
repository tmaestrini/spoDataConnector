import * as React from 'react';
import styles from './GraphConnector.module.scss';
import { IRequestResult, SharePointError, SharePointResult } from '../models/types';
import * as Handlebars from 'handlebars';
import { Icon, MessageBar, MessageBarType, Stack } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import * as strings from 'GraphConnectorWebPartStrings';
import RequestResults from '../../../common/components/RequestResults';
import { ISharePointConnectorProps } from './ISharePointConnectorProps';

const SharePointConnector: React.FunctionComponent<ISharePointConnectorProps> = (props) => {
  const [spoData, setSpoData] = React.useState<IRequestResult>({} as SharePointResult);
  const [apiError, setApiError] = React.useState<SharePointError | undefined>(undefined);
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

    const path = tryIngestDynamicData(props.api);

    const spoQueryParams = new URLSearchParams();
    if (props.select) spoQueryParams.append('$select', props.select);
    if (props.expand) spoQueryParams.append('$expand', props.expand);
    if (props.filter) spoQueryParams.append('$filter', tryIngestDynamicData(props.filter));

    const spoQuery = new URL(path);
    if (props.version && props.version === 'v2.0') spoQuery.pathname = `/v2.0${spoQuery.pathname}`;
    spoQuery.search = spoQueryParams.toString();

    try {
      setApiCall(`${spoQuery.href}`);
      const response = await props.sharePointClient.get(spoQuery.href, SPHttpClient.configurations.v1);

      if (!response.ok) throw new Error(response.statusText);
      const data = (await response.json());

      setSpoData({ value: { ...data } } as SharePointResult);
      if (props.onSharePointDataResult) props.onSharePointDataResult({ value: { ...data } } as SharePointResult);
    } catch (error) {
      setSpoData({} as SharePointResult);
      setApiError({ ...error } as SharePointError);
      if (props.onSharePointDataError) props.onSharePointDataError({ ...error } as SharePointError);
    }
  }

  return (
    <div className={styles.graphConnector}>
      <h2><Icon iconName="PlugConnected" /> SharePoint API Connection</h2>
      <div className={styles.breakContent}>SharePoint api call: {apiCall && <code>{apiCall}</code>}</div>
      {apiError &&
        <Stack tokens={{ childrenGap: 1 }} style={{ margin: '1rem 0' }}>
          <MessageBar messageBarType={MessageBarType.error}>
            <div>Error in api call: <br /><span className={styles.error}>{apiError.body}</span></div>
          </MessageBar>
        </Stack>}


      {!apiError && <>
        <RequestResults data={spoData as SharePointResult}
          dataFromDynamicSource={props.dataFromDynamicSource}
          labels={{ apiRequestResults: strings.SharePointConnector.ShowSPOResultsLabel, dynamicDataResults: strings.DataSource.ShowDynamicDataLabel }} />
      </>}
    </div>
  );
}

export default SharePointConnector;