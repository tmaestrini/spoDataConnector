import * as React from 'react';
import { MessageBar, MessageBarType, Stack } from '@fluentui/react';
import CollapsibleSection from './CollapsibleSection';
import { GraphResult, SharePointResult } from '../../webparts/apiConnector/models/types';

type ResultsSectionProps = {
  data: GraphResult | SharePointResult;
  dataFromDynamicSource?: never;
  labels: {
    referencePropertyInfo?: string;
    apiRequestResults: string;
    dynamicDataResults: string;
  }
}

export default function RequestResults(props: ResultsSectionProps): JSX.Element {
  const { data, dataFromDynamicSource, labels } = props;
  return (
    <>
      <Stack tokens={{ childrenGap: 1 }} style={{ margin: '1rem 0' }}>
        <MessageBar messageBarType={MessageBarType.success}>
          <div>{data.result?.['@odata.count'] && <code>{JSON.stringify((data.result)['@odata.count'])}</code>} valid record(s) found.</div>
          {labels.referencePropertyInfo && <div dangerouslySetInnerHTML={{ __html: labels.referencePropertyInfo }} />}
        </MessageBar>
      </Stack>

      <Stack tokens={{ childrenGap: 10 }}>
        <CollapsibleSection label={labels.apiRequestResults} value={data.result} />
        {dataFromDynamicSource && <CollapsibleSection label={labels.dynamicDataResults} value={dataFromDynamicSource ?? ''} />}
      </Stack>
    </>
  );
}