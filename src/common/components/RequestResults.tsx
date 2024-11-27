import * as React from 'react';
import { MessageBar, MessageBarType, Stack } from '@fluentui/react';
import CollapsibleSection from './CollapsibleSection';
import { GraphResult, SharePointResult } from '../../webparts/graphApiConnector/models/types';

type ResultsSectionProps = {
  data: GraphResult | SharePointResult;
  dataFromDynamicSource?: never;
  labels: {
    apiRequestResults: string;
    dynamicDataResults: string;
  }
}

export default function RequestResults(props: ResultsSectionProps): JSX.Element {
  const { data, dataFromDynamicSource, labels } = props;
  return (
    <>
      <Stack tokens={{ childrenGap: 1 }} style={{ margin: '1em 0' }}>
        <MessageBar messageBarType={MessageBarType.success}>
          <div>👉 <code>{JSON.stringify((data.value)['@odata.count'])}</code> valid record(s) found.</div>
          <div>Reference <code>value</code> property in connected webparts for results</div>
        </MessageBar>
      </Stack>

      <Stack tokens={{ childrenGap: 10 }}>
        <CollapsibleSection label={labels.apiRequestResults} value={data.value} />
        {dataFromDynamicSource && <CollapsibleSection label={labels.dynamicDataResults} value={dataFromDynamicSource ?? ''} />}
      </Stack>
    </>
  );
}