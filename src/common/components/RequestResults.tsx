import * as React from 'react';
import { Stack } from '@fluentui/react';
import CollapsibleSection from './CollapsibleSection';
import { GraphResult } from '../../webparts/graphApiConnector/models/types';

type ResultsSectionProps = {
  data: GraphResult;
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
      <div style={{ marginBottom: '1em' }}>
        👉 <code>{JSON.stringify((data.value)['@odata.count'])}</code> valid record(s) found.
        Refrence <code>value</code> property in connected webparts for results.
      </div>

      <Stack tokens={{ childrenGap: 10 }}>
        <CollapsibleSection label={labels.apiRequestResults} value={data.value} />
        {dataFromDynamicSource && <CollapsibleSection label={labels.dynamicDataResults} value={dataFromDynamicSource ?? ''} />}
      </Stack>
    </>
  );
}