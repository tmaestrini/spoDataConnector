import * as React from 'react';
import styles from './GraphConnector.module.scss';
import type { IGraphConnectorProps } from './IGraphConnectorProps';
import { Icon } from '@fluentui/react';

const GraphConnector: React.FunctionComponent<IGraphConnectorProps> = (props) => {
  return (
    <div className={styles.graphConnector}>
      <h2><Icon iconName="PlugConnected" /> Graph Connector</h2>
    </div>
  );
}

export default GraphConnector;