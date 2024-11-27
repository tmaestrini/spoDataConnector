import { prettyPrintJson } from "pretty-print-json";
import * as React from "react";
import styles from "../../webparts/graphApiConnector/components/GraphConnector.module.scss";

type CollapsibleSectionProps = {
  label: string;
  value: string
}

export default function CollapsibleSection(props: CollapsibleSectionProps): JSX.Element {
  const { label, value } = props;
  return (
    <details>
      <summary>
        {label}
      </summary>
      <div className={styles.breakContent}>
        <pre dangerouslySetInnerHTML={{ __html: prettyPrintJson.toHtml(value) }} />
      </div>
    </details>
  );
}
