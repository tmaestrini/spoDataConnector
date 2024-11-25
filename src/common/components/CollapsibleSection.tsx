import { prettyPrintJson } from "pretty-print-json";
import * as React from "react";

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
      <pre dangerouslySetInnerHTML={{ __html: prettyPrintJson.toHtml(value) }} />
    </details>);
}
