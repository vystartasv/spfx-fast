import * as React from "react";
import { Button, Card, CardHeader, MessageBar, MessageBarBody, Spinner, Text } from "@fluentui/react-components";
import type { createClient } from "@spfx-fast/core";

type Client = ReturnType<typeof createClient>;
type Scenario = { returnBound: number; rows: number; medianMs: number; p95Ms: number };
type Measurement = { availableRows: number; scenarios: Scenario[]; storage?: StorageEstimate };
const QUERY_BOUNDS = [10, 50, 250, 1000];
const SAMPLES = 7;

function percentile(values: number[], fraction: number): number {
  const ordered = [...values].sort((left, right) => left - right);
  return ordered[Math.min(ordered.length - 1, Math.ceil(ordered.length * fraction) - 1)];
}

/** Small real-browser probe for local query latency and origin storage estimates. */
export function PerformanceProbe({ client, siteId, listId }: { client: Client; siteId: string; listId: string }): React.ReactElement {
  const [measurement, setMeasurement] = React.useState<Measurement>();
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string>();

  async function measure(): Promise<void> {
    setLoading(true);
    setError(undefined);
    try {
      const availableRows = (await client.queryItems(siteId, listId, { limit: 1000 })).length;
      const scenarios: Scenario[] = [];
      for (const returnBound of QUERY_BOUNDS) {
        const timings: number[] = [];
        let rows = 0;
        for (let index = 0; index < SAMPLES; index++) {
          const start = performance.now();
          rows = (await client.queryItems(siteId, listId, { limit: returnBound })).length;
          timings.push(performance.now() - start);
        }
        scenarios.push({ returnBound, rows, medianMs: percentile(timings, 0.5), p95Ms: percentile(timings, 0.95) });
      }
      const storage = await navigator.storage?.estimate();
      setMeasurement({ availableRows, scenarios, storage });
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setLoading(false);
    }
  }

  return (
    <Card>
      <CardHeader header={<Text>Browser cache probe</Text>} description={<Text>Seven local query samples at four return bounds; no network request is initiated.</Text>} />
      <Button disabled={loading} onClick={measure}>{loading ? "Measuring…" : "Measure local query scaling"}</Button>
      {loading ? <Spinner label="Measuring local list query" /> : null}
      {error ? <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar> : null}
      {measurement ? <dl>
        <dt>Cached rows available (capped at 1,000)</dt><dd>{measurement.availableRows}</dd>
        <table>
          <caption>Local query latency by maximum returned rows (7 runs each)</caption>
          <thead><tr><th scope="col">Return bound</th><th scope="col">Rows returned</th><th scope="col">Median</th><th scope="col">P95</th></tr></thead>
          <tbody>{measurement.scenarios.map((scenario) => <tr key={scenario.returnBound}>
            <th scope="row">{scenario.returnBound}</th>
            <td>{scenario.rows}</td>
            <td>{scenario.medianMs.toFixed(2)} ms</td>
            <td>{scenario.p95Ms.toFixed(2)} ms</td>
          </tr>)}</tbody>
        </table>
        {measurement.storage?.usage !== undefined ? <><dt>Origin storage usage</dt><dd>{formatBytes(measurement.storage.usage)}</dd></> : null}
        {measurement.storage?.quota !== undefined ? <><dt>Origin storage quota</dt><dd>{formatBytes(measurement.storage.quota)}</dd></> : null}
      </dl> : null}
      <Text size={200}>The return bound limits output rows, not how many cached records the local store scans. Storage figures cover the browser origin, not this client alone. These measurements vary by device and cache size; use tenant workloads for release decisions.</Text>
    </Card>
  );
}

function formatBytes(value: number): string {
  const units = ["B", "KB", "MB", "GB"];
  let size = value;
  let unit = 0;
  while (size >= 1024 && unit < units.length - 1) { size /= 1024; unit++; }
  return `${size.toFixed(1)} ${units[unit]}`;
}
