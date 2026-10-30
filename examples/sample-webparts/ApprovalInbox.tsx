import * as React from "react";
import { Button, Card, CardHeader, FluentProvider, MessageBar, MessageBarBody, Spinner, Text } from "@fluentui/react-components";
import type { createClient } from "@spfx-fast/core";
import { useListItems } from "@spfx-fast/core/react";
import { sampleListIds, sampleSiteId } from "./sampleData.js";
import { SampleSection } from "./SampleSection.js";
import { sampleDarkTheme } from "./sampleTheme.js";

type Client = ReturnType<typeof createClient>;

/** Approval queue with an explicit write action, using a contrasting Fluent UI v9 Teams theme. */
export function ApprovalInbox({ client, siteId = sampleSiteId, listId = sampleListIds.approvals }: { client: Client; siteId?: string; listId?: string }): React.ReactElement {
  const approvals = useListItems(client, siteId, listId, {
    fields: ["Title", "Requester", "Amount", "Status"],
    filter: { field: "Status", operator: "eq", value: "Pending" },
    pageSize: 20,
    maxItems: 200,
  });
  const [busyId, setBusyId] = React.useState<string>();
  const [error, setError] = React.useState(false);

  return (
    <FluentProvider theme={sampleDarkTheme}>
      <SampleSection label="Approval inbox" title="Manager approvals" description="Review purchase requests and update their SharePoint status.">
        <Card><CardHeader header={<Text weight="semibold">{approvals.items.length} approvals waiting</Text>} />
          {error ? <MessageBar intent="error"><MessageBarBody>Approval could not be saved. Try again.</MessageBarBody></MessageBar> : null}
          {approvals.loading && approvals.items.length === 0 ? <Spinner label="Loading approvals" /> : null}
          {approvals.items.map((item) => (
            <div key={item.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16, padding: 12, borderTop: "1px solid var(--colorNeutralStroke2)" }}>
              <div><Text weight="semibold">{String(item.fields?.Title ?? "")}</Text><br /><Text size={200}>{String(item.fields?.Requester ?? "")} · {new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP", maximumFractionDigits: 0 }).format(Number(item.fields?.Amount ?? 0))}</Text></div>
              <Button appearance="primary" disabled={busyId === item.id} onClick={async () => {
                setBusyId(item.id);
                setError(false);
                try {
                  await client.updateItem(siteId, listId, item.id, { Status: "Approved" }, { eTag: item.eTag });
                  await approvals.refresh();
                } catch {
                  setError(true);
                } finally {
                  setBusyId(undefined);
                }
              }}>{busyId === item.id ? "Saving…" : "Approve"}</Button>
            </div>
          ))}
        </Card>
      </SampleSection>
    </FluentProvider>
  );
}
