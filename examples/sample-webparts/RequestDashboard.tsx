import * as React from "react";
import { Badge, Button, Card, CardHeader, FluentProvider, makeStyles, MessageBar, MessageBarBody, Spinner, Text, tokens } from "@fluentui/react-components";
import type { createClient } from "@spfx-fast/core";
import { ItemForm, ListView, useItems, type FieldDefinition } from "@spfx-fast/core/react";
import { ItemAttachmentsPanel } from "../ItemAttachmentsPanel.js";
import { sampleListIds, sampleSiteId } from "./sampleData.js";
import { SampleSection } from "./SampleSection.js";
import { sampleLightTheme } from "./sampleTheme.js";

type Client = ReturnType<typeof createClient>;
type SharePointApi = React.ComponentProps<typeof ItemAttachmentsPanel>["api"];
const fields: FieldDefinition[] = [
  { name: "Title", title: "Request", type: "text", required: true },
  { name: "Requester", title: "Requester", type: "text", required: true },
  { name: "Priority", title: "Priority", type: "choice", choices: ["High", "Normal"] },
  { name: "Amount", title: "Estimated amount", type: "number" },
  { name: "Status", title: "Status" },
];
const useStyles = makeStyles({
  panels: {
    display: "grid",
    gridTemplateColumns: "minmax(280px, 0.8fr) minmax(0, 1.2fr)",
    gap: tokens.spacingHorizontalL,
    alignItems: "start",
    "@media (max-width: 760px)": { gridTemplateColumns: "1fr" },
  },
});

/** Request intake, status counts, and a cache-first list that uses the authenticated host client. */
export function RequestDashboard({ client, siteId = sampleSiteId, listId = sampleListIds.requests, sharepoint, attachmentItemId }: {
  client: Client;
  siteId?: string;
  listId?: string;
  sharepoint?: SharePointApi;
  attachmentItemId?: number;
}): React.ReactElement {
  const requests = useItems(client, siteId, listId, { fields: fields.map((field) => field.name), maxItems: 1000 });
  const [saved, setSaved] = React.useState(false);
  const [activeAttachmentItemId, setActiveAttachmentItemId] = React.useState(attachmentItemId);
  const [attachmentWarning, setAttachmentWarning] = React.useState<string>();
  React.useEffect(() => { setActiveAttachmentItemId(attachmentItemId); }, [attachmentItemId]);
  const pending = requests.items.filter((item) => item.fields?.Status === "Pending" || item.fields?.Status === "In review").length;
  const styles = useStyles();

  return (
    <FluentProvider theme={sampleLightTheme}>
      <SampleSection label="Request dashboard" title="Request dashboard" description="Raise a purchase or facilities request and track it through review.">
        <Card><CardHeader header={<Text weight="semibold">{pending} requests need review</Text>} description={<Text>{requests.items.length} sample requests</Text>} /></Card>
        {saved ? <MessageBar intent="success"><MessageBarBody>Request saved.</MessageBarBody></MessageBar> : null}
        {attachmentWarning ? <MessageBar intent="warning"><MessageBarBody>{attachmentWarning}</MessageBarBody></MessageBar> : null}
        {requests.error ? <MessageBar intent="error"><MessageBarBody>Requests could not be loaded.</MessageBarBody></MessageBar> : null}
        {requests.loading && requests.items.length === 0 ? <Spinner label="Loading requests" /> : null}
        <div className={styles.panels}>
          <Card><CardHeader header={<Text weight="semibold">New request</Text>} description={<Text>Include an estimated cost for faster approval.</Text>} />
            <ItemForm fields={fields.slice(0, 4)} submitLabel="Submit request" attachments={sharepoint && activeAttachmentItemId ? <ItemAttachmentsPanel api={sharepoint} listId={listId} itemId={activeAttachmentItemId} /> : null} onSubmit={async (values) => {
            const created = await client.createItem(siteId, listId, values);
            setSaved(true);
            setAttachmentWarning(undefined);
            if (sharepoint) {
              const numericItemId = Number(created.id);
              if (Number.isSafeInteger(numericItemId) && numericItemId > 0) setActiveAttachmentItemId(numericItemId);
              else setAttachmentWarning("Request saved. Pass its numeric SharePoint item ID to enable attachments.");
            }
            await requests.refresh();
          }} />
          </Card>
          <Card><CardHeader header={<Text weight="semibold">Recent requests</Text>} description={<Text>Current items across North Campus.</Text>} />
            <ListView items={requests.items} fields={fields.filter((field) => field.name !== "Requester")} loading={requests.loading} error={requests.error} renderField={(item, field) => {
              const value = String(item.fields?.[field.name] ?? "");
              if (field.name === "Priority") return <Badge appearance="tint" color={value === "High" ? "danger" : "informative"}>{value}</Badge>;
              if (field.name === "Status") return <Badge appearance="tint" color={value === "Approved" ? "success" : value === "Pending" || value === "In review" ? "warning" : "informative"}>{value}</Badge>;
              if (field.name === "Amount") return new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP", maximumFractionDigits: 0 }).format(Number(item.fields?.Amount ?? 0));
              return undefined;
            }} />
            <Button appearance="secondary" onClick={() => void requests.refresh()}>Refresh requests</Button>
          </Card>
        </div>
      </SampleSection>
    </FluentProvider>
  );
}
