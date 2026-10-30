import * as React from "react";
import { Badge, Button, Card, CardHeader, FluentProvider, MessageBar, MessageBarBody, Spinner, Text } from "@fluentui/react-components";
import type { createClient } from "@spfx-fast/core";
import { ListView, useListItems } from "@spfx-fast/core/react";
import { sampleListIds, sampleSiteId } from "./sampleData.js";
import { SampleSection } from "./SampleSection.js";
import { sampleLightTheme } from "./sampleTheme.js";

type Client = ReturnType<typeof createClient>;

/** A bounded server-filtered work queue for a SharePoint list with an indexed Status column. */
export function TeamTaskQueue({ client, siteId = sampleSiteId, listId = sampleListIds.tasks }: { client: Client; siteId?: string; listId?: string }): React.ReactElement {
  const tasks = useListItems(client, siteId, listId, { fields: ["Title", "Owner", "Priority", "Status"], filter: { field: "Status", operator: "eq", value: "Open" }, pageSize: 50, maxItems: 500 });
  return (
    <FluentProvider theme={sampleLightTheme}>
      <SampleSection label="Team task queue" title="Team task queue" description="The next actions for keeping North Campus running.">
        {tasks.loading && tasks.items.length === 0 ? <Spinner label="Loading open tasks" /> : null}
        {tasks.error ? <MessageBar intent="error"><MessageBarBody>Open tasks could not be loaded.</MessageBarBody></MessageBar> : null}
        <Card><CardHeader header={<Text weight="semibold">{tasks.items.length} open tasks</Text>} />
          <ListView items={tasks.items} fields={[{ name: "Title", title: "Task" }, { name: "Owner", title: "Owner" }, { name: "Priority", title: "Priority" }, { name: "Status", title: "Status" }]} loading={tasks.loading} error={tasks.error} renderField={(item, field) => {
            if (field.name === "Status") return <Badge appearance="tint" color="informative">{String(item.fields?.Status ?? "Unknown")}</Badge>;
            if (field.name === "Priority") {
              const priority = String(item.fields?.Priority ?? "Normal");
              return <Badge appearance="tint" color={priority === "High" ? "danger" : priority === "Low" ? "success" : "warning"}>{priority}</Badge>;
            }
            return undefined;
          }} />
          {tasks.hasMore ? <Button onClick={() => void tasks.loadMore()}>Load more tasks</Button> : null}
        </Card>
      </SampleSection>
    </FluentProvider>
  );
}
