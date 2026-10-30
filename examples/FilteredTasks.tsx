import * as React from "react";
import { Button, FluentProvider, webLightTheme } from "@fluentui/react-components";
import { ListView, useListItems } from "@spfx-fast/core/react";
import type { createClient } from "@spfx-fast/core";

type Client = ReturnType<typeof createClient>;

/** Server-filters a large list and lets the user page through a bounded result set. */
export function FilteredTasks({ client, siteId, listId }: { client: Client; siteId: string; listId: string }): React.ReactElement {
  const tasks = useListItems(client, siteId, listId, {
    fields: ["Title", "Status"],
    filter: { field: "Status", operator: "eq", value: "Open" },
    pageSize: 50,
    maxItems: 500,
  });

  return (
    <FluentProvider theme={webLightTheme}>
    <section aria-label="Open tasks">
      <h2>Open tasks</h2>
      <ListView items={tasks.items} fields={[{ name: "Title", title: "Task" }, { name: "Status", title: "Status" }]} loading={tasks.loading} error={tasks.error} />
      {tasks.hasMore && <Button onClick={() => void tasks.loadMore()}>Load more tasks</Button>}
    </section>
    </FluentProvider>
  );
}
