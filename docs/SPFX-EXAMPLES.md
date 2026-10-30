# Authenticated SPFx examples

Examples use authenticated `SPHttpClient` and `MSGraphClientV3`. Tokens are not acquired or stored by the app.

## Create one authenticated data client

Create services once and share them. Cache scope uses tenant and user IDs; it does not grant access.

```ts
import { SPHttpClient } from "@microsoft/sp-http";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import { createSPFxServices } from "@spfx-fast/core/spfx";

export async function createDashboardServices(context: WebPartContext) {
  return createSPFxServices(context, SPHttpClient.configurations.v1);
}
```

Typed list reads, delta sync, and CRUD use Graph. `sharepoint` covers CAML, `RenderListDataAsStream`, change tokens, and audience targeting through the authenticated host.

## Fluent UI v9 single-page dashboard

Share one client and IndexedDB store. `useItems` subscribes to local changes and refreshes each list by delta.

```tsx
import { Card, CardHeader, Text } from "@fluentui/react-components";
import { ListView, useItems, type FieldDefinition } from "@spfx-fast/core/react";
import type { createClient } from "@spfx-fast/core";

type Client = ReturnType<typeof createClient>;

export function WorkDashboard({
  client,
  siteId,
  taskListId,
  projectListId,
}: {
  client: Client;
  siteId: string;
  taskListId: string;
  projectListId: string;
}) {
  const tasks = useItems(client, siteId, taskListId, {
    fields: ["Title", "Status", "DueDate"],
  });
  const projects = useItems(client, siteId, projectListId, {
    fields: ["Title", "Status"],
  });
  const openTasks = tasks.items.filter((item) => item.fields?.Status !== "Complete").length;
  const taskFields: FieldDefinition[] = [
    { name: "Title", title: "Task" },
    { name: "Status", title: "Status" },
    { name: "DueDate", title: "Due" },
  ];

  return (
    <main>
      <h1>Work dashboard</h1>
      <section aria-label="Work summary">
        <Card><CardHeader header={<Text>Open tasks</Text>} description={<Text>{openTasks}</Text>} /></Card>
        <Card><CardHeader header={<Text>Projects</Text>} description={<Text>{projects.items.length}</Text>} /></Card>
      </section>
      <section aria-label="Tasks">
        <ListView items={tasks.items} fields={taskFields} loading={tasks.loading} error={tasks.error} />
      </section>
    </main>
  );
}
```

Synced lists load locally. For large lists, use Graph filters and paging.

## Validate authenticated tenant access

Run the [SPFx 1.24 fixture](../examples/fixtures/spfx-fast-dashboard/README.md) and select **Run signed-in checks**. It probes REST, Graph, cache, storage, and drive access. Separate confirmed probes test CRUD, Person/Lookup, drive delta, and file upload/download. Results apply only to the tested user and resources; they are not a permission audit or benchmark.

## Vanilla DOM page inside an SPFx component

Render HTML without React inside a web part. A standalone `.html` page has no SPFx context; the host must supply the client.

```ts
import type { createClient } from "@spfx-fast/core";

type Client = ReturnType<typeof createClient>;

export async function renderTaskList(
  root: HTMLElement,
  client: Client,
  siteId: string,
  listId: string,
) {
  root.replaceChildren();
  const heading = document.createElement("h2");
  heading.textContent = "Tasks";
  const list = document.createElement("ul");
  root.append(heading, list);

  for await (const item of client.listItems(siteId, listId, {
    fields: ["Title", "Status"],
    maxItems: 100,
  })) {
    const row = document.createElement("li");
    row.textContent = `${String(item.fields?.Title ?? "Untitled")} — ${String(item.fields?.Status ?? "")}`;
    list.append(row);
  }
}
```

Call after `createDashboardServices(context)`. Use `textContent` to keep SharePoint values inert.

## SharePoint Markdown files and interactive Markdown

SharePoint Online and OneDrive store and display `.md` files. Apps can read them as authenticated content. Native Markdown does not run scripts or SPFx components.

Use `useMarkdownFile` or `readDriveItemContent(..., { as: "text" })` for cached-first text. Render with `MarkdownContent` from `@spfx-fast/core/markdown` (GFM; raw HTML stays inert). Keep interactivity in compiled React. Never evaluate SharePoint Markdown as MDX:

```mdx
import { WorkDashboard } from "./WorkDashboard";

# Team overview

Updated from the signed-in SharePoint session.

<WorkDashboard client={client} siteId={siteId} taskListId={taskListId} projectListId={projectListId} />
```

This is compiled React/MDX, not runtime `.md`. The SPFx host supplies values and renders it. Use Markdown for documents and web parts for interactive data.

## Graph permissions

Sign-in does not grant every Graph scope. Request only required scopes in `config/package-solution.json` and get admin consent. Check the endpoint's current permission table; Graph list reads document delegated `Sites.Read.All`, while writes require write access. [Graph list items](https://learn.microsoft.com/en-us/graph/api/listitem-list?view=graph-rest-1.0), [SPFx API permission requests](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient).

The generator emits a Fluent UI v9 app and SPFx host wrapper with authenticated services and connection/loading states. Adapt the generated React code or use authenticated raw REST/Graph for uncovered browser APIs. See the [API map](API-COVERAGE.md) for limits.
