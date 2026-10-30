import { Card, CardHeader, FluentProvider, Text, webLightTheme, type Theme } from "@fluentui/react-components";
import { ListView, useItems } from "@spfx-fast/core/react";
import type { createClient } from "@spfx-fast/core";

type Client = ReturnType<typeof createClient>;

/** Render inside an SPFx web part after createSPFxServices() provides the client. */
export function WorkDashboard({ client, siteId, taskListId, projectListId, theme = webLightTheme }: {
  client: Client;
  siteId: string;
  taskListId: string;
  projectListId: string;
  theme?: Theme;
}) {
  const tasks = useItems(client, siteId, taskListId, { fields: ["Title", "Status"], maxItems: 1000 });
  const projects = useItems(client, siteId, projectListId, { fields: ["Title", "Status"], maxItems: 500 });
  return (
    <FluentProvider theme={theme}>
      <main>
        <h1>Work dashboard</h1>
        <section aria-label="Summary">
          <Card><CardHeader header={<Text>Tasks</Text>} description={<Text>{tasks.items.length} loaded · {tasks.source ?? "loading"}</Text>} /></Card>
          <Card><CardHeader header={<Text>Projects</Text>} description={<Text>{projects.items.length} loaded · {projects.source ?? "loading"}</Text>} /></Card>
        </section>
        <section aria-label="Tasks"><h2>Tasks</h2><ListView items={tasks.items} fields={[{ name: "Title", title: "Task" }, { name: "Status", title: "Status" }]} loading={tasks.loading} error={tasks.error} /></section>
        <section aria-label="Projects"><h2>Projects</h2><ListView items={projects.items} fields={[{ name: "Title", title: "Project" }, { name: "Status", title: "Status" }]} loading={projects.loading} error={projects.error} /></section>
      </main>
    </FluentProvider>
  );
}
