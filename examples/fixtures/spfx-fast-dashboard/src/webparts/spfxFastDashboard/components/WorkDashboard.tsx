import * as React from "react";
import { Card, CardHeader, FluentProvider, Tab, TabList, Text, webLightTheme, type Theme } from "@fluentui/react-components";
import { ListView, useItems } from "@spfx-fast/core/react";
import type { createClient } from "@spfx-fast/core";
import { DriveDeltaPanel } from "./DriveDeltaPanel";
import { DocumentLibraryPanel } from "./DocumentLibraryPanel";
import { DashboardStartupProbe } from "./DashboardStartupProbe";
import { PerformanceProbe } from "./PerformanceProbe";
import { TenantValidationPanel } from "./TenantValidationPanel";
import { CalendarAgenda } from "./CalendarAgenda";
import { PrivateFloorPlanProbe } from "./PrivateFloorPlanProbe";
import type { ISpfxFastDashboardProps } from "./ISpfxFastDashboardProps";

type Client = ReturnType<typeof createClient>;

/** Render inside an SPFx web part after its authenticated Graph client is created. */
function WorkLists({ client, siteId, taskListId, projectListId }: { client: Client; siteId: string; taskListId: string; projectListId: string }): React.ReactElement {
  const tasks = useItems(client, siteId, taskListId, { fields: ["Title", "Status"], maxItems: 1000 });
  const projects = useItems(client, siteId, projectListId, { fields: ["Title", "Status"], maxItems: 500 });
  const startedAt = React.useRef(performance.now()).current;
  return <>
    <DashboardStartupProbe startedAt={startedAt} tasks={tasks} projects={projects} />
    <section aria-label="Summary">
      <Card><CardHeader header={<Text>Tasks</Text>} description={<Text>{tasks.loading ? "Loading…" : `${tasks.items.length} loaded`}</Text>} /></Card>
      <Card><CardHeader header={<Text>Projects</Text>} description={<Text>{projects.loading ? "Loading…" : `${projects.items.length} loaded`}</Text>} /></Card>
    </section>
    <section aria-label="Tasks"><h2>Tasks</h2><ListView items={tasks.items} fields={[{ name: "Title", title: "Task" }, { name: "Status", title: "Status" }]} loading={tasks.loading} error={tasks.error} /></section>
    <section aria-label="Projects"><h2>Projects</h2><ListView items={projects.items} fields={[{ name: "Title", title: "Project" }, { name: "Status", title: "Status" }]} loading={projects.loading} error={projects.error} /></section>
  </>;
}

export function WorkDashboard({ services, client, siteId, webTitle, taskListId, projectListId, driveId, folderItemId = "root", theme = webLightTheme }: ISpfxFastDashboardProps & {
  theme?: Theme;
}): React.ReactElement {
  const [view, setView] = React.useState<"work" | "calendar" | "floorplan">("work");
  const viewId = React.useId();
  return (
    <FluentProvider theme={theme}>
      <main>
        <h1>Work dashboard</h1>
        <Text>Connected to {webTitle} through the signed-in SharePoint session.</Text>
        <TabList aria-label="Dashboard views" selectedValue={view} onTabSelect={(_, data) => setView(data.value as "work" | "calendar" | "floorplan")}>
          <Tab id={`${viewId}-work`} value="work">Work</Tab>
          <Tab id={`${viewId}-calendar`} value="calendar">Calendar</Tab>
          {driveId ? <Tab id={`${viewId}-floorplan`} value="floorplan">Floor plan</Tab> : null}
        </TabList>
        <div role="tabpanel" aria-labelledby={`${viewId}-${view}`}>
        {view === "calendar" ? <CalendarAgenda client={client} /> : view === "floorplan" && driveId ? <PrivateFloorPlanProbe client={client} driveId={driveId} /> : <section aria-label="Work dashboard">
          {taskListId && projectListId ? <><WorkLists client={client} siteId={siteId} taskListId={taskListId} projectListId={projectListId} /><PerformanceProbe client={client} siteId={siteId} listId={taskListId} /></> : <Text>Configure both list IDs to display the list dashboard.</Text>}
          {driveId ? <section aria-label="Documents"><h2>Documents</h2><DriveDeltaPanel client={client} driveId={driveId} /><DocumentLibraryPanel client={client} driveId={driveId} folderItemId={folderItemId} /></section> : <Text>The default document library is unavailable to this account.</Text>}
          {taskListId ? <TenantValidationPanel services={services} siteId={siteId} listId={taskListId} driveId={driveId} folderItemId={folderItemId} /> : null}
        </section>}
        </div>
      </main>
    </FluentProvider>
  );
}
