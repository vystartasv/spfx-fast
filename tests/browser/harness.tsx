import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider } from "@fluentui/react-components";
import { ItemForm, ListView, useDriveItemContent, useListItems, type FieldDefinition } from "../../src/react.js";
import { defineSPFxFastListElement } from "../../src/elements.js";
import { createClient, createMemoryStore, createResilientLocalStore, DriveUploadInterruptedError } from "../../src/index.js";
import { DashboardStartupProbe, type DashboardReadState } from "../../examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/components/DashboardStartupProbe";
import { SampleWebpartsGallery } from "../../examples/sample-webparts/SampleWebpartsGallery.js";
import { sampleLightTheme } from "../../examples/sample-webparts/sampleTheme.js";
import { RequestDashboard } from "../../examples/sample-webparts/RequestDashboard.js";
import { MarkdownContent } from "../../src/markdown.js";
import { DriveDeltaPanel } from "../../examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/components/DriveDeltaPanel.js";
import { DocumentLibraryPanel } from "../../examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/components/DocumentLibraryPanel.js";
import { ItemAttachmentsPanel } from "../../examples/ItemAttachmentsPanel.js";
import { CalendarAgenda } from "../../examples/CalendarAgenda.js";
import { WorkDashboard } from "../../examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/components/WorkDashboard.js";
import { TenantValidationPanel } from "../../examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/components/TenantValidationPanel.js";
import { DeviceCapture } from "../../examples/DeviceCapture.js";
import { FloorPlanMap } from "../../examples/FloorPlanMap.js";
import { DriveFloorPlanMap } from "../../examples/DriveFloorPlanMap.js";
import { createSampleClient } from "../../examples/sample-webparts/createSampleClient.js";
import { sampleListIds, sampleSiteId } from "../../examples/sample-webparts/sampleData.js";
import { PrivateFloorPlanProbe } from "../../examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/components/PrivateFloorPlanProbe.js";

const fields: FieldDefinition[] = [
  { name: "Title", title: "Title", type: "text", required: true },
  { name: "Priority", title: "Priority", type: "choice", choices: ["High", "Normal"], required: true },
];
const items = Array.from({ length: 1000 }, (_, index) => ({
  id: `item-${index}`,
  fields: { Title: `Task ${index}`, Owner: `Owner ${index % 5}`, Status: index % 2 === 0 ? "Open" : "Closed", Priority: index % 3 === 0 ? "High" : "Normal", Modified: `2026-10-${String((index % 28) + 1).padStart(2, "0")}` },
}));
const root = createRoot(document.getElementById("app")!);
type Client = ReturnType<typeof createClient>;

export function mountForm(): void {
  root.render(
    <FluentProvider theme={sampleLightTheme}>
      <ItemForm fields={fields} onSubmit={() => { window.__spfxFastSubmitCount++; }} />
    </FluentProvider>,
  );
}

export function mountPersonLookupForm(): void {
  root.render(<FluentProvider theme={sampleLightTheme}><ItemForm
    fields={[
      { name: "Owner", title: "Owner", type: "person", required: true, options: [{ value: 13, label: "Synthetic owner" }] },
      { name: "Project", title: "Project", type: "lookup", required: true, options: [{ value: "42", label: "Synthetic project" }] },
    ]}
    onSubmit={async (values) => { window.__spfxFastReferenceSubmit = values; }}
  /></FluentProvider>);
}

export function mountSampleWebparts(): void {
  root.render(<SampleWebpartsGallery />);
}

export function mountRequestDashboardAttachments(): void {
  const sample = createSampleClient();
  const client = {
    ...sample,
    async createItem(siteId: string, listId: string, fields: Record<string, unknown>) {
      return { ...await sample.createItem(siteId, listId, fields), id: "42" };
    },
  } as Client;
  const attachments = [{ FileName: "request-quote.txt", Length: "18", ETag: '"1"' }];
  const sharepoint = {
    async listItemAttachments() { return { value: attachments.map((item) => ({ ...item })) }; },
    async readItemAttachment() { window.__spfxAttachmentDownloads++; return new Blob(["sample attachment"]); },
    async addItemAttachment(_listId: string, _itemId: number, fileName: string) { window.__spfxAttachmentMutations++; attachments.push({ FileName: fileName, Length: "10", ETag: '"1"' }); },
    async updateItemAttachment() { window.__spfxAttachmentMutations++; },
    async deleteItemAttachment() { window.__spfxAttachmentMutations++; },
  } as unknown as React.ComponentProps<typeof ItemAttachmentsPanel>["api"];
  root.render(<RequestDashboard client={client} sharepoint={sharepoint} />);
}

export function mountMarkdownSecurity(): void {
  root.render(<FluentProvider theme={sampleLightTheme}><section aria-label="Markdown safety test"><MarkdownContent>{`# Safe heading\n\n**Bold text** and [unsafe](javascript:alert(1))\n\n<img src=x onerror=alert(1)>\n\n<script>window.__spfxFastXss = true</script>\n\n| Item | State |\n| --- | --- |\n| Example | Ready |`}</MarkdownContent></section></FluentProvider>);
}

export async function probeSampleWrite(): Promise<{ id?: string; error?: string }> {
  const client = createSampleClient();
  await client.syncList(sampleSiteId, sampleListIds.requests, { fields: ["Title"] });
  try {
    const item = await client.createItem(sampleSiteId, sampleListIds.requests, { Title: "Transport probe" });
    return { id: item.id };
  } catch (error) {
    return { error: error instanceof Error ? error.message : String(error) };
  }
}

export function mountList(): void {
  root.render(
    <FluentProvider theme={sampleLightTheme}>
      <ListView items={items} fields={[
        { name: "Title", title: "Task" },
        { name: "Owner", title: "Owner" },
        { name: "Status", title: "Status" },
        { name: "Priority", title: "Priority" },
        { name: "Modified", title: "Modified" },
      ]} pageSize={10} />
    </FluentProvider>,
  );
}

export function mountCustomElement(): void {
  const tagName = "spfx-fast-browser-list";
  defineSPFxFastListElement(tagName);
  const element = document.createElement(tagName) as HTMLElement & { client?: Client };
  element.setAttribute("title", "Agent-built list");
  element.setAttribute("site-id", "tenant,site,web");
  element.setAttribute("list-id", "tasks-id");
  element.setAttribute("max-items", "1");
  element.setAttribute("fields", "Title,Status");
  document.getElementById("app")!.replaceChildren(element);

  // Assign after connection to exercise the HTML-host integration timing.
  element.client = {
    async *listItems(_siteId: string, _listId: string, options: { maxItems: number }) {
      window.__spfxFastListCalls++;
      const values = [
        { id: "safe", fields: { Title: '<p><strong>Rich text title</strong> <img src=x onerror=window.__spfxFastXss=true> <a href="javascript:window.__spfxFastXss=true">unsafe link</a><script>window.__spfxFastXss=true</script></p>', Status: "Open" } },
        { id: "bounded-away", fields: { Title: "Second row", Status: "Closed" } },
      ];
      for (const item of values.slice(0, options.maxItems)) yield item;
    },
  } as unknown as Client;
}

export function mountCachedRefresh(): void {
  const scope = "browser:cache-refresh";
  const store = createMemoryStore();
  const url = new URL("/v1.0/sites/site-id/lists/list-id/items", "https://graph.microsoft.com");
  url.searchParams.set("$top", "2");
  url.searchParams.set("$select", "id,eTag,fields");
  url.searchParams.set("$expand", "fields($select=Title)");
  void (async () => {
    await store.set(JSON.stringify([scope, url.toString()]), { value: [{ id: "cached", fields: { Title: "Cached record" } }] });
    const client = createClient({
      scope,
      store,
      get: async () => {
        await new Promise((resolve) => window.setTimeout(resolve, 350));
        window.__spfxFastRefreshCalls++;
        return { value: [{ id: "fresh", fields: { Title: "Fresh record" } }] };
      },
    });
    root.render(<FluentProvider theme={sampleLightTheme}><CachedRefreshSurface client={client} /></FluentProvider>);
  })();
}

function CachedRefreshSurface({ client }: { client: Client }): React.ReactElement {
  const result = useListItems(client, "site-id", "list-id", { fields: ["Title"], pageSize: 2 });
  return <section aria-label="Cached refresh sample">
    <p role="status">{result.source ?? "loading"}</p>
    <ul>{result.items.map((item) => <li key={item.id}>{String(item.fields?.Title ?? "")}</li>)}</ul>
    {result.error ? <p role="alert">{String(result.error instanceof Error ? result.error.message : result.error)}</p> : null}
  </section>;
}

export function mountCacheFailover(): void {
  const primary = createMemoryStore();
  primary.get = async () => { throw new Error("simulated primary cache failure"); };
  const store = createResilientLocalStore(primary, createMemoryStore(), () => { window.__spfxFastStorageErrors++; });
  const client = createClient({
    scope: "browser:cache-failover",
    store,
    get: async () => ({ value: [{ id: "network", fields: { Title: "Available from SharePoint" } }] }),
  });
  const tagName = "spfx-fast-failover-list";
  defineSPFxFastListElement(tagName);
  const element = document.createElement(tagName) as HTMLElement & { client?: Client };
  element.setAttribute("site-id", "site-id");
  element.setAttribute("list-id", "list-id");
  element.setAttribute("fields", "Title");
  element.client = client;
  document.getElementById("app")!.replaceChildren(element);
}

export function mountQuotaDenial(): void {
  const primary = createMemoryStore();
  primary.set = async () => { throw new DOMException("Simulated browser quota exhaustion", "QuotaExceededError"); };
  let storageErrorName = "";
  const store = createResilientLocalStore(primary, createMemoryStore(), (error) => {
    storageErrorName = error instanceof DOMException ? error.name : "unknown";
  });
  const client = createClient({
    scope: "browser:quota-denial",
    store,
    get: async () => ({ value: [{ id: "cached-after-quota", fields: { Title: "Available from memory after quota denial" } }] }),
  });
  void (async () => {
    const url = new URL("/v1.0/sites/site/lists/list/items", "https://graph.microsoft.com").toString();
    const first = await client.readPage(url);
    await first.refresh;
    await new Promise((resolve) => window.setTimeout(resolve, 0));
    const second = await client.readPage(url);
    await second.refresh;
    window.__spfxFastQuotaResult = { firstSource: first.source, secondSource: second.source, storageErrorName };
    root.render(<p role="status">Quota fallback source: {second.source}</p>);
  })();
}

export function mountStartupProbe(): void {
  const startedAt = performance.now();
  let tasks: DashboardReadState = { loading: true };
  let projects: DashboardReadState = { loading: true };
  const render = () => root.render(<FluentProvider theme={sampleLightTheme}>
    <DashboardStartupProbe startedAt={startedAt} tasks={tasks} projects={projects} />
  </FluentProvider>);
  render();
  window.setTimeout(() => {
    tasks = { loading: true, source: "cache" };
    render();
  }, 40);
  window.setTimeout(() => {
    tasks = { loading: false, source: "network" };
    projects = { loading: false, source: "network" };
    render();
  }, 100);
}

export function mountDriveDelta(): void {
  const scope = "browser:drive-delta";
  const store = createMemoryStore();
  void (async () => {
    const cursor = "https://graph.microsoft.com/v1.0/drives/drive-id/root/delta?token=old";
    await store.putItems(scope, "__spfx_fast_drive_items__", "drive-id", [
      { id: "cached-file", name: "Cached file.md", file: {}, lastModifiedDateTime: "2026-10-20T10:00:00Z" },
      { id: "new-sheet", name: "Old Costs.xlsx", file: {}, lastModifiedDateTime: "2026-10-19T10:00:00Z" },
      { id: "cached-folder", name: "Archive", folder: { childCount: 3 } },
    ]);
    await store.setCursor(scope, "__spfx_fast_drive_items__", "drive-id", cursor);
    const client = createClient({
      scope,
      store,
      get: async () => {
        await new Promise((resolve) => window.setTimeout(resolve, 1200));
        return {
          value: [
            { id: "cached-file", name: "Fresh file.md", file: {}, size: 120, webUrl: "https://contoso.sharepoint.com/sites/demo/Shared%20Documents/Fresh-file.md", lastModifiedDateTime: "2026-10-20T10:00:00Z" },
            { id: "new-file", name: "New file.txt", file: {}, size: 80, webUrl: "https://contoso.sharepoint.com/sites/demo/Shared%20Documents/New-file.txt", lastModifiedDateTime: "2026-10-29T10:00:00Z" },
            { id: "new-sheet", name: "Costs.xlsx", file: {}, size: 1024, lastModifiedDateTime: "2026-10-21T10:00:00Z" },
            { id: "new-image", name: "Floorplan.png", file: {}, size: 2000, lastModifiedDateTime: "2026-10-22T10:00:00Z" },
            { id: "new-folder", name: "Archive", folder: { childCount: 3 } },
          ],
          "@odata.deltaLink": "https://graph.microsoft.com/v1.0/drives/drive-id/root/delta?token=new",
        };
      },
      request: async () => { window.__spfxDriveDownloads++; return new Blob(["sample file"]); },
    });
    root.render(<FluentProvider theme={sampleLightTheme}><DriveDeltaPanel client={client} driveId="drive-id" /></FluentProvider>);
  })();
}

export function mountDocumentLibraryPanel(): void {
  const markdown = `# Operations guide\n\nCached file content from the signed-in library.\n\n| Topic | Owner |\n| --- | --- |\n| Onboarding | People team |`;
  const items = [
    { id: "guide", name: "Operations-guide.md", size: 240, webUrl: "https://contoso.sharepoint.com/sites/demo/Shared%20Documents/Operations-guide.md", file: {} },
    { id: "plan", name: "Quarterly-plan.docx", size: 82000, webUrl: "https://contoso.sharepoint.com/sites/demo/Shared%20Documents/Quarterly-plan.docx", file: {} },
    { id: "budget", name: "Team-budget.xlsx", size: 14000, file: {} },
    { id: "banner", name: "Office-banner.png", size: 48000, file: {} },
    { id: "folder", name: "Archive", folder: { childCount: 3 } },
  ];
  let uploadAttempts = 0;
  const mockClient = {
    async getDriveItemPage() { return { items, source: "network" as const, refresh: Promise.resolve({ value: items, "@odata.nextLink": undefined }) }; },
    async readDriveItemContent() { return { data: markdown, source: "cache" as const, refresh: Promise.resolve(markdown) }; },
    async downloadDriveItem() { window.__spfxDriveDownloads++; return new Blob(["sample file"]); },
    async uploadDriveItem() {
      if (uploadAttempts++ === 0) throw new DriveUploadInterruptedError(async () => ({ id: "uploaded" }));
      return { id: "uploaded" };
    },
  } as unknown as Client;
  root.render(<FluentProvider theme={sampleLightTheme}><DocumentLibraryPanel client={mockClient} driveId="demo-drive" folderItemId="root" /></FluentProvider>);
}

export function mountItemAttachmentsPanel(): void {
  const attachments = [{ FileName: "meeting-notes.txt", Length: "18", ETag: '"1"' }];
  const api = {
    async listItemAttachments() { return { value: attachments.map((attachment) => ({ ...attachment })) }; },
    async readItemAttachment() { window.__spfxAttachmentDownloads++; return new Blob(["sample attachment"]); },
    async addItemAttachment(_listId: string, _itemId: number, fileName: string) { window.__spfxAttachmentMutations++; attachments.push({ FileName: fileName, Length: "10", ETag: '"1"' }); },
    async updateItemAttachment(_listId: string, _itemId: number, fileName: string) { window.__spfxAttachmentMutations++; const item = attachments.find((attachment) => attachment.FileName === fileName); if (item) item.ETag = '"2"'; },
    async deleteItemAttachment(_listId: string, _itemId: number, fileName: string) { window.__spfxAttachmentMutations++; const index = attachments.findIndex((attachment) => attachment.FileName === fileName); if (index >= 0) attachments.splice(index, 1); },
  } as unknown as React.ComponentProps<typeof ItemAttachmentsPanel>["api"];
  root.render(<FluentProvider theme={sampleLightTheme}><ItemAttachmentsPanel api={api} listId="sample-list-guid" itemId={42} /></FluentProvider>);
}

export function mountCalendarAgenda(): void {
  const client = {
    async getCalendarView(_start: string, _end: string, options?: { cursor?: string }) {
      return options?.cursor
        ? { items: [{ id: "event-2", subject: "Next review", start: { dateTime: "2026-10-30T15:00:00", timeZone: "UTC" }, end: { dateTime: "2026-10-30T15:30:00", timeZone: "UTC" } }], complete: true }
        : { items: [{ id: "event-1", subject: "Planning meeting", start: { dateTime: "2026-10-29T09:00:00", timeZone: "UTC" }, end: { dateTime: "2026-10-29T10:00:00", timeZone: "UTC" }, location: { displayName: "Room 4" } }], nextCursor: "next", complete: false };
    },
  } as unknown as Client;
  root.render(<FluentProvider theme={sampleLightTheme}><CalendarAgenda client={client} /></FluentProvider>);
}

export function mountTenantValidationPanel(): void {
  let title = "SPFx Fast CRUD test";
  let listPageCalls = 0;
  let uploaded = new Blob();
  const client = {
    async getLists() { return [{ id: "list" }]; },
    async batch() { return ["site", "list", "columns", "contentTypes", "items"].map((id) => ({ id, status: 200, body: {} })); },
    async getCalendarView() { return { items: [], complete: true }; },
    async inspectList() { return { list: { id: "list" }, columns: [{ name: "Title" }] }; },
    async getListContentTypes() { return [{ id: "item" }]; },
    async getDriveItemPage() { return { items: [], source: "network", refresh: Promise.resolve({ value: [] }) }; },
    async getItemPage() {
      listPageCalls++;
      return { items: [], source: listPageCalls === 1 ? "network" : "cache", refresh: Promise.resolve({ value: [] }) };
    },
    async syncDrive(_driveId: string, options: { maxPages?: number }) {
      window.__spfxFastDeltaMaxPages = options.maxPages ?? 0;
      return { pages: 1, changed: 2, complete: false, nextCursor: "next-page" };
    },
    async createItem(_siteId: string, _listId: string, fields: Record<string, unknown>) { title = String(fields.Title); return { id: "temporary-item", eTag: '"1"' }; },
    async updateItem(_siteId: string, _listId: string, _itemId: string, fields: Record<string, unknown>) { title = String(fields.Title); },
    async deleteItem() { window.__spfxFastTenantDeletes++; },
    async uploadDriveItem(_driveId: string, _folderId: string, _name: string, content: Blob) { uploaded = content; return { id: "temporary-file" }; },
    async downloadDriveItem() { return uploaded; },
    async requestGraph(url: string, options?: { method?: string }) {
      if (options?.method === "DELETE") { window.__spfxFastTenantDeletes++; return undefined; }
      return { fields: { Title: title } };
    },
  };
  const services = {
    client,
    rest: { async get() { return { Id: "web-id" }; } },
    sharepoint: {
      async getCurrentUser() { return { Id: 1 }; },
      async getListViewsPage() { return { items: [{ Id: "view" }] }; },
    },
  } as unknown as React.ComponentProps<typeof TenantValidationPanel>["services"];
  root.render(<FluentProvider theme={sampleLightTheme}><TenantValidationPanel services={services} siteId="site-id" listId="list-id" driveId="drive-id" /></FluentProvider>);
}

export function mountDashboardCalendar(): void {
  const client = {
    async getCalendarView() {
      window.__spfxDashboardCalendarCalls++;
      return { items: [{ id: "meeting", subject: "Synthetic planning meeting", start: { dateTime: "2026-10-29T09:00:00", timeZone: "UTC" } }], complete: true };
    },
  } as unknown as Client;
  root.render(<WorkDashboard
    client={client}
    services={{} as React.ComponentProps<typeof WorkDashboard>["services"]}
    siteId="tenant,site,web"
    webTitle="Demo site"
    taskListId=""
    projectListId=""
  />);
}

export function mountDeviceCapture(): void {
  window.__spfxFastBarcodeScans = 0;
  window.__spfxFastBitmapCloses = 0;
  Object.defineProperty(globalThis, "BarcodeDetector", {
    configurable: true,
    value: class {
      async detect(_source: ImageBitmap): Promise<Array<{ rawValue: string }>> {
        window.__spfxFastBarcodeScans++;
        return [{ rawValue: "ASSET-42" }];
      }
    },
  });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({ close() { window.__spfxFastBitmapCloses++; } }),
  });
  Object.defineProperty(navigator, "geolocation", {
    configurable: true,
    value: { getCurrentPosition(success: (position: { coords: { latitude: number; longitude: number; accuracy: number } }) => void) {
      window.__spfxFastGeolocationCalls++;
      success({ coords: { latitude: 51.5, longitude: -0.12, accuracy: 20 } });
    } },
  });
  root.render(<FluentProvider theme={sampleLightTheme}><DeviceCapture
    onBarcode={(value) => { window.__spfxFastLastBarcode = value; }}
    onLocation={(point) => { window.__spfxFastLocation = point; }}
  /></FluentProvider>);
}

export function mountDeviceCaptureWithoutScanner(): void {
  Object.defineProperty(globalThis, "BarcodeDetector", { configurable: true, value: undefined });
  root.render(<FluentProvider theme={sampleLightTheme}><DeviceCapture onBarcode={(value) => { window.__spfxFastLastBarcode = value; }} onLocation={() => undefined} /></FluentProvider>);
}

export function mountDeviceCaptureLocationDenied(): void {
  Object.defineProperty(navigator, "geolocation", {
    configurable: true,
    value: { getCurrentPosition(_success: unknown, error?: (failure: { message: string }) => void) { error?.({ message: "Location permission denied" }); } },
  });
  root.render(<FluentProvider theme={sampleLightTheme}><DeviceCapture onBarcode={() => undefined} onLocation={() => undefined} /></FluentProvider>);
}

export function mountFloorPlanMap(): void {
  root.render(<FluentProvider theme={sampleLightTheme}><FloorPlanMap
    title="Office locations"
    locations={[
      { id: "desk-a", label: "Desk A", x: 24, y: 32, detail: "Available" },
      { id: "room-4", label: "Room 4", x: 72, y: 66, detail: "Meeting room" },
      { id: "invalid", label: "Outside map", x: 120, y: -4 },
    ]}
    onSelect={(location) => { window.__spfxFastSelectedLocation = location.id; }}
  /></FluentProvider>);
}

export async function mountDriveFloorPlanMap(): Promise<void> {
  const canvas = document.createElement("canvas");
  canvas.width = 800;
  canvas.height = 600;
  const context = canvas.getContext("2d")!;
  context.fillStyle = "#f4f8f7";
  context.fillRect(0, 0, 800, 600);
  context.fillStyle = "#fff";
  context.fillRect(24, 24, 752, 552);
  const rooms = [
    { x: 40, y: 40, width: 300, height: 208, color: "#eaf4f8" },
    { x: 360, y: 40, width: 400, height: 208, color: "#fff4d9" },
    { x: 40, y: 288, width: 300, height: 272, color: "#e8f4ed" },
    { x: 360, y: 288, width: 400, height: 272, color: "#edf0fa" },
  ];
  for (const room of rooms) {
    context.fillStyle = room.color;
    context.fillRect(room.x, room.y, room.width, room.height);
  }
  context.strokeStyle = "#66818a";
  context.lineWidth = 8;
  context.strokeRect(24, 24, 752, 552);
  context.lineWidth = 6;
  context.beginPath();
  context.moveTo(350, 40); context.lineTo(350, 560); context.moveTo(40, 268); context.lineTo(760, 268);
  context.stroke();
  context.fillStyle = "#294c55";
  context.font = "600 21px sans-serif";
  context.fillText("RECEPTION", 64, 78);
  context.fillText("MEETING ROOM", 392, 78);
  context.fillText("OPEN WORKSPACE", 64, 326);
  context.fillText("QUIET ROOMS", 392, 326);
  context.fillStyle = "#fff";
  context.strokeStyle = "#9bb0b4";
  context.lineWidth = 3;
  context.fillRect(78, 108, 194, 48); context.strokeRect(78, 108, 194, 48);
  context.fillRect(450, 106, 220, 78); context.strokeRect(450, 106, 220, 78);
  context.fillRect(80, 362, 88, 64); context.strokeRect(80, 362, 88, 64);
  context.fillRect(194, 362, 88, 64); context.strokeRect(194, 362, 88, 64);
  context.fillRect(432, 368, 118, 112); context.strokeRect(432, 368, 118, 112);
  context.fillRect(574, 368, 142, 112); context.strokeRect(574, 368, 142, 112);
  context.fillStyle = "#58757d";
  context.font = "600 14px sans-serif";
  context.fillText("FRONT DESK", 116, 190);
  context.fillText("6 SEATS · DISPLAY", 482, 210);
  context.fillText("DESK BANK A", 104, 454);
  context.fillText("QUIET ZONE · 4 ROOMS", 474, 516);
  const image = await new Promise<Blob>((resolve, reject) => canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("Could not create sample image")), "image/png"));
  const createUrl = URL.createObjectURL.bind(URL);
  const revokeUrl = URL.revokeObjectURL.bind(URL);
  URL.createObjectURL = (blob) => { window.__spfxFloorPlanUrlsCreated++; return createUrl(blob); };
  URL.revokeObjectURL = (url) => { window.__spfxFloorPlanUrlsRevoked++; revokeUrl(url); };
  const client = {
    async readDriveItemContent() {
      window.__spfxFloorPlanReads++;
      return { data: image, source: "cache" as const, refresh: Promise.resolve(image) };
    },
  } as unknown as Client;
  root.render(<FluentProvider theme={sampleLightTheme}><DriveFloorPlanMap
    client={client}
    driveId="drive-one"
    itemId="plan-one"
    title="Secure office plan"
    locations={[
      { id: "reception", label: "Reception desk", x: 18, y: 22, detail: "Front desk" },
      { id: "meeting-1", label: "Meeting 1", x: 68, y: 25, detail: "6 seats" },
      { id: "desk-1", label: "Desk 1", x: 26, y: 64, detail: "Available" },
      { id: "quiet-room", label: "Quiet rooms", x: 80, y: 70, detail: "Four phone rooms" },
    ]}
  /></FluentProvider>);
}

export function unmountDriveFloorPlanMap(): void {
  root.unmount();
}

export function mountDriveContentIdentityProbe(): void {
  const client = {
    readDriveItemContent(_driveId: string, itemId: string) {
      if (itemId === "first-file") {
        return Promise.resolve({ data: "First cached file", source: "cache" as const, refresh: new Promise<Blob | string>(() => undefined) });
      }
      window.__spfxSecondFileRequested = true;
      return new Promise<never>(() => undefined);
    },
  } as unknown as Client;
  function Probe(): React.ReactElement {
    const [itemId, setItemId] = React.useState("first-file");
    const file = useDriveItemContent(client, "drive-one", itemId, { as: "text" });
    return <FluentProvider theme={sampleLightTheme}>
      <button type="button" onClick={() => setItemId("second-file")}>Load another file</button>
      <p role="status">{file.loading ? "Loading file" : ""}</p>
      <output aria-label="Loaded file">{typeof file.content === "string" ? file.content : ""}</output>
    </FluentProvider>;
  }
  root.render(<Probe />);
}

export async function mountPrivateFloorPlanProbe(): Promise<void> {
  const canvas = document.createElement("canvas");
  canvas.width = 800;
  canvas.height = 600;
  const image = await new Promise<Blob>((resolve, reject) => canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("Could not create test image")), "image/png"));
  const client = {
    async uploadDriveItem(_driveId: string, _parentId: string, name: string, _file: Blob) {
      window.__spfxPrivateFloorPlanUploads++;
      return { id: "temporary-plan", name };
    },
    async readDriveItemContent() {
      window.__spfxPrivateFloorPlanReads++;
      return { data: image, source: "network" as const, refresh: Promise.resolve(image) };
    },
    async deleteDriveItem() { window.__spfxPrivateFloorPlanDeletes++; },
  } as unknown as Client;
  root.render(<FluentProvider theme={sampleLightTheme}><PrivateFloorPlanProbe client={client} driveId="synthetic-drive" /></FluentProvider>);
}

declare global {
interface Window { __spfxFastSubmitCount: number; __spfxFastReferenceSubmit?: Record<string, unknown>; __spfxFastListCalls: number; __spfxFastRefreshCalls: number; __spfxFastStorageErrors: number; __spfxFastQuotaResult: { firstSource: string; secondSource: string; storageErrorName: string }; __spfxFastXss?: boolean; __spfxDriveDownloads: number; __spfxAttachmentDownloads: number; __spfxAttachmentMutations: number; __spfxFastGeolocationCalls: number; __spfxFastBarcodeScans: number; __spfxFastBitmapCloses: number; __spfxFastLastBarcode?: string; __spfxFastLocation?: { latitude: number; longitude: number; accuracy: number }; __spfxFastSelectedLocation?: string; __spfxDashboardCalendarCalls: number; __spfxFastDeltaMaxPages?: number; __spfxFastTenantDeletes: number; __spfxFloorPlanReads: number; __spfxFloorPlanUrlsCreated: number; __spfxFloorPlanUrlsRevoked: number; __spfxPrivateFloorPlanUploads: number; __spfxPrivateFloorPlanReads: number; __spfxPrivateFloorPlanDeletes: number; __spfxSecondFileRequested?: boolean; }
}

window.__spfxFastSubmitCount = 0;
window.__spfxFastReferenceSubmit = undefined;
window.__spfxFastTenantDeletes = 0;
window.__spfxFloorPlanReads = 0;
window.__spfxFloorPlanUrlsCreated = 0;
window.__spfxFloorPlanUrlsRevoked = 0;
window.__spfxSecondFileRequested = false;
window.__spfxPrivateFloorPlanUploads = 0;
window.__spfxPrivateFloorPlanReads = 0;
window.__spfxPrivateFloorPlanDeletes = 0;
window.__spfxFastListCalls = 0;
window.__spfxFastRefreshCalls = 0;
window.__spfxFastStorageErrors = 0;
window.__spfxFastQuotaResult = { firstSource: "", secondSource: "", storageErrorName: "" };
window.__spfxDriveDownloads = 0;
window.__spfxAttachmentDownloads = 0;
window.__spfxAttachmentMutations = 0;
window.__spfxFastGeolocationCalls = 0;
window.__spfxDashboardCalendarCalls = 0;
