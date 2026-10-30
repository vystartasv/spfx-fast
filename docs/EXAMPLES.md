# Copyable recipes for SPFx Fast

Recipes use the signed-in SPFx identity. See [authenticated services](SPFX-EXAMPLES.md#create-one-authenticated-data-client). Keep cache scope stable; declare Graph permissions and obtain admin approval.

| # | Recipe | Surface | Capability | Tenant verification |
|---:|---|---|---|---|
| 1 | [Create authenticated services](SPFX-EXAMPLES.md#create-one-authenticated-data-client) | SPFx | Graph + REST | Build and signed-in tenant |
| 2 | [Resolve site by path](#resolve-a-site) | Graph | `site.resolve` | Signed-in site access |
| 3 | [Discover lists](#discover-lists) | Graph | `list.discover` | Signed-in site access |
| 4 | [Inspect fields](#inspect-list-fields) | Graph | `list.inspect` | List read permission |
| 5 | [Read bounded pages](#read-a-bounded-page) | Graph | `list.items.page` | List read permission |
| 6 | [Read all pages safely](#iterate-pages-with-a-bound) | Graph | `list.items.page` | List read permission |
| 7 | [Cached-first read](#show-cached-data-and-refresh) | Graph/cache | `client.storage` | Browser storage enabled |
| 8 | [Sync incrementally](#sync-a-list-with-delta) | Graph delta | `list.sync.delta` | List read permission |
| 9 | [Create and update](#create-and-update-items) | Graph | `item.create`, `item.update` | List write permission |
| 10 | [Batch independent requests](#batch-independent-requests) | Graph | `graph.batch` | Endpoint permissions |
| 11 | [Run CAML](#run-caml-through-rest) | SharePoint REST | `sharepoint.caml` | Site/list access |
| 12 | [Enable audience targeting](#enable-audience-targeting) | SharePoint REST | `sharepoint.audienceTargeting` | Site/list access; tenant test |
| 13 | [Download a file](#download-a-file) | Graph | `graph.raw` | File read permission |
| 14 | [Read Markdown safely](#read-markdown-as-data) | Graph | `graph.raw` | File read permission |
| 15 | [Render safe vanilla DOM](SPFX-EXAMPLES.md#vanilla-dom-page-inside-an-spfx-component) | DOM | `list.items.page` | Signed-in SPFx host |
| 16 | [Build a Fluent UI v9 dashboard](SPFX-EXAMPLES.md#fluent-ui-v9-single-page-dashboard) | React + Fluent UI v9 | `useItems` | Signed-in SPFx host |
| 17 | [Use Graph's raw escape hatch](#call-an-untyped-graph-endpoint) | Graph | `graph.raw` | Endpoint permission |
| 18 | [Use SharePoint Search REST](#search-with-sharepoint-rest) | SharePoint REST | `sharepoint.search` | Search permission/access |
| 19 | [Embed the custom list element](#embed-the-custom-list-element) | Custom element | `ui.custom-element` | Must be initialized inside authenticated SPFx host |
| 20 | [Upload a large file](#upload-a-large-file) | Graph upload session | `sharepoint.drive-items` | Graph file write permission and browser upload URL availability |
| 21 | [Page through a document library folder](#page-through-a-document-library-folder) | Graph | `sharepoint.drive-items` | File read permission |
| 22 | [Manage list item attachments](#manage-list-item-attachments) | SharePoint REST | `sharepoint.attachments` | List read/write access |
| 23 | [Use a cached Markdown file hook](#use-a-cached-markdown-file-hook) | React + Graph | `useMarkdownFile` | File read permission |
| 24 | [Run the SPFx 1.24 dashboard and tenant checks](../examples/fixtures/spfx-fast-dashboard/README.md) | SPFx 1.24 + React 18 + Fluent UI v9 | Authenticated Graph and REST | `Sites.ReadWrite.All` for list/drive actions; `Calendars.ReadBasic` for the calendar view |
| 25 | [Filter a list on the server](#filter-a-list-on-the-server) | Graph | `list.items.page` | Indexed field and list read permission |
| 26 | [Use the paged filtered React hook](#use-the-paged-filtered-react-hook) | React + Fluent UI v9 | `useListItems` | Indexed field and list read permission |
| 27 | [Run the SharePoint REST playground](#run-the-sharepoint-rest-playground) | SPFx REST + Fluent UI v9 | `SharePointRestPlayground.tsx` | Current user's SharePoint access; use a disposable list for writes |
| 28 | [Build a Copilot UX list browser](#build-a-copilot-ux-list-browser) | SPFx Copilot UX + Graph + Fluent UI v9 | `spfx-fast-copilot` fixture | Sites.Read.All approval and Microsoft 365 Copilot add-on license |
| 29 | [Read visible file sharing permissions](#read-visible-drive-item-permissions) | Authenticated Graph | `sharepoint.drive-items` | File permission read access |
| 30 | [Build an authenticated file workflow](../examples/fixtures/spfx-fast-dashboard/README.md) | SPFx + Fluent UI v9 | `sharepoint.drive-items` | File read/write access and upload-session host validation |
| 31 | [Inspect list content types](#inspect-list-content-types) | Authenticated Graph | `list.content-types` | Sites.Read.All approval and list access |
| 32 | [Load managed metadata for a picker](#load-managed-metadata-for-a-picker) | Authenticated Graph | `sharepoint.taxonomy` | TermStore.Read.All admin approval and user term access |
| 33 | [Inspect and manage list permission scope](#inspect-and-manage-a-lists-permission-scope) | Authenticated SharePoint REST | `sharepoint.role-assignments` | Current user must be allowed to read/change the list's security configuration; writes require confirmation |
| 34 | [Load site people and groups](#load-site-people-and-groups) | Authenticated SharePoint REST | `sharepoint.users-profiles-groups` | Requires visibility to users and groups in the hosting web; output is paged and does not resolve Entra membership |
| 35 | [Configure SharePoint read retries](#configure-sharepoint-read-retries) | Authenticated SPFx transports | `client.retry-policy` | Reads retry 429/503 responses; writes are not retried by default |
| 36 | [Keep local list data fresh](#keep-local-list-data-fresh) | React hook + Graph delta | `client.sync-policy` | Sync interval and foreground behavior are browser-tab local |
| 37 | [Delete after explicit confirmation](#delete-after-explicit-confirmation) | Graph write | `item.delete` | The hosting UI or agent must obtain the user's confirmation immediately before deletion |
| 38 | [Render a SharePoint list view](#render-a-sharepoint-list-view) | Authenticated SharePoint REST | `sharepoint.renderListDataAsStream` | Legacy response shape and view behavior need tenant validation |
| 39 | [Read a SharePoint file with REST](#read-a-sharepoint-file-with-rest) | Authenticated SharePoint REST | `sharepoint.files-folders` | Use a server-relative path from inspected metadata; binary response is not cached |
| 40 | [Inspect SharePoint list views](#inspect-sharepoint-list-views) | Authenticated SharePoint REST | `sharepoint.views-content-types` | Public views and the current user's personal views; view definitions may contain legacy CAML |
| 41 | [Read SharePoint user profile properties](#read-sharepoint-user-profile-properties) | Authenticated SharePoint REST PeopleManager | `sharepoint.user-profiles` | Profile visibility and property availability vary by tenant; use a claims login, not an Entra object ID |
| 42 | [Read list and file version history](#read-list-and-file-version-history) | Authenticated Microsoft Graph v1.0 | `sharepoint.versions` | Requires list `Sites.Read.All` or work/school file `Files.Read`; version retention is service-controlled |
| 43 | [Synchronize a document library with delta](#sync-a-document-library-with-delta) | Authenticated Microsoft Graph v1.0 | `sharepoint.drive-items` | Requires delegated work/school `Files.Read`; large initial enumerations resume across bounded calls |
| 44 | [Create, move, and delete drive items](#create-move-and-delete-drive-items) | Authenticated Microsoft Graph v1.0 | `sharepoint.drive-items` | Requires delegated work/school `Files.ReadWrite`; deletion moves items to the recycle bin |
| 45 | [Download or restore a file version](#download-or-restore-a-file-version) | Authenticated Microsoft Graph v1.0 | `sharepoint.versions` | Requires `Files.Read` to download or `Files.ReadWrite` to restore; restore adds a new current version |
| 46 | [Create a sharing link or remove direct access](#share-a-drive-item) | Authenticated Microsoft Graph v1.0 | `sharepoint.drive-items` | Requires `Files.ReadWrite`; confirm the explicit audience and any direct permission removal |
| 47 | [Copy a drive item with progress](#copy-a-drive-item-with-progress) | Authenticated Graph + SharePoint REST monitor | `sharepoint.drive-items` | Requires `Files.ReadWrite`; copy runs asynchronously and uses destination-folder permissions |
| 48 | [Capture a barcode or current location](#capture-a-barcode-or-current-location) | Browser device APIs + Fluent UI v9 | `DeviceCapture.tsx` | Requires a secure context and user-granted browser permissions; barcode detection varies by browser |
| 49 | [Read a calendar window](#read-a-calendar-window) | Authenticated Microsoft Graph v1.0 | `graph.calendar.view` | Requires `Calendars.ReadBasic`; reads the signed-in user's Outlook calendar |
| 50 | [Query cached list items](#query-cached-list-items) | Local cache | `client.queryItems` | Sync the list first; server access still controls freshness and permissions |
| 51 | [Show locations on a floor plan](#show-locations-on-a-floor-plan) | Authenticated Graph + React | `DriveFloorPlanMap.tsx` | Work/school delegated `Files.Read`; loads a private image through the cached signed-in file hook |
| 52 | [Manage the SharePoint recycle bin](#manage-the-recycle-bin) | Authenticated SharePoint REST | `sharepoint.recycle-bin` | Restore requires access to the item; permanent deletion requires explicit user confirmation |
| 53 | [Read web regional settings](#read-web-regional-settings) | Authenticated SharePoint REST | `sharepoint.regional-settings` | Read-only settings for the hosting web; site usage remains raw-only |
| 54 | [Read SharePoint navigation](#read-sharepoint-navigation) | Authenticated SharePoint REST | `sharepoint.navigation` | Bounded reads of web or site-root navigation; writes remain raw-only |

## Build a Copilot UX list browser

The [SPFx 1.24.0 Copilot fixture](../examples/fixtures/spfx-fast-copilot/README.md) reads an exact user-named list via `createSPFxServices`, shows up to 50 items, and shares up to 10 titles after user action. It validates SharePoint URLs and never writes.

```sh
pnpm test:copilot-fixture
# npm: npm run test:copilot-fixture
```

The build validates the manifest and creates a `.sppkg`. Use requires a Copilot add-on license and admin-approved `Sites.Read.All`. GA supports browser/desktop Microsoft 365 Copilot and Copilot in Teams, not mobile. Check [SPFx 1.24 release notes](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/release-1.24.0) before deployment.

## Run the SharePoint REST playground

`SharePointRestPlayground` is an embeddable React example for Search REST, CAML reads, list attachments, and audience targeting. It uses the caller's SPFx-authenticated `SPHttpClient` transport and includes no token handling.

```tsx
import { SharePointRestPlayground } from "./SharePointRestPlayground";
import { createSPFxServices } from "@spfx-fast/core/spfx";
import { SPHttpClient } from "@microsoft/sp-http";

const services = await createSPFxServices(this.context, SPHttpClient.configurations.v1);
return <SharePointRestPlayground rest={services.rest} listId={testListId} itemId={numericTestItemId} />;
```

Use a disposable list with attachments and a test item. Search/CAML use current SharePoint access; attachments need item edit rights. Audience targeting changes a list property, so use list-management rights and disable it afterward. Remove the test attachment with `services.sharepoint.deleteItemAttachment(listId, itemId, fileName)`, then delete test data and restore the setting. No Graph permission is needed.

See [SharePointRestPlayground.tsx](../examples/SharePointRestPlayground.tsx). `pnpm build` or `npm run build` type-checks it; verify audience targeting in the target tenant.

## Resolve a site

```ts
const site = await client.getSite("contoso.sharepoint.com", "/sites/operations");
console.log(site.id, site.webUrl);
```

## Discover lists

```ts
const lists = await client.getLists(site.id);
const tasks = lists.find((list) => list.displayName === "Tasks");
if (!tasks) throw new Error("Tasks list not found; inspect the site before scaffolding.");
```

## Inspect list fields

```ts
const { list, columns } = await client.inspectList(site.id, tasks.id);
console.table(columns.map((column) => ({ name: column.name, displayName: column.displayName })));
```

Use the internal `name` for API bindings and the display name for labels.

## Read a bounded page

```ts
const page = await client.getItemPage(site.id, tasks.id, { fields: ["Title", "Status"], pageSize: 50 });
console.log(page.items, page.nextCursor); // Keep the opaque cursor for the next call.
```

## Iterate pages with a bound

```ts
for await (const item of client.listItems(site.id, tasks.id, { fields: ["Title"], maxItems: 500 })) {
  renderRow(item);
}
```

## Filter a list on the server

Use typed filters for focused reads so Graph returns only matching rows. Filtering is most reliable on indexed columns; SharePoint supports `eq`, `ne`, comparisons, and `startswith` for list items. The builder escapes string literals and rejects untrusted field/operator syntax.

```ts
for await (const task of client.listItems(siteId, listId, {
  fields: ["Title", "Status"],
  filter: { field: "Status", operator: "eq", value: "Open" },
  maxItems: 500,
})) {
  renderTask(task);
}
```

The iterator follows Graph continuation links and enforces `maxItems`. [Microsoft Graph list-item filtering](https://learn.microsoft.com/en-us/graph/api/listitem-list?view=graph-rest-1.0)

## Use the paged filtered React hook

`useListItems` serves cached Graph pages immediately, refreshes them in the background, and exposes explicit `loadMore` and `maxItems` controls. Render it inside a `FluentProvider`.

```tsx
const tasks = useListItems(client, siteId, listId, {
  fields: ["Title", "Status"],
  filter: { field: "Status", operator: "eq", value: "Open" },
  pageSize: 50,
  maxItems: 500,
});
return <>
  <ListView items={tasks.items} fields={fields} loading={tasks.loading} error={tasks.error} />
  {tasks.hasMore && <Button onClick={() => void tasks.loadMore()}>Load more</Button>}
</>;
```

See the typechecked [filtered task example](../examples/FilteredTasks.tsx).

The dashboard CLI can generate the same mode from inspected schema JSON. Add a `filter`, optional `pageSize` (1–100), and `maxItems` (1–10,000) to a list entry. Lists without `filter` keep cache-first delta sync; filtered entries use bounded server paging and render a Load more button. `ListView` paginates the loaded rows with previous/next controls.

Generated dashboards use `ItemForm` for create and edit. Required text/date/number controls use native browser validation. Required choice, Person, and Lookup fields receive an accessible error message, invalid field state, and focus when submitted empty; `getMissingRequiredFieldNames` is available from `@spfx-fast/core/react` when an app needs the same check in a custom form.

```json
{
  "siteId": "tenant,site,web",
  "lists": [{
    "title": "Open requests",
    "listId": "inspected-list-id",
    "filter": { "field": "Status", "operator": "eq", "value": "Open" },
    "pageSize": 50,
    "maxItems": 500,
    "fields": [
      { "name": "Title", "type": "text", "required": true },
      { "name": "Status", "type": "choice", "choices": ["Open", "Closed"] }
    ]
  }]
}
```

Use an indexed SharePoint column for the filter and confirm the internal field name during inspection. This generated view is server-filtered and paged; it does not populate the full local delta cache.

## Use Person and Lookup form fields

For `person` and `lookup`, pass `{ value, label }` options from authenticated SPFx services. `ItemForm` displays labels and submits Graph's `<FieldName>LookupId` field. Generated dashboards also load that ID for edits:

```tsx
const fields: FieldDefinition[] = [
  { name: "Title", title: "Request", type: "text", required: true },
  { name: "Owner", title: "Owner", type: "person", options: users.map((user) => ({ value: user.Id, label: user.Title ?? user.Email ?? String(user.Id) })) },
  { name: "Project", title: "Project", type: "lookup", options: projects.map((item) => ({ value: item.id, label: String(item.fields?.Title ?? item.id) })) },
];
```

Page people with `getSiteUsersPage`; page lookup items with `client.listItems`. The app chooses and filters the target list. See the typechecked [Person and Lookup form](../examples/PersonLookupForm.tsx). It caps each source at 1,000 entries. Verify writes in the target tenant.

## Show cached data and refresh

```ts
const page = await client.readPage("https://graph.microsoft.com/v1.0/sites/.../items");
render(page.data.value, { source: page.source });
void page.refresh.then((fresh) => render(fresh.value, { source: "network" }));
```

## Sync a list with delta

```ts
const result = await client.syncList(site.id, tasks.id, { maxPages: 100 });
console.log(`Updated ${result.changed} items across ${result.pages} pages`);
```

The initial run enumerates the list. Later runs use the saved delta cursor; an expired cursor triggers a bounded full rebuild.

## Query cached list items

Sync the list first, then filter and sort only the fields the app needs:

```ts
const sync = await client.syncList(site.id, tasks.id, { fields: ["Title", "Status"], maxPages: 25 });
const openTasks = await client.queryItems(site.id, tasks.id, {
  equals: { Status: "Open" },
  sort: { field: "Title" },
  limit: 50,
});
```

If `sync.complete` is false, call `syncList` again to resume. Exact local queries with limits up to 100 stay in memory for one second and clear when that list changes. A cold sorted query reads the whole local list; use a server filter and `getItemPage` for large remote lists. `queryCachedTasks` in [`recipes.ts`](../examples/recipes.ts) shows this pattern.

## Create and update items

```ts
const created = await client.createItem(site.id, tasks.id, { Title: "Draft plan", Status: "Open" });
const updated = await client.updateItem(site.id, tasks.id, created.id, { Status: "In progress" });
```

Show users the proposed write and require explicit confirmation for agent-triggered updates or deletes.

## Batch independent requests

```ts
const results = await client.batch([
  { id: "tasks", method: "GET", url: `/sites/${site.id}/lists/${tasks.id}/items?$top=10` },
  { id: "projects", method: "GET", url: `/sites/${site.id}/lists/${projects.id}/items?$top=10` },
]);
for (const response of results) if (response.status >= 400) console.error(response.id, response.status, response.body);
```

Graph batches contain at most 20 subrequests. A `200` batch response does not mean every subrequest succeeded. Check each status and handle throttling per subrequest.

## Run CAML through REST

```ts
const rows = await sharepoint.getItemsByCaml(listId, `<View><Query><Where><Eq><FieldRef Name="Status"/><Value Type="Text">Open</Value></Eq></Where></Query><RowLimit>100</RowLimit></View>`);
```

`getItemsByCaml` and other legacy wrapper names are defined by the installed version's `@spfx-fast/core/sharepoint` declarations; consult them before copying this recipe.

## Enable audience targeting

```ts
await sharepoint.setAudienceTargeting(listId, true);
```

This endpoint has tenant-specific behavior; validate it in a test list before production use. The operation requires the signed-in user to have list administration rights.

## Download a file

```ts
const response = await client.readDriveItemContent(driveId, fileId, { as: "blob" });
const file = response.data; // Blob; warm reads return source: "cache" and refresh in the background.
```

For binary calls, confirm that the configured SPFx Graph client returns a Blob for `responseType("blob")` in the target host.

## Page through a document library folder

```ts
const page = await client.getDriveItemPage(driveId, folderItemId, { pageSize: 50 });
renderFiles(page.items);
if (page.nextCursor) renderMoreButton(() => client.getDriveItemPage(driveId, folderItemId, { cursor: page.nextCursor }));
```

Continuation cursors are bound to the requested drive and folder path. Keep reads bounded and use each opaque cursor only with the same folder.

## Read Markdown as data

```ts
const response = await client.readDriveItemContent(driveId, fileId, { as: "text" });
const markdown = String(response.data);
markdownElement.textContent = markdown; // Render as literal text, or use a renderer with raw HTML disabled.
```

SharePoint's native Markdown viewer is the simplest link target. A `.md` file is document content; it does not execute SPFx modules.

## Use a cached Markdown file hook

Inside an SPFx-hosted React component, `useMarkdownFile` reads authenticated Graph file content as text and exposes cached-first refresh state. Render it with the safe Markdown component for headings, links, lists, code, and tables:

```tsx
import { useMarkdownFile } from "@spfx-fast/core/react";
import { MarkdownContent } from "@spfx-fast/core/markdown";

const file = useMarkdownFile(client, driveId, itemId);
return file.markdown === undefined ? <p>Loading…</p> : <MarkdownContent>{file.markdown}</MarkdownContent>;
```

`@spfx-fast/core/markdown` supports GitHub-flavored tables and strikethrough, sanitizes rendered elements, and leaves raw HTML inert. Keep interactive behavior in compiled SPFx React; do not evaluate Markdown files as MDX. SharePoint's native `.md` viewer is for file content, not an SPFx runtime.

In React, `useDriveItemContent(client, driveId, fileId, { as: "text" })` exposes `content`, `source`, `loading`, `error`, and `refresh`. Use `String(content ?? "")` as text and never feed the file to `dangerouslySetInnerHTML`.

## Call an untyped Graph endpoint

```ts
const result = await client.requestGraph(`https://graph.microsoft.com/v1.0/sites/${site.id}/lists/${tasks.id}/items?$expand=fields`);
```

Only HTTPS Microsoft Graph URLs are accepted by `requestGraph`; file uploads and other host-specific paths may need their official SPFx client flow.

## Search with SharePoint REST

```ts
const response = await sharepoint.search('"Project Plan"', {
  rowLimit: 20,
  selectProperties: ["Title", "Path", "LastModifiedTime"],
  sortList: "LastModifiedTime:descending",
});
```

The helper returns SharePoint's native result object. Microsoft documents `/_api/search/query` GET parameters, JSON accept headers, paging, managed properties, and refiners in the [Search REST API reference](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/sharepoint-search-rest-api-overview). Use `sharepoint.raw` for complex POST queries and `client.requestGraph` for Microsoft Graph Search; their payloads and result shapes differ.

## Embed the custom list element

```ts
import { defineSPFxFastListElement } from "@spfx-fast/core/elements";
import { createSPFxServices } from "@spfx-fast/core/spfx";
import { SPHttpClient } from "@microsoft/sp-http";

defineSPFxFastListElement();
const services = await createSPFxServices(this.context, SPHttpClient.configurations.v1);
const list = document.createElement("spfx-fast-list") as HTMLElement & { client: typeof services.client };
list.client = services.client; // Pass the existing signed-in client; no token is exposed.
list.setAttribute("site-id", siteId);
list.setAttribute("list-id", listId);
list.setAttribute("title", "Team tasks");
list.setAttribute("fields", "Title,Status,DueDate");
this.domElement.replaceChildren(list);
```

The element uses a shadow root, a bounded read, accessible status text and table headers, and `textContent` for SharePoint values. A standalone `.html` page still needs an SPFx host to pass the authenticated client; a Markdown file opened in SharePoint's viewer cannot execute this code.

## Upload a large file

```ts
try {
  const result = await client.uploadDriveItem(driveId, folderItemId, file.name, file, {
    onProgress: (uploaded, total) => setProgress(Math.round(uploaded / total * 100)),
  });
  return result;
} catch (error) {
  if (isDriveUploadInterruptedError(error)) setUploadRecovery(error);
  else throw error;
}
```

Uploads use authenticated session creation, then sequential chunks to Graph's preauthenticated URL (no Authorization header). Chunks default to 10 MiB and must be a 320 KiB multiple below 60 MiB. If retries expire, keep `DriveUploadInterruptedError` in page state and call `await error.resume()` after reconnecting. It retains the file/session in memory; reload loses it. Never expose the session URL. Validate CORS/CSP. See [Graph resumable upload](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0).

Default conflict behavior is `fail`. Confirm before `replace`. After a lost final response, `fail`/`replace` report success only after read-back and byte comparison. `rename` checks a bounded set of Graph suffixes and accepts only an exact byte match. Fixture 1.0.0.60 passed in a signed-in tenant. For large uploads, do not repeat `item.name` in the body when the path already names it; this tenant returned `400 invalidRequest`. See [tenant notes](TENANT-VALIDATION.md#lessons-captured).

See the [SPFx dashboard fixture](../examples/fixtures/spfx-fast-dashboard/README.md) for a Fluent UI v9 file browser. `downloadDriveItem(driveId, itemId)` fetches fresh bytes; use `readDriveItemContent` or `useDriveItemContent` for cached-first preview.

## Read visible drive item permissions

```ts
const permissions = await client.getDriveItemPermissions(driveId, itemId);
for (const permission of permissions) {
  console.log(permission.roles, permission.grantedToV2?.user?.displayName, permission.link?.scope);
}
```

Uses signed-in Graph with bounded paging. Permissions are uncached; never log `link.webUrl`, `shareId`, or invitation secrets. Work/school delegated `Files.Read` is least privilege. Owners may see more than non-owners; this is not a full ACL audit. See [list permissions](https://learn.microsoft.com/en-us/graph/api/driveitem-list-permissions?view=graph-rest-1.0) and [Permission resource](https://learn.microsoft.com/en-us/graph/api/resources/permission?view=graph-rest-1.0).

## Share a drive item

```ts
// Ask the user to confirm the audience and access level before creating the link.
const permission = await client.createDriveItemSharingLink(driveId, itemId, {
  type: "view",
  scope: "organization",
  expirationDateTime: "2026-12-31T23:59:59Z",
});
showLink(permission.link?.webUrl);

// After the user confirms this specific direct grant should be revoked:
await client.deleteDriveItemPermission(driveId, itemId, permission.id);
```

Specify the link scope; anonymous access is never implicit. Creation preserves inherited permissions. Removal re-reads the grant and rejects inherited access. Show recipient/scope and get confirmation before removal. Delegated `Files.ReadWrite` is least privilege; tenant policy may block anonymous links or set expiry. See [`DriveItemPermissions.ts`](../examples/DriveItemPermissions.ts), [create a link](https://learn.microsoft.com/en-us/graph/api/driveitem-createlink?view=graph-rest-1.0), and [remove permission](https://learn.microsoft.com/en-us/graph/api/permission-delete?view=graph-rest-1.0).

## Copy a drive item with progress

```ts
const operation = await services.client.copyDriveItem(sourceDriveId, sourceItemId, {
  targetDriveId,
  targetFolderId,
  conflictBehavior: "rename",
});
const result = await services.waitForDriveCopy(operation, {
  onProgress: ({ percentComplete }) => setCopyProgress(percentComplete ?? 0),
});
console.log(result.resourceLocation);
```

Graph copies asynchronously. `waitForDriveCopy` polls its monitor URL through same-origin REST (2-minute default, 10-minute max), reports progress, and supports abort. Copies inherit destination permissions and omit metadata. Confirm `conflictBehavior: "replace"`; it deletes the destination and its history. Renaming and preserving version history cannot be combined. See [`DriveCopy.ts`](../examples/DriveCopy.ts) and Microsoft's [driveItem copy API](https://learn.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0).

## Capture a barcode or current location

```tsx
<DeviceCapture
  onBarcode={(value) => setAssetTag(value)}
  onLocation={(point) => setCoordinates(point)}
/>
```

`DeviceCapture` uses the browser's photo picker/camera capture, `BarcodeDetector` when available, and geolocation after a button click. Manual barcode entry remains available when scanning is unsupported. These browser APIs require HTTPS and may be blocked by host or device policy. The component returns values to your app; save them to an inspected SharePoint list with the authenticated client. See [`DeviceCapture.tsx`](../examples/DeviceCapture.tsx).

## Read a calendar window

```tsx
<CalendarAgenda client={services.client} days={7} />
```

The helper loads expanded event instances for a bounded date window, follows Graph pages, and keeps a cursor when more events remain. Calendar data is read directly and never cached. Declare delegated `Calendars.ReadBasic` in the SPFx solution, obtain admin approval, and ensure the signed-in account has an active Exchange mailbox enabled for REST. This reads only the signed-in user's default Outlook calendar; it does not write events. See [`CalendarAgenda.tsx`](../examples/CalendarAgenda.tsx) and Microsoft's [calendarView API](https://learn.microsoft.com/en-us/graph/api/user-list-calendarview?view=graph-rest-1.0).

## Show locations on a floor plan

Use the component when list rows already contain normalized map coordinates. Inspect the list schema first; store `MapX` and `MapY` as numbers from 0 to 100. Read rows through the authenticated SPFx client, then map them into the view. `DriveFloorPlanMap` also loads a private PNG, JPEG, or WebP plan through Graph and revokes its temporary object URL when the file changes or the component unmounts:

```tsx
import { DriveFloorPlanMap } from "@spfx-fast/core/react";

<DriveFloorPlanMap
  client={services.client}
  driveId={driveId}
  itemId={floorPlanItemId}
  title="Office locations"
  locations={items.map((item) => ({
    id: item.id,
    label: String(item.fields?.Title ?? "Unnamed location"),
    x: Number(item.fields?.MapX),
    y: Number(item.fields?.MapY),
    detail: String(item.fields?.Status ?? ""),
  }))}
/>
```

The image uses `useDriveItemContent` and defaults to an 8 MiB cache limit. For local SVGs or other custom visuals, use the lower-level [`FloorPlanMap`](../examples/FloorPlanMap.tsx) `plan` slot. Both components ignore out-of-range coordinates, offer map markers and a keyboard-friendly location list, and make no map-provider calls. This is a floor-plan overlay, not a geocoding or street-map service. See [`DriveFloorPlanMap.tsx`](../examples/DriveFloorPlanMap.tsx).

## Inspect list content types

```ts
const contentTypes = await client.getListContentTypes(siteId, listId);
console.table(contentTypes.map(({ id, name, group, hidden }) => ({ id, name, group, hidden })));
```

This uses Graph v1.0 and follows continuation pages; Microsoft documents delegated `Sites.Read.All` as the least-privileged permission. For SharePoint list views, use the authenticated SharePoint REST transport: Graph has no list views endpoint. The shipped [`ListContentTypes.ts` example](../examples/ListContentTypes.ts) shows the typed helper.

## Load managed metadata for a picker

```ts
const groups = await client.getTermStoreGroups(siteId);
const sets = groups[0] ? await client.getTermStoreSets(siteId, groups[0].id) : [];
const terms = sets[0] ? await client.getTermSetChildren(siteId, sets[0].id) : [];
const nested = terms[0] && sets[0]
  ? await client.getTermChildren(siteId, sets[0].id, terms[0].id)
  : [];
```

Each call loads one level and follows Graph pages, so the UI can fetch children on demand. Requires delegated `TermStore.Read.All`; Graph returns only accessible terms. See [`TaxonomyPickerData.ts`](../examples/TaxonomyPickerData.ts), [groups](https://learn.microsoft.com/en-us/graph/api/termstore-list-groups?view=graph-rest-1.0), [sets](https://learn.microsoft.com/en-us/graph/api/termstore-group-list-sets?view=graph-rest-1.0), and [children](https://learn.microsoft.com/en-us/graph/api/termstore-term-list-children?view=graph-rest-1.0).

## Inspect and manage a list's permission scope

```ts
const scope = await services.sharepoint.getListPermissionScope(listId);
const assignments = await services.sharepoint.getListRoleAssignments(listId);
console.log(scope.HasUniqueRoleAssignments, assignments.map(({ Member, RoleDefinitionBindings }) => ({
  principal: Member?.Title,
  roles: RoleDefinitionBindings?.map(({ Name }) => Name),
})));
```

These signed-in SharePoint REST reads inspect one list scope. A `false` unique-permissions flag means assignments come from a parent scope; this helper does not walk the hierarchy or expand Entra group membership, sharing links, or item-level scopes. Do not use it to claim complete effective access. See the [`ListPermissions.ts` example](../examples/ListPermissions.ts), [SharePoint REST and CSOM API index](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/sharepoint-net-server-csom-jsom-and-rest-api-index), and [SecurableObject inheritance properties](https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-visio/jj245813%28v%3Doffice.15%29).

Permission writes require the UI to confirm the exact list, principal, role, and impact. The helpers reject calls unless `confirmed: true` is supplied. Breaking inheritance copies current assignments and always leaves child-item scopes alone; resetting inheritance removes the list's unique assignments. Inspect the target and existing bindings first:

```ts
const roles = await services.sharepoint.getRoleDefinitionsPage({ pageSize: 100 });
const readRole = roles.items.find((role) => role.Name === "Read");
if (!readRole) throw new Error("The Read permission level is unavailable.");

if (window.confirm(`Grant Read on ${listTitle} to ${principalTitle}?`)) {
  const scope = await services.sharepoint.getListPermissionScope(listId);
  if (!scope.HasUniqueRoleAssignments) {
    await services.sharepoint.breakListRoleInheritance(listId, {
      confirmed: true,
      copyRoleAssignments: true,
    });
  }
  await services.sharepoint.addListRoleAssignment(listId, principalId, readRole.Id, { confirmed: true });
}
```

`removeListRoleAssignment` removes all role bindings for that principal on this list. `resetListRoleInheritance` discards list-specific bindings and restores the parent scope; call either only after a separate impact confirmation. SharePoint enforces the caller's permissions. These operations affect one list and do not grant or calculate effective access for Entra groups, links, or child items. Microsoft documents the [SharePoint REST custom-permissions flow](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/set-custom-permissions-on-a-list-by-using-the-rest-interface).

## Load site people and groups

```ts
const current = await services.sharepoint.getCurrentUser();
const people = await services.sharepoint.getSiteUsersPage({ pageSize: 50 });
const groups = await services.sharepoint.getSiteGroupsPage({ pageSize: 50 });
const members = groups.items[0]
  ? await services.sharepoint.getSiteGroupMembersPage(groups.items[0].Id, { pageSize: 50 })
  : undefined;
// Continue with the opaque cursor returned by each page.
const morePeople = people.nextCursor
  ? await services.sharepoint.getSiteUsersPage({ cursor: people.nextCursor })
  : undefined;
```

These helpers use the authenticated hosting web's `SPHttpClient`, cap page size at 100, and validate that a continuation stays on the same endpoint and web. They are suitable for person pickers and SharePoint group selection; they do not enumerate Entra ID or expand nested group membership. See the [`SitePeople.ts` example](../examples/SitePeople.ts).

## Configure SharePoint read retries

```ts
const retryEvents: Array<{ delayMs: number; error: unknown }> = [];
const services = await createSPFxServices(context, SPHttpClient.configurations.v1, {
  transport: {
    maxRetries: 4,
    maxBackoffMs: 12_000,
    onRetry: (delayMs, error) => retryEvents.push({ delayMs, error }),
  },
  onMetric: (metric) => console.debug("SPFx Fast", metric.type, metric.operation),
});
```

The Graph and REST host adapters honor `Retry-After` for throttled reads and retry HTTP 429/503 responses within the configured attempt/backoff limits. Graph and REST reads also retry recognized network `TypeError`s such as `Failed to fetch`; writes are never retried for a lost network response. Non-idempotent writes are not retried unless the caller explicitly marks a supported REST request `retryable` and knows that retry is safe.

## Keep local list data fresh

```tsx
const result = useItems(client, siteId, listId, {
  fields: ["Title", "Status", "Modified"],
  syncIntervalMs: 60_000,
  syncOnVisible: true,
  maxItems: 2_000,
});
```

`useItems` serves synchronized local records, refreshes through Graph delta, and starts a refresh when the tab becomes visible or focused. `syncIntervalMs` prevents repeated refreshes within one mounted hook; it is not a cross-tab scheduler. Set an explicit `maxItems` for the network fallback used when the first delta sync fails. See the [cached-first and delta recipes](#show-cached-data-and-refresh) and [sync a list with delta](#sync-a-list-with-delta).

## Delete after explicit confirmation

```tsx
async function deleteConfirmedItem(itemId: string) {
  const accepted = await confirmDialog("Delete this SharePoint item?");
  if (!accepted) return;
  await client.deleteItem(siteId, listId, itemId);
}
```

Keep the confirmation in the UI or agent host, close to the destructive action. The client method performs the write and optimistic cache update; it cannot infer whether a user approved it. For agent integrations, mark the delete tool as confirmation-required and verify an explicit confirmation field before executing.

## Render a SharePoint list view

```ts
const result = await services.sharepoint.renderListDataAsStream(listId, "<View><RowLimit>50</RowLimit><ViewFields><FieldRef Name=\"Title\" /></ViewFields></View>");
const rows = (result as { Row?: unknown[] })?.Row ?? [];
```

`RenderListDataAsStream` is useful when SharePoint's view rendering semantics are needed. Its request parameters and response shape are SharePoint-specific; inspect the actual response in the target tenant and prefer Graph list-item paging for ordinary CRUD screens.

## Read a SharePoint file with REST

```ts
const encodedPath = encodeURIComponent(serverRelativePath).replaceAll("%2F", "/");
const response = await services.rest.raw(
  `/_api/web/GetFileByServerRelativeUrl('${encodedPath}')/$value`,
  { responseType: "blob" },
);
const file = await response.blob();
```

Use a server-relative path obtained from the authenticated site or library metadata, not a URL supplied by an untrusted caller. The REST transport stays on the hosting web's HTTPS origin. This raw path bypasses the Graph content cache; use `readDriveItemContent` or `useDriveItemContent` for cached-first file previews.

## Inspect SharePoint list views

```ts
const views = await services.sharepoint.getListViewsPage(listId, { pageSize: 50 });
const selected = views.items.find((view) => view.DefaultView) ?? views.items[0];
const fields = selected?.Id
  ? await services.sharepoint.getListViewFields(listId, selected.Id)
  : [];
console.log(views.items.map(({ Title, DefaultView, ViewQuery }) => ({ Title, DefaultView, ViewQuery })), fields);
```

Helpers page public and caller-visible personal views, then load selected field names. Continuations stay on the same list endpoint and web. Treat legacy CAML `ViewQuery` as untrusted data. See [`ListViews.ts`](../examples/ListViews.ts), [REST/CSOM API index](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/sharepoint-net-server-csom-jsom-and-rest-api-index), and [CSOM `ViewCollection`](https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-visio/jj245667%28v%3Doffice.15%29).

## Manage list item attachments

Use the reusable [ItemAttachmentsPanel](../examples/ItemAttachmentsPanel.tsx) with `ItemForm` for classic list-item attachments. It lists, uploads or replaces, downloads, and confirms deletes. Pass SharePoint's numeric item ID; a Graph item ID is not accepted.

```tsx
import { ItemAttachmentsPanel } from "./ItemAttachmentsPanel";

return <ItemAttachmentsPanel api={services.sharepoint} listId={listId} itemId={sharePointItemId} />;
```

`ItemForm` renders an optional `attachments` slot outside its form element. This keeps file controls independent from the field submit action. The request sample also shows the panel after create when Graph returns a numeric SharePoint item ID:

```tsx
<ItemForm
  fields={fields}
  initial={item.fields}
  onSubmit={saveItem}
  attachments={<ItemAttachmentsPanel api={services.sharepoint} listId={listId} itemId={sharePointItemId} />}
/>
```

For a create flow, pass `sharepoint={services.sharepoint}` to `RequestDashboard`. If the created Graph item ID is not numeric, provide `attachmentItemId` after resolving SharePoint's numeric ID.

```ts
const files = await sharepoint.listItemAttachments(listId, itemId);
const attachment = await sharepoint.readItemAttachment(listId, itemId, "brief.txt");
await sharepoint.addItemAttachment(listId, itemId, "brief.txt", new Blob(["Draft"]));
await sharepoint.updateItemAttachment(listId, itemId, "brief.txt", replacementBlob, currentETag);
// Confirm with the user before deleting an attachment.
await sharepoint.deleteItemAttachment(listId, itemId, "brief.txt", currentETag);
```

These helpers use REST `AttachmentFiles` for list-item attachments; libraries use drive-item APIs. See Microsoft's [attachment REST docs](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest).

## Notes for agents

- Read the installed [`agent/SKILL.md`](../agent/SKILL.md) and [`agent/capabilities.json`](../agent/capabilities.json) first.
- Inspect site/list/columns; never infer IDs, field internal names, or permissions.
- Use one authenticated services instance per tenant/user session and keep all writes user-confirmed.
- Verify each recipe against the installed package declarations. Recipes labeled tenant verification need a signed-in SPFx host; a local build alone does not prove endpoint behavior.

## Read SharePoint user profile properties

```ts
const currentProfile = await services.sharepoint.getMyProperties();
const colleague = await services.sharepoint.getPropertiesFor("i:0#.f|membership|person@contoso.com");
const department = await services.sharepoint.getUserProfilePropertyFor(
  "i:0#.f|membership|person@contoso.com",
  "SPS-Department",
);
```

These read-only `PeopleManager` calls use SharePoint REST. `getPropertiesFor` needs a claims login. Results are uncached and tenant-dependent; they do not enumerate Entra ID or support profile administration/social actions. See [`UserProfiles.ts`](../examples/UserProfiles.ts) and Microsoft's [profile guidance](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/work-with-user-profiles-in-sharepoint).

## Read list and file version history

```ts
const listVersions = await client.getListItemVersions(siteId, listId, itemId, { maxItems: 100 });
const fileVersions = await client.getDriveItemVersions(driveId, driveItemId, { maxItems: 100 });
```

Reads are uncached and bounded. List versions may include historical fields; drive versions return metadata. Permissions: `Sites.Read.All` for lists, `Files.Read` for work/school files. Retention is site-controlled. Use `/items/{id}/content` for current bytes. See [`Versions.ts`](../examples/Versions.ts), [list versions](https://learn.microsoft.com/en-us/graph/api/listitem-list-versions?view=graph-rest-1.0), and [drive versions](https://learn.microsoft.com/en-us/graph/api/driveitem-list-versions?view=graph-rest-1.0).

## Download or restore a file version

```ts
const bytes = await client.downloadDriveItemVersion(driveId, driveItemId, selectedVersion.id);
// Save or inspect the Blob without putting it in the cache.

if (await confirmDialog("Restore this file version as the current version?")) {
  await client.restoreDriveItemVersion(driveId, driveItemId, selectedVersion.id);
}
```

The download uses Graph's version-content endpoint and returns fresh bytes. Restore creates a new current version and preserves the history. Confirm before restoring. Permissions: `Files.Read` for download and `Files.ReadWrite` for restore. See [`Versions.ts`](../examples/Versions.ts), [download version content](https://learn.microsoft.com/en-us/graph/api/driveitemversion-get-contents?view=graph-rest-1.0), and [restore a version](https://learn.microsoft.com/en-us/graph/api/driveitemversion-restoreversion?view=graph-rest-1.0).

## Sync a document library with delta

```ts
const result = await client.syncDrive(driveId, { maxPages: 100 });
const items = await client.getDriveItems(driveId);
if (!result.complete) {
  // The checkpoint is persisted. Call syncDrive again to continue.
}
```

`syncDrive` stores root-drive items and a tenant/user-scoped cursor, validates continuations, applies deletions, and resumes bounded pages. Delta is current state, not history. Default limit: 100 pages (max 1,000); repeat while `complete` is false. `useDriveItems` serves cache first and syncs on mount/visibility. The fixture's [DriveDeltaPanel](../examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/components/DriveDeltaPanel.tsx) shows inventory, status, file types, and actions. Modified-date order is not event history. See [`DriveDelta.ts`](../examples/DriveDelta.ts) and [Graph delta](https://learn.microsoft.com/en-us/graph/api/driveitem-delta?view=graph-rest-1.0).

## Create, move, and delete drive items

```ts
const folder = await client.createDriveFolder(driveId, parentFolderId, "Reports", { conflictBehavior: "rename" });
const moved = await client.moveDriveItem(driveId, fileId, folder.id, { name: "Quarterly report.docx" });
// Invoke this only after the user confirms the target in the host UI.
await client.deleteDriveItem(driveId, moved.id, { eTag: moved.eTag });
```

Uses authenticated Graph, validates names, applies ETags when available, and updates local drive state. Move is same-drive only; delete uses the recycle bin. Work/school accounts need `Files.ReadWrite`. See [`DriveMutations.ts`](../examples/DriveMutations.ts), [create folder](https://learn.microsoft.com/en-us/graph/api/driveitem-post-children?view=graph-rest-1.0), [move](https://learn.microsoft.com/en-us/graph/api/driveitem-move?view=graph-rest-1.0), and [delete](https://learn.microsoft.com/en-us/graph/api/driveitem-delete?view=graph-rest-1.0).

## Manage the recycle bin

```ts
const page = await services.sharepoint.getRecycleBinItemsPage({ pageSize: 50 });
const item = page.items[0];

if (item && await confirmDialog(`Restore ${item.LeafName} to its original location?`)) {
  await services.sharepoint.restoreRecycleBinItem(item.Id);
}

if (item && await confirmDialog(`Permanently delete ${item.LeafName}?`)) {
  await services.sharepoint.permanentlyDeleteRecycleBinItem(item.Id);
}
```

This uses SharePoint REST, not Graph's SharePoint Embedded recycle-bin API. Reads are paged; choose `scope: "site"` to read or act on the site-collection bin. Restore can fail when the original location has a conflict. Permanent deletion cannot be undone. See [`RecycleBin.ts`](../examples/RecycleBin.ts) and Microsoft's [RecycleBinItem REST methods](https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-visio/jj246549%28v%3Doffice.15%29).

## Read web regional settings

```ts
const settings = await services.sharepoint.getRegionalSettings();
console.log(settings.LocaleId, settings.DateFormat, settings.FirstDayOfWeek);
```

This reads the current web's SharePoint regional settings through the signed-in SPFx REST client. Values are not cached; site usage statistics remain available through raw REST only. See [`RegionalSettings.ts`](../examples/RegionalSettings.ts) and Microsoft's [SharePoint REST API index](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/sharepoint-net-server-csom-jsom-and-rest-api-index).

## Read SharePoint navigation

```ts
const page = await services.sharepoint.getNavigationNodesPage({
  scope: "web", // or "site" to read the site collection's root web
  bar: "quickLaunch", // or "topNavigationBar"
  includeChildren: true,
  pageSize: 50,
});

const links = page.items.flatMap((node) => [node, ...(node.Children ?? [])]);
// If page.nextCursor exists, pass it back to getNavigationNodesPage to continue.
```

Reads use authenticated REST, normalize modern/verbose OData, and expand one child level. Web and site-root scopes differ. Writes and deeper expansion remain raw; inspect the target model first. See [`Navigation.ts`](../examples/Navigation.ts), [REST endpoint guidance](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/determine-sharepoint-rest-service-endpoint-uris?tabs=http), and [CSOM navigation properties](https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-csom/ee545015%28v%3Doffice.15%29).
