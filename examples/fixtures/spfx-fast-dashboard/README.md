# SPFx Fast authenticated dashboard fixture

Runnable SPFx 1.24.0 web part using React 18, Fluent UI v9, authenticated Graph/REST, cached lists, and a paged file panel.

## Requirements

- Node.js 22.14+, 24.x, or 26.x.
- SharePoint development site and App Catalog.
- Optional Tasks/Projects lists (`Title`, `Status`) and access to the default library.
- Admin approval for Graph `Sites.ReadWrite.All` and `Calendars.ReadBasic`.

## Install and run

```sh
cd <repository-root>
pnpm run build
cd examples/fixtures/spfx-fast-dashboard
pnpm install --frozen-lockfile
pnpm start -- --nobrowser
```

Open a modern site page and load the manifest with the [SPFx Debug Toolbar](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/debug-toolbar). Add **SPFx Fast Dashboard** and optionally enter Tasks/Projects list IDs. Use **Work** for the list and file panels, **Calendar** for the signed-in user's bounded Outlook agenda, or **Floor plan** to upload, display, and remove a synthetic private image through Graph. Calendar data is not cached or written. The default library appears automatically; set a folder ID to show a subfolder. The drive panel shows current delta-sync status, file-type distribution, and latest-modified files with open/download actions. The library panel pages a folder, offers upload/download actions, and renders safe Markdown previews; other files link to SharePoint. Delta is current inventory, not full edit history.

Using npm: run `npm ci` and `npm start -- --nobrowser` in this fixture.

The dashboard shows cache-first and network timings, local-query timings, and storage estimates. **Run signed-in checks** runs eleven bounded reads across SharePoint REST, Graph, cache, and browser storage. **Test list delta paging** creates 3,000 synthetic work tickets with choice, number, text, and date fields; it measures Graph delta pages and cold/warm IndexedDB queries, then deletes the list and cache. **Test drive delta paging** resumes one-page syncs over 2,000 synthetic folders, creating Graph batches with concurrency three. **Test malformed delta cursor** seeds a bad cursor in a disposable list, confirms Graph rejects it without clearing cached rows, and deletes the list; this does not simulate token expiry. **Test drive delta (one page)** syncs at most 200 library entries and saves a checkpoint to the local cache; repeat it to resume. **Test list CRUD** creates, reads, updates, and deletes one synthetic item. **Test Person/Lookup writes** creates two disposable lists, tests Graph create/read/update/delete, then removes both lists. The 10, 50, 100, and 128 MiB upload checks upload, verify, and delete temporary files. Write and paging probes ask before running; manually remove `SPFxFastRefProbe*`, `SPFxFastWorkload_*`, `SPFxFastDriveDelta_*`, `SPFxFastRoundTrip-*`, or `SPFxFastDeltaRecovery_*` data if cleanup fails. Reports omit identities and file contents.

Deploy and approve `Sites.ReadWrite.All` and `Calendars.ReadBasic` in SharePoint admin center before using Graph list, drive, or calendar calls. The list-write scope is needed for the optional CRUD check; tailor the app manifest if you only need read access. SPFx provides authentication. Validate upload-session/CORS behavior in the tenant. Checks prove only the current user's access; they are not an ACL audit or production benchmark.

## Build

```sh
pnpm build
```

From the repository root, `pnpm test:spfx-fixture` builds the library and fixture, then packages the `.sppkg`. Tenant permissions and page rendering need live SharePoint; local builds do not verify them.

## Implementation notes

- `SpfxFastDashboardWebPart.ts`: creates authenticated services and resolves the Graph site.
- `WorkDashboard.tsx`: selectable work and Outlook calendar views; calendar reads only after the user selects it.
- `PrivateFloorPlanProbe.tsx`: confirm-before-write live check that uploads a synthetic private image, renders it through `DriveFloorPlanMap`, then removes it.
- `CalendarAgenda.tsx`: bounded Graph calendar view, shared by the standalone example and SPFx fixture.
- `DocumentLibraryPanel.tsx`: bounded paging, sanitized Markdown preview, upload progress and same-page resume, fresh downloads.
- `PerformanceProbe.tsx`: median/p95 local query timings and storage estimates. Diagnostic only; bounds limit results, not rows scanned.
- `TenantValidationPanel.tsx`: bounded Graph/REST checks and confirmed disposable CRUD, Person/Lookup, version-history, 2,000-item IndexedDB-backed work-queue and drive delta benchmarks, 10/50/100/128 MiB upload, and interrupted-session recovery probes; reports no identity or file content.
- Cache uses tenant/user-scoped IndexedDB with memory fallback.
