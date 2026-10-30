# Compiled SPFx Fast examples

See the [scenario catalog](../docs/SCENARIOS.md) and [`scenarios.json`](scenarios.json) for app starters and Power Automate patterns.

`pnpm build` / `npm run build` type-checks examples. npm ships their TypeScript source. Examples use the caller's SPFx client; builds make no tenant calls.

- `WorkDashboard.tsx`: multi-list Fluent UI React v9 dashboard.
- `FilteredTasks.tsx`: bounded server-filtered list view using the `useListItems` React hook.
- `MarkdownFilePanel.tsx`: cached-first, sanitized Markdown panel using `useMarkdownFile` inside the SPFx React host.
- `SharePointRestPlayground.tsx`: interactive SharePoint Search, CAML, attachment, and audience-targeting examples using the signed-in SPFx REST transport.
- `ItemAttachmentsPanel.tsx`: authenticated list-item attachment list, upload/replace, download, and confirmed delete workflow.
- `DriveItemPermissions.ts`: typed, uncached Graph read of sharing permissions visible to the current SPFx caller.
- `ListContentTypes.ts`: typed Graph v1.0 list content-type discovery; notes that views use SharePoint REST.
- `TaxonomyPickerData.ts`: read-only, level-by-level Graph term store discovery for managed metadata pickers.
- `ListPermissions.ts`: inspect one list's role bindings and scope, then add/remove principal roles after explicit user confirmation.
- `ListViews.ts`: bounded SharePoint REST view discovery and lazy loading of a selected view's internal field names.
- `SitePeople.ts`: authenticated current-user lookup and bounded SharePoint site-user, group, and group-member pages for pickers.
- `UserProfiles.ts`: read-only SharePoint PeopleManager lookups for the current user, a visible user's profile, or one profile property; profile values are not cached.
- `Versions.ts`: bounded, uncached Graph version metadata for list items and document-library files.
- `DriveDelta.ts`: resumable Graph drive-delta synchronization with persisted drive items and cursor state.
- The dashboard fixture's `DriveDeltaPanel.tsx`: cache-first drive inventory, file-type chart, sync status, latest-modified files, and file actions.
- `DriveMutations.ts`: authenticated folder creation, same-drive moves, and deletion to the SharePoint recycle bin with ETag support.
- `PersonLookupForm.tsx`: authenticated people and lookup options for `ItemForm`.
- `DeviceCapture.tsx`: user-triggered photo/barcode and geolocation inputs with manual fallbacks.
- `CalendarAgenda.tsx`: a bounded, cache-free agenda from the signed-in user's Outlook calendar.
- `FloorPlanMap.tsx`: accessible map markers and location list over a local or authenticated floor-plan asset; consumes normalized list coordinates and makes no external map calls.
- `DriveFloorPlanMap.tsx`: loads a private PNG, JPEG, or WebP floor-plan from Graph through the cached authenticated file hook and cleans up its object URL.
- `sample-webparts/`: request dashboard, task queue, Markdown panel, approval inbox, and local synthetic data. See [screenshots](../docs/SCREENSHOTS.md).
- `fixtures/spfx-fast-copilot/`: SPFx 1.24.0 GA Copilot UX declarative agent that uses `createSPFxServices`, reads an exact list with a 50-item cap, and uses Fluent UI v9.
- `fixtures/spfx-fast-dashboard/`: complete SPFx 1.24.0 fixture that builds a `.sppkg` against the workspace package, uses authenticated `createSPFxServices`, renders selectable work/calendar views in Fluent UI v9, pages a document-library folder, renders safe Markdown, and measures local query latency in the browser.

Build the dashboard with `pnpm test:spfx-fixture` (`npm run test:spfx-fixture`); build/package Copilot with `pnpm test:copilot-fixture` (`npm run test:copilot-fixture`).
- `RecycleBin.ts`: bounded recycle-bin reads and confirmed restore/permanent-delete calls.
- `RegionalSettings.ts`: typed, uncached regional settings read for the hosting web.
- `Navigation.ts`: bounded Quick Launch and top-navigation reads, with optional one-level child expansion.
- `recipes.ts`: typed examples for discovery, bounded and server-filtered list/folder paging, delta sync, CRUD, local queries, Markdown-as-data, resumable uploads, list attachments, CAML, search, and raw Graph.

In SharePoint, create services with `createSPFxServices(context, SPHttpClient.configurations.v1)` and pass `client`. Request/admin-approve only needed Graph scopes. See `docs/EXAMPLES.md` for API details.
