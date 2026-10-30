# Tenant validation

**Checked:** 30 October 2026 · **Host:** signed-in SPFx 1.24 site

## Passed

| Batch | Result |
|---|---|
| Package and page | Site-scoped package 1.0.0.62 is valid, enabled, and installed. Published page loads current bundle `spfx-fast-dashboard-web-part_71724e0b094ae9f1bcd8.js`. |
| Signed-in reads | 10/11 passed: REST, user, Graph discovery/schema/content types, views, storage, library, paging, and cache. Calendar returned `MailboxNotEnabledForRESTAPI`. |
| Regional settings | `/_api/web/RegionalSettings` returned HTTP 200 with locale/date/calendar/week fields. |
| Graph `$batch` | Five signed-in site, list, column, content type, and item reads passed; all individual subrequest statuses were successful. |
| List CRUD | Created, read, updated, read, and deleted one synthetic item. |
| Person/Lookup Graph write | Created two temporary lists; tested Graph create/read/update/delete with Person and Lookup fields, then deleted both lists. |
| Drive delta | One page completed; the test library was empty. |
| Synthetic delta scaling | List delta: 250, 1,000, 2,000, 2,500, and 3,000 mixed-field rows passed in 2, 5, 10, 13, and 15 pages; 3,000 synced in 4.28 s. Drive delta: 250, 2,000, and 3,000 folders passed in 2, 18, and 23 pages (8.75 s and 11.99 s for the larger runs). Seeding 3,000 rows/folders took 113/378 s; cleanup passed. Synthetic data is not a business-list benchmark. [250 list](assets/screenshots/tenant-list-delta-paging.png) · [250 drive](assets/screenshots/tenant-drive-delta-paging.png) · [1,000 list](assets/screenshots/tenant-list-1000-workload.png) · [2,000 list/drive](assets/screenshots/tenant-delta-2026-10-30/list-delta-2000.png) · [2,500 list](assets/screenshots/tenant-delta-2500/list-delta-2500.png) · [3,000 list](assets/screenshots/tenant-delta-3000/list-delta-3000.png) · [3,000 drive](assets/screenshots/tenant-drive-delta-3000/drive-delta-3000.png). |
| File round-trip | Uploaded, downloaded, byte-verified, and deleted synthetic files of 10, 50, 100, and 128 MiB + 1 byte using Graph upload sessions. [50 MiB](assets/screenshots/tenant-upload-50mb.png) · [100 MiB](assets/screenshots/tenant-upload-100mb.png) · [128 MiB](assets/screenshots/tenant-upload-128mb.png). |
| Private floor-plan image | Fixture 1.0.0.50 loaded its current bundle, uploaded a synthetic PNG to the signed-in user's drive, rendered the private file through `DriveFloorPlanMap` at 800×600, and deleted it. [Screenshot](assets/screenshots/tenant-private-floor-plan.png). |
| Interrupted upload recovery | Injected a lost chunk response, resumed at Graph's real session range, verified a 20 MiB + 1 byte download, then cleaned up (6.9 s). This was not an offline test. [Screenshot](assets/screenshots/tenant-upload-recovery.png). |
| Actual offline upload recovery | Browser went offline after the first 10 MiB chunk; reconnect resumed the 128 MiB + 1 byte upload, verified bytes, and cleaned up (23.1 s). [Screenshot](assets/screenshots/tenant-upload-offline-recovery-0.2.4.png). |
| Post-window upload resume | After six failed status reads, the page resumed the same 128 MiB + 1 byte session. SHA-256/name matched; cleanup and UI reset passed (36.8 s). [Screenshot](assets/screenshots/tenant-upload-post-window-recovery-0.2.4.png). |
| Lost final upload response | Dropped the final HTTP 201 after a 20 MiB + 1 byte upload. Name, size, and bytes matched; cleanup passed (6.0 s). [Screenshot](assets/screenshots/tenant-upload-final-response-0.2.4.png). |
| Existing-file replace after lost response | Fixture 1.0.0.56 dropped the final HTTP 201 after replace; same item ID and bytes matched, then cleanup passed (6.0 s). [Screenshot](assets/screenshots/tenant-upload-replace-final-response-0.2.4.png). |
| Graph rename after lost response | Fixture 1.0.0.60 verified Graph's renamed 20 MiB + 1 byte file by bounded path probes and byte comparison, then deleted both files. [Screenshot](assets/screenshots/tenant-upload-rename-recovery-0.2.4.png). |
| List/file version history | On disposable versioned objects, Graph returned three list versions and two file versions. Prior/current file bytes matched; both objects were deleted. [Screenshot](assets/screenshots/tenant-version-history.png). |
| Legacy SharePoint REST batch | 17 checks passed: web Quick Launch (6 nodes, 4 children), web top navigation (empty), site-root Quick Launch (6 nodes), CAML, `RenderListDataAsStream`, three-page change-token XML, a live `InvalidToken` response, `GetChanges`, views, and audience targeting field add/remove. Both temporary lists and their items were deleted. |
| Malformed Graph delta cursor | A disposable list returned `400 invalidRequest`; the probe verified the client preserved its cached row and cursor. The temporary list was deleted. This does not simulate Graph's `410` resync response. [Screenshot](assets/screenshots/tenant-delta-invalid-cursor.png). |
| Simulated drive delta `410` recovery | The fixture injected one synthetic `410` for an isolated stale cursor, verified the cache was cleared, then completed one bounded read through the signed-in Graph client (201 items, incomplete checkpoint). No tenant data was written; this does not prove a service-driven `410` will occur. [Screenshot](assets/screenshots/tenant-drive-410-recovery.png). |
| Copilot package deployment | SPFx 1.24 Copilot package v1.0.0.0 is valid, enabled, and available to all sites. `Sites.Read.All` was already approved organization-wide. Add to Teams returned HTTP 400 with a remote 403; the test account has SharePoint Plan 1 only, and the agent is absent from the Microsoft 365 admin center registry. |

Tests used one signed-in user and synthetic content. They do not establish access for other users, sites, or tenants. No user identity, tenant URL, tokens, cookies, or item content is retained.

## Still unverified

- Delta paging against a real business list and libraries larger than 3,000 items. The synthetic schema and one test account do not establish performance for real list designs or other tenants. The 3,000-folder drive sync took 12.0 seconds, while creation took 378 seconds under throttling. The 2,000-folder run created in 106 seconds, synced in 8.7 seconds, and cleaned up in 1 second; setup-time variance is significant.
- A service-driven Graph `410` resync. The fixture injected a synthetic drive `410` and verified cache clearing plus a bounded authenticated restart; a malformed same-list cursor returned `400 invalidRequest` and preserved its cached row and cursor. Microsoft documents these `410` responses as service state recovery; they are not a predictable timed expiry.
- Calendar and device access in their target hosts. The calendar request returned `MailboxNotEnabledForRESTAPI`; this tenant account cannot verify the calendar flow until its mailbox is enabled for Graph REST.
- Screen-reader behavior, browser quota policy, and tenant performance.

## Lessons captured

- Route REST through the supplied site path; `/_api` may target the tenant root. Tenant and site-collection catalogs are separate.
- After deployment, verify `AvailableApps` versions, exact component ID, and the bundle loaded by the published page. `GetAppInstanceById` can report a false negative. Increment `solution.version` before rebuilds; same-version assets can stay stale.
- Find pages in **Site Pages**; never assume `Page.aspx`. Use synthetic records and verify cleanup. Adding children changes the parent folder ETag; refresh it before deletion.
- Legacy SharePoint writes may need verbose OData v3 headers even when SPFx defaults to v4. Read back fields after `CreateFieldAsXml`.
- Graph current-file bytes use `/items/{id}/content`; `/versions/{id}/content` is for older versions. Version timestamps can tie; use version IDs to order.
- `Calendars.ReadBasic` does not provision an Exchange mailbox. Check for `MailboxNotEnabledForRESTAPI` before debugging permissions.
- Graph delta malformed cursors return `400 invalidRequest`; preserve cached rows/cursors. `410` recovery is different and may provide a fresh `Location`; follow only same-resource links. This tenant has not produced a service-driven `410`.
- Upload recovery needs distinct tests: simulated lost response, actual offline interruption, and lost final response. Keep `DriveUploadInterruptedError` in page memory and resume after reconnecting; it cannot survive reload.
- For `replace`, create a Graph upload session by existing item ID. In this tenant, a duplicate `item.name` caused `400 invalidRequest`; omitting it worked. Treat this as tenant-tested behavior.
- For `rename` after a lost response, use bounded candidate names and verify parent, size, and full bytes. Do not enumerate the folder.
- Separate synthetic data seeding, delta, query, and cleanup timings. Do not present synthetic results as business-list performance.
