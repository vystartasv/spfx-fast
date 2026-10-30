# SPFx Fast Data and Agent Build Experience

## G. Goal

Help agents build useful SharePoint Framework apps quickly, with a reactive local cache and fast, reliable SharePoint data access.

## C. Constraints and limitations

- C1: Target SPFx 1.24 and later; validate Fluent UI React v9, React, and Copilot UX integration against the exact SPFx release used.
- C2: In SPFx, use the host-provided `SPHttpClient` for SharePoint REST and `MSGraphClientV3` for Graph authentication; the core must not own credentials.
- C3: Browser persistence is best-effort. Cached data can be evicted or cleared, and first-time data still requires a network request. Interrupted Graph upload recovery holds the file and private upload-session URL in memory; it cannot survive a page reload.
- C4: SharePoint/Graph throttling and permissions remain authoritative. Honor `Retry-After`; do not imply local cache bypasses authorization.
- C5: Keep the core framework-independent. Provide React and SPFx host adapters separately.
- C6: Auto-pagination must be bounded and cancellable. Agent-facing operations must have explicit result limits.
- C7: Graph delta and SharePoint change-token APIs are incremental sync sources, not complete item history. Notification delivery requires a callback service.
- C8: Cached file content is scoped to the signed-in tenant/user cache, expires after 5 minutes, and is persisted only when it is at most 1 MB by default. Markdown is untrusted document data.
- C9: Synthetic and local-browser benchmarks do not prove signed-in tenant performance.
- C10: Copilot UX is generally available from SPFx 1.24.0. It requires a Microsoft 365 Copilot add-on license, supports browser/desktop Copilot clients, and is not supported on mobile at GA; verify current tenant/client requirements before deployment.

## I. Interfaces

- `I1`: Framework-independent TypeScript data client and local store.
- `I2`: SPFx transport adapter for `SPHttpClient` and `MSGraphClientV3`.
- `I3`: React hooks and standard list/detail/form views.
- `I4`: Agent capability catalog, starter instructions, templates, CLI, and host-neutral tool descriptors; hosts register and expose tools.
- `I5`: SharePoint REST, Microsoft Graph, and explicitly scoped legacy CAML/CSOM-oriented adapters.
- `I6`: CLI commands for agent setup/diagnostics and Fluent UI React v9 app/dashboard scaffolding.

## T. User stories and acceptance tasks

| ID | Status | User story and observable acceptance | Track |
|---|---|---|---|
| T1 | implemented | As an app developer, I can read paged Graph list data through a framework-independent client, persist pages in IndexedDB, and get cached data before background refresh. | Core read path |
| T2 | implemented | As an app developer, I can subscribe to synchronized local items and see cache and sync updates without refetching the same entities. | Reactive local data |
| T3 | implemented | As an app developer, I can synchronize lists incrementally with Graph delta and recover from invalid cursors. | Sync |
| T4 | implemented | As an app developer, I can create, update, and delete items optimistically with ETag conflict recovery. | Writes |
| T5 | implemented | As an agent, I can scaffold a working list app from a small declarative specification and supported templates. | Agent build surface |
| T6 | implemented | As an agent, I can inspect machine-readable capabilities, schemas, permissions, limits, and escape hatches, and generate project guidance with static diagnostics. | Agent build surface |
| T7 | implemented | As an SPFx developer, I can use host-authenticated REST and Graph transports without the core owning credentials, or compose them through `createSPFxServices`. | SPFx adapter |
| T8 | implemented | As an app developer, I can use bounded automatic pagination, Fluent UI React v9 list/detail/form views, and bounded local queries over synced rows. Exact queries with an explicit limit up to 100 cache for one second and invalidate on list changes. | App primitives |
| T9 | implemented | As an advanced SharePoint developer, I can reach legacy REST/CAML operations and use raw REST or Graph for uncovered browser-capable operations; unsupported CSOM-only work is classified. | API coverage |
| T10 | implemented | As an operator, I can inspect cache, local query, network, sync, mutation, retry, and storage metrics, check storage status, and configure retry and sync policy. | Performance and operations |
| T11 | implemented | As an app developer, I can read Graph drive-item metadata and content through tenant/user-scoped cached helpers and React hooks, render untrusted Markdown with the sanitized GFM component, and keep current root-drive state through bounded, resumable delta sync with persisted cursors and invalidated folder pages. | Files and content |
| T12 | implemented | As an agent, I can generate a single-list app from one of 24 business scenario presets or build a custom one-list/multi-list Fluent UI React v9 dashboard from inspected list IDs, fields, types, required flags, and choices; the packed-artifact gate compiles scenario, one-list, and multi-list apps in SPFx 1.24. | Dashboard starters |
| T13 | implemented | As an operator, I can run fake IndexedDB and real Chromium benchmarks at 10,000, 50,000, and 100,000 rows, compare uncached and warm queries, and inspect per-list first-result/network timing and storage estimates. The browser gate checks cached-first reads, cross-instance invalidation, and memory fallback for generic and simulated `QuotaExceededError`. Actual browser quota policy and signed-in tenant performance remain unverified. | Performance validation |
| T14 | implemented | As a developer, I can create folders, move items within a drive, delete items to the recycle bin with ETag checks when available, upload files through resumable Graph sessions, resume an interrupted upload in the current page, and render a bounded list custom element with safely formatted rich text inside a host that supplies the authenticated client. | Files and interactive surfaces |
| T15 | implemented | As an SPFx developer, I can build separate Fluent UI v9 web-part and Copilot UX fixtures against SPFx 1.24.0 using `createSPFxServices`; the web-part fixture pages a document library and renders sanitized Markdown previews, and the Copilot fixture reads a bounded list through the signed-in user. Dashboard package 1.0.0.60 passed Graph `rename` recovery after a lost final response. Package 1.0.0.56 passed existing-file replacement recovery after a lost final response. Earlier versions passed private-image and post-window upload checks. Earlier checks cover CRUD, Person/Lookup writes, and paged list/drive delta. Copilot runtime validation is excluded as requested. | Host validation and examples |
| T16 | implemented | As an SPFx developer, I can manage a list item's attachments through the authenticated SharePoint REST panel and compose it with create/edit forms using the optional `ItemForm.attachments` slot. | SharePoint REST examples |
| T17 | implemented | As an SPFx developer, I can see cache-first drive inventory and resumable delta status, file-type distribution, latest-modified items, and open/download/continue-sync actions. Latest-modified ordering is not presented as delta event history. | Files and interactive surfaces |
