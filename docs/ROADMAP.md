# SPFx Fast roadmap

**Last reviewed:** 30 October 2026
**For:** SharePoint Framework 1.24 and later

SPFx Fast builds SharePoint apps with authenticated Graph/REST, caching, React components, a Fluent UI v9 generator, and examples.

## Start here

Choose an [app idea](USE-CASES.md), [form preset](SCENARIOS.md), or [recipe](EXAMPLES.md). See [API coverage](API-COVERAGE.md) and [SPFx setup](SPFX-EXAMPLES.md).

To generate a starter, inspect the signed-in site and list first, then run:

```sh
pnpm dlx --package @spfx-fast/core spfx-fast use-cases
pnpm dlx --package @spfx-fast/core spfx-fast dashboard IssueCenter --site-id <inspected-site-id> --use-case service-outage --list-id <inspected-list-id>
```

Match fields and choices to the inspected list, then build. Starters use Fluent UI React v9.

## What is ready

| Area | Ready |
|---|---|
| Agent setup | `pnpm dlx --package @spfx-fast/core spfx-fast agent init`, static checks, API catalog, agent guide, bounded tool descriptors. |
| Starters | 100 use cases, 24 form presets, single/multi-list Fluent UI v9 dashboards. |
| Auth | SPFx `MSGraphClientV3` and `SPHttpClient`; no token storage. |
| Data | IndexedDB, cached-first reads, bounded paging, list/drive delta, resumable checkpoints, subscriptions, memory fallback. |
| Actions | List CRUD; file upload/download/copy; folder creation; same-drive move; recycle-bin delete; Markdown reads; explicit-scope sharing links; confirmed permission removal and version restore. |
| SharePoint | Typed or raw CAML, `RenderListDataAsStream`, audience targeting, attachments, Search, taxonomy, people, permissions, versions. |
| Examples | React hooks, Fluent UI v9, cache-first drive delta panel, device capture, calendar agenda, floor-plan map, HTML element, SPFx 1.24 web part, Copilot UX. |
| Browser QA | Authenticated recipes, locked profiles, screenshots, traces, and Graph CRUD. Local browser checks cover sample web parts, device/floor-plan UI, file/Markdown, cache failover, upload resume, 28 axe scans, and narrow layouts. See [browser workflow](AGENT-BROWSER-WORKFLOW.md). |
| Validation | SPFx 1.24 dashboard build, synthetic benchmarks, browser/accessibility checks, and tenant probes. Local browser initial render: 28.5 ms. See [tenant validation](TENANT-VALIDATION.md). |

The npm package includes the CLI, guides, catalogs, examples, and fixtures. Check the packed package with `pnpm test:package` or `npm run test:package`.

## What still needs a real tenant

Live checks passed for REST/Graph reads, list/file versions, Person/Lookup writes, 10–128 MiB file round-trips, private images, legacy REST, regional settings, and disposable-list CAML, stream rendering, change tokens, views, audience targeting, and navigation. Synthetic delta passed for 3,000 list items (15 pages, 4.3 s) and 3,000 folders (23 pages, 12 s). These results do not predict business-data performance. See [tenant evidence and screenshots](TENANT-VALIDATION.md#passed). Remaining tenant checks:

- Delta on a real business list and libraries larger than 3,000 items. Synthetic list delta synced 3,000 mixed-field rows over 15 pages; the drive probe synced 3,000 folders over 23 pages. A cross-list token returned the `InvalidToken` marker handled by the reset path.
- A service-driven Graph `410` resync response. The signed-in fixture injected a synthetic drive `410`, verified stale-cache clearing, and resumed one bounded Graph page. The live tenant rejected a malformed same-list cursor with `400 invalidRequest`. Microsoft describes `410` as a service state-recovery response; SharePoint list/drive token invalidation is not a predictable timed expiry. See [list delta](https://learn.microsoft.com/en-us/graph/api/listitem-delta?view=graph-rest-1.0) and [drive delta](https://learn.microsoft.com/en-us/graph/api/driveitem-delta?view=graph-rest-1.0).
- Calendar and device access in their target hosts. The calendar probe returned `MailboxNotEnabledForRESTAPI`; device checks still need their target host.
- Screen-reader use and real browser storage limits.

Tenant results apply only to the tested user, site, list, and operation. SharePoint and Graph control access; cache does not grant it.

Use cases include generic Power Automate patterns; customers configure tenant-specific triggers, connectors, and permissions. The library does not provision flows. SharePoint Markdown is text, not SPFx or MDX.

## Product boundaries

- Browser-based SPFx only; server CSOM and farm APIs need server-side code.
- No Power Automate flow creation or deployment.
- SPFx supplies authentication; credentials stay outside the toolkit.
- IndexedDB is evictable cache, not backup or authorization. OPFS/SQLite is not enabled.
- Graph delta gives current state, not event history or notifications.
- Some APIs are raw-only, legacy, tenant-dependent, or service-hosted. See [API coverage](API-COVERAGE.md).

## Release status

The roadmap and catalog ship in npm. See [release notes](RELEASE.md). Local builds do not prove tenant behavior; run the signed-in checks above.
