---
name: spfx-fast
description: Build SharePoint Framework apps quickly with SPFx Fast. Use for list apps, cached data, forms, Graph/REST, and SPFx 1.24+ Copilot UX.
---

# Build SPFx apps with SPFx Fast

Use the installed library before writing Graph/REST plumbing. Install it with the project's package manager, then run `pnpm dlx --package @spfx-fast/core spfx-fast agent init` and `pnpm dlx --package @spfx-fast/core spfx-fast agent doctor`. With npm, use `npx --yes --package=@spfx-fast/core -- spfx-fast`. Read `capabilities.json` for coverage, scopes, limits, and recipe links. npm install does not register this guide; load it through the agent's instruction mechanism.

Tenant-validation commands under `node scripts/...` require a source checkout; they are not included in the npm package.

## Build a dashboard

1. Inspect the signed-in site, list, columns, internal names, and choices. Never guess IDs, URLs, fields, or permissions.
2. Choose with `pnpm dlx --package @spfx-fast/core spfx-fast use-cases` or `pnpm dlx --package @spfx-fast/core spfx-fast scenarios`; generate with `pnpm dlx --package @spfx-fast/core spfx-fast dashboard <Name> --site-id <GraphSiteId> --use-case <id> --list-id <ListId>`. Use `pnpm dlx --package @spfx-fast/core spfx-fast create` for one-list CRUD or `--schema-file` for custom fields/lists. For npm, use the `npx --yes --package=@spfx-fast/core -- spfx-fast` prefix.
3. Match generated fields and choices to the inspected schema. Use Fluent UI React v9 controls inside `FluentProvider`, including `Button`; use Fluent theme tokens for custom styling. Use color by meaning for priority and status. Ask an available design skill to refine layout, responsiveness, theme, and states.
4. Create services once with `createSPFxServices(context, SPHttpClient.configurations.v1)` from `@spfx-fast/core/spfx`. Reuse its `client`, `sharepoint`, and `rest`; never acquire tokens yourself or store credentials.
5. Use `useItems` for cached-first list delta, `useDriveItems` for drive delta, and `useListItems`/`client.listItems` with indexed filters and bounds for large lists. `ListView` supports paging and custom cells. Select only needed fields.
6. For agent reads, use `getItemPage`; pass back the opaque `nextCursor` and keep results bounded.
7. Show proposed writes and get user confirmation for updates, deletes, sharing, and permission changes. SharePoint/Graph enforce access; cache does not grant it.
8. Use authenticated raw REST/Graph when a typed helper is missing. Keep legacy response parsing local; see [API coverage](../docs/API-COVERAGE.md).

## Specialized recipes

| Need | Start here | Limits to preserve |
|---|---|---|
| CAML, Search, audience targeting, attachments | `examples/SharePointRestPlayground.tsx`, `docs/EXAMPLES.md` | Use the `createSPFxServices` REST transport. |
| Recycle bin | `getRecycleBinItemsPage`, `restoreRecycleBinItem`, `permanentlyDeleteRecycleBinItem` | Confirm permanent deletion. |
| Taxonomy | `getTermStoreGroups`, `getTermStoreSets`, `getTermSetChildren`, `getTermChildren`; `examples/TaxonomyPickerData.ts` | `TermStore.Read.All` needs admin consent; page one level at a time. |
| List permissions | `examples/ListPermissions.ts` | List scope only; does not resolve parent inheritance, group membership, links, or item scopes. Confirm the exact change. |
| Site users/groups | `examples/SitePeople.ts` | Hosting web only; not Entra search or effective access. Follow opaque cursors. |
| User profiles | `examples/UserProfiles.ts` | Uncached, visibility-dependent SharePoint profile reads; use Graph for directory search. |
| Versions | `examples/Versions.ts` | Uncached and bounded. Confirm restore; use `Files.ReadWrite` for drive restore. |
| Navigation/views | `examples/Navigation.ts`, `examples/ListViews.ts` | Child expansion is one level. Treat `ViewQuery` as untrusted CAML; writes remain raw. |
| Drive share/copy | `examples/DriveItemPermissions.ts`, `examples/DriveCopy.ts` | Set sharing scope explicitly. `replace` deletes destination history; copies inherit destination permissions and omit metadata. |
| Calendar/device inputs | `examples/CalendarAgenda.tsx`, `examples/DeviceCapture.tsx` | Calendar is uncached; needs `Calendars.ReadBasic` and an active REST mailbox. Device access needs user action and host permissions. |
| Floor plans | `examples/DriveFloorPlanMap.tsx`, `examples/FloorPlanMap.tsx` | Use inspected drive/item IDs and normalized 0–100 coordinates; no public map service. |
| Markdown/HTML | `useMarkdownFile`, `MarkdownContent`, `defineSPFxFastListElement`; [SPFx examples](../docs/SPFX-EXAMPLES.md) | Markdown HTML stays inert; never evaluate SharePoint Markdown as MDX. Standalone HTML has no SPFx auth. |
| App-catalog deploy | [Deployment guide](../docs/AGENT-SHAREPOINT-DEPLOYMENT.md) | Confirm catalog/app/component IDs and installed state from SharePoint; install test apps site-scoped. |
| Copilot UX | `examples/fixtures/spfx-fast-copilot` | Validate supplied site/list; bound reads; keep writes confirmed. Requires a Copilot license. |
| Browser validation | [Browser workflow](../docs/AGENT-BROWSER-WORKFLOW.md) | Use a dedicated signed-in profile, published test page, synthetic data, and one job per profile. |

## Common recipe

```tsx
import type { createClient } from "@spfx-fast/core";
import { ListView, useItems } from "@spfx-fast/core/react";

type Client = ReturnType<typeof createClient>;

function TaskList({ client, siteId, listId }: { client: Client; siteId: string; listId: string }) {
  const { items, loading, error } = useItems(client, siteId, listId, { fields: ["Title", "Status"] });
  return <ListView items={items} loading={loading} error={error} fields={[{ name: "Title", title: "Task" }]} />;
}
```

Read [examples](../docs/EXAMPLES.md) for copyable recipes and [scenarios](../docs/SCENARIOS.md) for app/flow patterns. Power Automate patterns are guidance; the package does not create flows. Local builds do not prove tenant access. For checks and limits, see [roadmap](../docs/ROADMAP.md) and [tenant validation](../docs/TENANT-VALIDATION.md).
