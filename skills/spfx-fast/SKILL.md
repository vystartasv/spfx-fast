---
name: spfx-fast
description: 'Build SharePoint Framework apps with SPFx Fast. Use for SharePoint list forms, request trackers, dashboards, and authenticated Graph/REST features in SPFx 1.24+.'
---

# SPFx specialist skill

Use this skill for SharePoint Framework work: understand the host, package, tenant APIs, and deployment lifecycle before changing code. Keep existing PnP and Microsoft clients when they fit.

## Route the task

1. Inspect the repo's SPFx version, package manager, current dependencies, and existing patterns. Read `@spfx-fast/core`'s `agent/SKILL.md` and `agent/capabilities.json` before using its APIs.
2. For component selection, Fluent UI and SharePoint UI patterns, route to [`@spfx-fast/sharepoint-ui`](https://github.com/vystartasv/sharepoint-ui)'s skill and catalog. Keep that focused catalog separate from this SPFx workflow.
3. For authenticated browser debugging, route to [`@spfx-fast/playwright`](../packages/playwright/README.md) and the [debugging playbook](../docs/DEBUGGING.md).
4. For other concerns, use the narrow reference below; do not load every guide for every task.

## Build workflow

1. Inspect the signed-in site context, lists, columns, internal names, and permissions. Never guess tenant IDs, fields, URLs, or access.
2. Choose a starter with `pnpm dlx --package @spfx-fast/core spfx-fast use-cases` or `pnpm dlx --package @spfx-fast/core spfx-fast scenarios`. With npm, use `npx --yes --package=@spfx-fast/core -- spfx-fast`. Generate only after inspecting the target schema:

   ```sh
   pnpm dlx --package @spfx-fast/core spfx-fast dashboard <Name> --site-id <inspected-site-id> --use-case <id> --list-id <inspected-list-id>
   ```

   Replace the pnpm prefix with `npx --yes --package=@spfx-fast/core -- spfx-fast` to use npm.
3. Run `spfx-fast agent init` and `spfx-fast agent doctor` through the same package runner. Doctor checks local setup, not tenant permissions.
4. Use `createSPFxServices(context, SPHttpClient.configurations.v1)` for authenticated SharePoint calls. Do not acquire tokens independently or store credentials.
5. Preserve permission checks, bounded reads, loading/error states, accessible controls, and confirmation before writes or deletes.

## Common requests

| Request | Use case |
|---|---|
| Laptop or purchase request | `laptop-request` |
| IT help desk | `it-help-desk` |
| Employee onboarding | `employee-onboarding` |
| Facility inspection | `facility-inspection` |
| Asset register | `asset-register` |
| Stock reorder tracker | `reorder-workflow` |

## References by task

| Task | Read |
|---|---|
| Common failures and field-tested gotchas | [Debugging playbook](../../packages/knowledge/docs/DEBUGGING.md) |
| API, permissions, and limits | [API coverage](../../docs/API-COVERAGE.md) |
| Deployment and stale bundle diagnosis | [Deployment guide](../../docs/AGENT-SHAREPOINT-DEPLOYMENT.md) |
| Signed-in browser checks | [Browser workflow](../../docs/AGENT-BROWSER-WORKFLOW.md) |
| Reusable app patterns | [Examples](../../docs/EXAMPLES.md), [SPFx examples](../../docs/SPFX-EXAMPLES.md) |
| Scenario and starter selection | [Use cases](../../docs/USE-CASES.md), [scenarios](../../docs/SCENARIOS.md) |
| Share a reusable finding | [Contribution guide](../../CONTRIBUTING.md#knowledge-contributions) |

When a helper is missing, use authenticated raw REST or Graph through supplied services. Check API coverage first. Keep special response parsing local and small; do not replace authentication, caching, paging, or delta sync without a concrete need.

## References

- [Package guide](https://github.com/vystartasv/spfx-fast/blob/main/agent/SKILL.md)
- [Use cases](https://github.com/vystartasv/spfx-fast/blob/main/docs/USE-CASES.md)
- [Examples](https://github.com/vystartasv/spfx-fast/blob/main/docs/EXAMPLES.md)
- [SPFx examples](https://github.com/vystartasv/spfx-fast/blob/main/docs/SPFX-EXAMPLES.md)
- [API coverage](https://github.com/vystartasv/spfx-fast/blob/main/docs/API-COVERAGE.md)
- [Deployment guide](https://github.com/vystartasv/spfx-fast/blob/main/docs/AGENT-SHAREPOINT-DEPLOYMENT.md)
