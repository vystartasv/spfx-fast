---
name: spfx-fast
description: Specialist workflow for building, debugging, testing, and deploying SharePoint Framework applications. Load for SPFx projects and SharePoint-hosted extensions.
---

# SPFx specialist skill

Use this one skill for SPFx work. Inspect the project's SPFx version, package manager, current dependencies, host, and existing patterns first. Read the installed `@spfx-fast/core` agent guide and capabilities before using its APIs. Never guess tenant IDs, site URLs, list fields, permissions, or deployment state.

## Task routing

- Component choice, Fluent UI, accessibility, RTL, and SharePoint UI patterns: use the [`@spfx-fast/sharepoint-ui` skill and catalog](https://github.com/vystartasv/sharepoint-ui).
- Authenticated browser checks: use [`@spfx-fast/playwright`](https://github.com/vystartasv/spfx-fast/tree/main/packages/playwright).
- Troubleshooting: load `docs/DEBUGGING.md` from this package.
- API, deployment, examples, and scenarios: follow the focused references in the SPFx Fast repository's [`docs/LEARNINGS.md`](https://github.com/vystartasv/spfx-fast/blob/main/docs/LEARNINGS.md).

## Build safely

1. Inspect site/list schema and permissions before generating code.
2. Prefer existing project libraries and `createSPFxServices(context, SPHttpClient.configurations.v1)` for authenticated calls; do not acquire tokens independently or store credentials.
3. Keep reads bounded, preserve permission checks, expose loading and error states, and use accessible controls.
4. Confirm before writes, deletes, sharing, or permission changes. Use synthetic data and a dedicated test page for validation.
5. Distinguish local/static checks from proof in the target tenant. Record applicability and limitations in reusable findings.

## Share a discovery

Use `spfx-fast-knowledge learn init`, fill in the reusable behavior and evidence level without tenant-specific details, then run `learn validate` and `learn patch`. Review the generated patch manually. Keep URLs, IDs, identities, content, credentials, browser profiles, screenshots, and traces out of contributions. See the repository's [contribution guide](https://github.com/vystartasv/spfx-fast/blob/main/CONTRIBUTING.md#knowledge-contributions).
