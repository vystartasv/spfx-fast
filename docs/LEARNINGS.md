# SharePoint learning index

Use the narrowest source that answers the question. Treat tenant validation as evidence for the tested user, site, and operation; do not generalize it without reproduction.

| Need | Start here |
|---|---|
| Build or adapt an SPFx app | [Agent guide](../agent/SKILL.md), [use cases](USE-CASES.md), [scenarios](SCENARIOS.md) |
| Debug SPFx and SharePoint failures | [SPFx knowledge package](../packages/knowledge/README.md), [debugging playbook](../packages/knowledge/docs/DEBUGGING.md) |
| Find an API, permission, or known limit | [API coverage](API-COVERAGE.md) |
| Copy a working implementation | [Examples](EXAMPLES.md), [SPFx examples](SPFX-EXAMPLES.md) |
| Check actual tenant behavior | [Tenant validation](TENANT-VALIDATION.md), [browser workflow](AGENT-BROWSER-WORKFLOW.md) |
| Test SharePoint UI and publish a reusable finding | [`@spfx-fast/playwright`](../packages/playwright/README.md#promote-a-sharepoint-ui-discovery) |
| Draft and submit a reusable knowledge patch | [`@spfx-fast/knowledge`](../packages/knowledge/README.md), [contribution guide](../CONTRIBUTING.md#knowledge-contributions) |
| Verify deployment and app-catalog lifecycle | [Deployment guide](AGENT-SHAREPOINT-DEPLOYMENT.md) |

## Connected projects

`spfx-fast` is the discovery and routing hub. These projects remain independently versioned; this index points agents to their current source of truth instead of copying their code or docs.

| Knowledge or tool | Canonical source | Connection to SPFx Fast |
|---|---|---|
| SharePoint UI components, tokens, RTL, accessibility, and component APIs | [`sharepoint-ui`](https://github.com/vystartasv/sharepoint-ui) and its [agent skill](https://github.com/vystartasv/sharepoint-ui/blob/main/docs-site/public/agents/SKILL.md) | Use `@spfx-fast/sharepoint-ui` for UI; keep its component catalog and docs authoritative. [`sharepoint-ui-docs`](https://github.com/vystartasv/sharepoint-ui-docs) is the published static site output, not a second source of truth. |
| Manual SPFx host and preview behavior | [`spfx-fieldcheck`](https://github.com/vystartasv/spfx-fieldcheck) | Local-only observations are a source for candidate learnings. **Integration gap:** it exports Markdown today; adapt it to the `spfx-fast-learning` format before importing reports into the package. |
| App data, Graph/REST patterns, examples, and agent setup | [`spfx-fast`](https://github.com/vystartasv/spfx-fast) | Canonical home for reusable app patterns, API knowledge, and the base agent skill. |
| SPFx agent skill, debugging playbook, and contribution patch CLI | [`@spfx-fast/knowledge`](../packages/knowledge/README.md) | Portable SPFx specialist skill; debugging/gotchas are the first-class reference. Its local patch generator has no upload path. |
| Generate a constrained SPFx feature | [`m365-feature-assembler`](https://github.com/vystartasv/m365-feature-assembler) | Keep its deterministic generator standalone; link it from agent task routing when its input contract fits. |
| Scaffold and verify a new project | [`spfx-factory`](https://github.com/vystartasv/spfx-factory) | Use for new-solution generation and evidence-gated local checks. |
| Check one existing SPFx project | [`spfx-assurance`](https://github.com/vystartasv/spfx-assurance) | Independent offline project checks; reports are local evidence, not tenant proof. |
| Inventory multiple local projects | [`spfx-estate-doctor`](https://github.com/vystartasv/spfx-estate-doctor) | Bounded, read-only estate inventory. |
| Compare declared estate state with project files | [`m365-extension-control`](https://github.com/vystartasv/m365-extension-control) | Release evidence and preflight; does not deploy or approve tenant changes. |
| SPFx implementation and upgrade guidance | [`spfx-dev-skills`](https://github.com/SharePoint/spfx-dev-skills) and [PnP SharePoint Skills](https://github.com/pnp/sharepoint-skills) | Treat as external skill sources. Preserve attribution and check their current guidance; do not fork their docs into the base skill. |
| Recurring fixes and agent troubleshooting | [`hermes-harness`](https://github.com/vystartasv/hermes-harness) | Candidate miner for recurring patterns. Review and sanitize every finding before it becomes public guidance. |
| Logs, screenshots, media, and test-report inspection | [`agent-efficiency-tools`](https://github.com/vystartasv/agent-efficiency-tools) | Optional local diagnostics for the agent; raw artifacts stay in the user's workspace. |
| Evaluate whether agent changes improve results | [`evalcase`](https://github.com/vystartasv/evalcase) and [`agent-observer`](https://github.com/vystartasv/agent-observer) | Use synthetic task outcomes and local traces to measure quality and speed without collecting user telemetry. |

`spfx-data-kit` has useful framework-independent transport work but overlaps the runtime in `@spfx-fast/core`. Compare capabilities before adopting any code; do not create two competing defaults for SharePoint data access. Keep private work logs and tenant evidence private, and publish only reviewed, generalized findings.

## Promote a new discovery

1. Create a generalized finding with `npx --yes @spfx-fast/knowledge learn init`.
2. Record symptoms, diagnostic evidence, cause, fix, applicability, reproduction, limitations, and honest validation level. Keep tenant URLs, IDs, names, content, profiles, authentication state, screenshots, and traces out of the draft.
3. Run `learn validate`, then `learn patch`. Review the patch and source evidence yourself; the check catches common identifiers but cannot guarantee anonymity.
4. Reproduce the behavior when possible. Keep unverified observations labeled as such.
5. Apply the patch in a repository checkout and propose a reviewed pull request. Promote the finding into the durable guide, API/scenario document, or Playwright recipe that best fits. Add an offline fixture assertion where practical.
6. Keep supporting notes specific about versions, host, validation scope, and known exceptions. Never turn a one-tenant result into a universal claim.

Draft creation and validation are local-only. There is no telemetry, automatic upload, or account identity collected by this workflow. Repository contributions may still expose contributor identity through GitHub; the published learning artifact itself should contain no author field.
