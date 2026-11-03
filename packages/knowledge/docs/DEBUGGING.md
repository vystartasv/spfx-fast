# SPFx and SharePoint debugging playbook

Start with observable evidence. Record the SPFx version, host, page state, request status, component ID, package version, and exact loaded bundle when relevant. Do not put tenant URLs, IDs, user data, tokens, screenshots, or traces in public findings.

| Symptom | Check | Likely explanation / action | Evidence scope |
|---|---|---|---|
| REST succeeds on one web but fails elsewhere | Inspect the request path and target web. Compare `/sites/team/_api/...` with `/_api/...`. | Route REST through the intended site path. `/_api` may resolve at the tenant root. | Tenant-validated in one signed-in site; check target web and permissions. |
| App install/deploy returns 200 but the web part is absent or old | Read target site's `AvailableApps`; compare `InstalledVersion` and `AppCatalogVersion`. Then inspect the published page's component ID and hashed bundle. | HTTP 200 can mean request accepted only. Catalog metadata can lag the bundle served. Increment `solution.version` before rebuilding; verify catalog and runtime state separately. | Tenant-tested. `GetAppInstanceById` returned a false negative for an app present in Site Contents. |
| Install reports an existing instance but read-back cannot find it | Inspect Site Contents and app-instance state. | Stop; repeating Install can reproduce the conflict. Do not retry lifecycle calls blindly. | Tenant-tested. |
| Instance remains `Initialized` (status 9) after install | Read the exact instance status and job result. | Avoid repeating the catalog lifecycle request. A guarded `AppInstance.Install()` retry exists in the repo helper and only runs for one matching status-9 instance; stop if it does not reach Installed (status 5). | Tenant-tested for the helper's guarded path; still inspect each response. |
| Component works in Workbench/editor but not on published page | Confirm page is published, exact manifest title/component ID, and loaded release bundle. Find pages in Site Pages rather than assuming a filename. | Draft pages and Workbench can differ from published SharePoint host behavior. | Browser workflow guidance; validate on the actual target page. |
| `MailboxNotEnabledForRESTAPI` appears despite calendar scope | Check whether the signed-in account has an active Exchange mailbox enabled for REST. | Treat this as mailbox/host setup before changing Graph consent. A permission scope alone does not provision a mailbox. | Tenant-validated failure; verify account and target host. |
| Graph delta returns 400 for a malformed cursor | Preserve current cache and cursor; inspect the actual response and cursor source. | `400 invalidRequest` is not the same as a service `410` resync. Do not clear useful state based on a malformed cursor. | Malformed cursor reproduced in a disposable list; service-driven 410 remains unverified in the recorded tenant run. |
| File version content is wrong | Check whether the request is `/items/{id}/content` or `/versions/{versionId}/content`. | Current bytes use `/items/{id}/content`; `/versions/{id}/content` addresses prior versions. Use version IDs to order tied timestamps. | Tenant-validated. |
| Upload resumes after network failure but fails after reload | Check whether the in-memory interrupted upload object is still available. | The interrupted upload retains the Blob and private session URL in page memory; it cannot survive reload. Never serialize or expose the session URL. | Tenant-validated for bounded upload scenarios; not a durable-resume guarantee. |
| Playwright locator or click fails on SharePoint chrome | Inspect screenshot, trace, console, failed requests, page state, and accessibility tree before changing selectors. Close duplicate toolboxes and handle the debug-manifest prompt after navigation. | SharePoint chrome, draft state, and prompts can change the accessible tree or block interactions. Prefer role/label locators; after two similar failures inspect DOM and network evidence. | Workflow guidance; host UI can vary. |
| A local check passes but tenant behavior differs | Compare host, site, identity, scopes, package bundle, and real response. | Static checks, Workbench, one user's success, and synthetic workloads do not prove other tenants or users. Reproduce in the target scope and state the limit. | General validation rule. |

## Debugging loop

1. Reproduce once and capture the smallest useful evidence locally.
2. Locate the failing boundary: build/compiler, package manifest, page host, network/auth, SharePoint service, or UI state.
3. Compare expected and observed values, including status/body and exact asset/version. Redact evidence before sharing.
4. Change one cause, then rerun the narrow check and one end-to-end check in the real host when needed.
5. Record the verified scope and known limits. Treat hypotheses as unverified until reproduced.

Sources: repository [tenant validation](https://github.com/vystartasv/spfx-fast/blob/main/docs/TENANT-VALIDATION.md), [deployment guide](https://github.com/vystartasv/spfx-fast/blob/main/docs/AGENT-SHAREPOINT-DEPLOYMENT.md), and [browser workflow](https://github.com/vystartasv/spfx-fast/blob/main/docs/AGENT-BROWSER-WORKFLOW.md). Findings inherit those documents' validation limits.
