# SharePoint app and automation scenarios

Generate SPFx starters from common app and Power Automate patterns. SPFx Fast does not create or run flows.

## Generate a starter

1. Reuse the signed-in SPFx client. Inspect the site, list, and internal column names.
2. Choose a use case with `pnpm dlx --package @spfx-fast/core spfx-fast use-cases` or a preset with `pnpm dlx --package @spfx-fast/core spfx-fast scenarios`.
3. Generate a Fluent UI v9 app:

```sh
pnpm dlx --package @spfx-fast/core spfx-fast dashboard IssueCenter --site-id <inspected-GraphSiteId> --use-case service-outage --list-id <inspected-ListId>
```

For npm, replace `pnpm dlx --package @spfx-fast/core spfx-fast` with `npx --yes --package=@spfx-fast/core -- spfx-fast`.

4. Match field names and choices to the list. Use `--schema-file` for custom fields.
5. Build any suggested workflow separately in Power Automate or another service.

Presets: [`examples/scenarios.json`](../examples/scenarios.json). Use cases: [`examples/use-cases.json`](../examples/use-cases.json) and [catalog](USE-CASES.md). Forms support text, multiline, number, date, boolean, choice, multi-choice, Person, Lookup, and attachments. Multi-choice uses string arrays; Person/Lookup need authenticated options. See the [form recipe](EXAMPLES.md#use-person-and-lookup-form-fields). [`DeviceCapture.tsx`](../examples/DeviceCapture.tsx) shows camera, barcode, and location inputs; validate host permissions. [`CalendarAgenda.tsx`](../examples/CalendarAgenda.tsx) shows a bounded Outlook agenda. [`FloorPlanMap.tsx`](../examples/FloorPlanMap.tsx) maps rooms/assets without a map service. Validate tenant permissions and coordinates. Offline writes are out of scope.

## Common app starters

| User request | Preset | App behavior | Common automation to pair |
|---|---|---|---|
| Request a laptop, monitor, or phone | `equipment-request` | Create, edit, search, and track requests | Approval by manager; notify requester; mark ordered/fulfilled |
| Request time off | `leave-request` | Capture dates, leave type, reason, and status | Route approval; notify employee and update shared calendar through an approved connector |
| Buy supplies or services | `purchase-request` | Track supplier, amount, need-by date, and status | Amount-based approval; receipt or rejection notice |
| Report a workplace, IT, or safety issue | `issue-reporting` | Intake, triage, assignment, and resolution | Teams/email alert; create Planner task; escalate critical severity |
| Inspect a site, vehicle, or product | `inspection` | Record location, result, findings, and follow-up | Escalate failed inspections; remind owners about corrective actions |
| Track company equipment | `asset-register` | Search assets by tag, category, owner, location, and lifecycle | Assignment notice; warranty and maintenance reminders |
| Manage stock or supplies | `inventory` | Update quantities, reorder levels, and locations | Low-stock alert; purchase request; scheduled reconciliation |
| Onboard an employee | `onboarding-checklist` | Assign and track onboarding tasks | Create checklist rows from a starter template; due reminders; notify owners |
| Track projects and milestones | `project-tracker` | Browse milestones, owners, dates, progress, and risk status | Weekly digest; overdue reminder; escalate at-risk items |
| Plan events and published content | `event-calendar` | Track dates, channels, owners, and editorial status | Review approval; publish-date reminder; announce release |
| Collect employee ideas | `employee-ideas` | Submit, categorize, review, and track implementation | Acknowledge submission; notify campaign reviewers |
| Run a service desk | `service-desk` | Triage by category/priority and track resolution | Route by category; create task; send status and resolution updates |
| Review controlled documents | `document-review` | Track owner, reviewer, due date, and approval status | File-created approval; set content approval status; move approved files |
| Schedule maintenance | `maintenance-work-order` | Record assets, priority, owner, resolution, and due date | Route urgent work; remind assignees; notify on completion |
| Manage organizational risks | `risk-register` | Rate likelihood/impact and track mitigation | Escalate critical risks; scheduled review digest |
| Track meeting actions | `meeting-actions` | Capture decisions, owners, due dates, and state | Send assignments; recurring overdue-action digest |
| Collect policy acknowledgements | `policy-acknowledgement` | Record policy version, assignment, and response | Assign to a group; remind employees who haven't responded |
| Publish internal announcements | `bulletin-board` | Manage announcement copy, audience notes, dates, and review state | Notify the intended audience at publish time; remind about expiry |
| Host topic discussions and Q&A | `topic-qa` | Submit questions, assign answers, and track open questions | Notify topic owners about unanswered questions |
| Share colleague skills | `skills-directory` | Maintain a searchable, opt-in skills and interests list | Use profile APIs for live directory data where suitable |
| Assign and track training | `training-tracker` | Track courses, learners, due dates, completion, and scores | Remind learners; report overdue assignments |
| Request business travel | `travel-request` | Capture destination, dates, purpose, estimate, and approval state | Route approval and notify the traveler |
| Claim expenses | `expense-claim` | Track amount, cost center, receipt link, and reimbursement state | Route by amount and notify after review |
| Register visitors | `visitor-log` | Track visits, host, purpose, and check-in/out | Notify the host and send reception a daily digest |

Patterns are not fixed schemas. Inspect columns and index fields used for large-list filters. Person/Lookup fields are not plain text in Graph writes. Secure sensitive records with SharePoint permissions; app filters are not security. Flows use their configured identity, which may differ from the app user.

## Common flow patterns

Microsoft documents SharePoint triggers for item/file create, update, delete, and selected item/file actions, plus actions for file/list operations. Start with these flow recipes when the requirement includes:

| Flow pattern | Trigger and action outline | Design note |
|---|---|---|
| Submit-and-approve | On item create → start/wait approval → branch on outcome → update status → notify requester | Choose single, sequential, or “everyone must approve”; use separate approval state fields. |
| Document approval and routing | On file create → get metadata/ETag → approval → set content approval status → move approved file | Preserve the ETag required by the content approval action. |
| Change notification | On item/file create or modified → condition on relevant fields → email or Teams message | Use “Get changes” or compare selected fields. Add a trigger condition or idempotency marker to prevent a flow that updates its trigger item from looping. |
| Due-date reminder | Recurrence → query indexed due/status fields → send reminder → record reminder state | Recurrence belongs in Power Automate; browser tabs are not reliable schedulers. |
| Daily/weekly digest | Recurrence → retrieve bounded/filtered changes → group by owner/team → send digest | Bound reads and index filter columns; batch large digests. |
| Manual selected-item action | User selects a list item/file → validate status → perform approved action → write result | Do not expose anonymous flow HTTP URLs in client code. |
| File processing | File created → inspect metadata/content → classify or extract → write metadata → route to folder | AI Builder/Syntex and some connectors need licenses or premium capacity. |
| Cross-service synchronization | Connector trigger or recurrence → map stable key → upsert SharePoint item → record outcome | Use idempotency keys and explicit conflict handling to prevent duplicate rows. |
| Attachment/document generation | Form or item submitted → generate document → save file and URL → notify submitter | Document generation and connectors can be preview or license gated. |
| Lifecycle and archive | Status/retention date reached → approval or retention action → move/archive → audit result | Use Purview retention controls for records; do not substitute custom client-side deletion. |

Flow definitions depend on environment connection references, connector licensing, run-only users, and tenant policy. This repository gives portable blueprints rather than a tenant-independent flow package ready to import. Browser apps use signed-in SPFx Graph/REST clients; cloud flows use configured connector identities and their own authorization. Never pass access tokens or secrets between them.

## Research basis and boundaries

Microsoft's Power Apps samples cover inspections, issues, ideas, bulletins, onboarding, milestones, and directories. SharePoint's Power Automate docs cover approvals, files, lists, reminders, triggers, and actions. Sources: [Power Apps from lists](https://learn.microsoft.com/en-us/power-apps/maker/canvas-apps/app-from-sharepoint), [Teams samples](https://learn.microsoft.com/en-us/power-apps/teams/use-sample-apps), [Power Automate scenarios](https://learn.microsoft.com/en-us/power-automate/sharepoint-overview), [connector reference](https://learn.microsoft.com/en-us/sharepoint/dev/business-apps/power-automate/sharepoint-connector-actions-triggers), [document approval](https://learn.microsoft.com/en-us/sharepoint/dev/business-apps/power-automate/guidance/require-doc-approval).
