# 100 SharePoint app use cases

A practical starting point for common SharePoint apps and the automations that often go with them. These patterns reuse 24 tested form schemas. Inspect the real SharePoint list before generating an app; update internal field names and choices to match it.

```sh
pnpm dlx --package @spfx-fast/core spfx-fast use-cases
pnpm dlx --package @spfx-fast/core spfx-fast dashboard MyApp --site-id <inspected-site-id> --use-case laptop-request --list-id <inspected-list-id>
```

For npm, replace the `pnpm dlx --package @spfx-fast/core spfx-fast` prefix with `npx --yes --package=@spfx-fast/core -- spfx-fast`.

The app starter runs inside SPFx and uses the signed-in SharePoint client. A flow is a separate Power Automate setup. Sensitive records still need SharePoint permissions; hiding rows in the UI does not secure them.

**10 groups · 100 use cases · 10 reusable flow patterns**

## Flow patterns

| Pattern | What it does |
|---|---|
| **No flow required** | The app handles the work directly. |
| **Approval** | Start approval, update the decision, and notify the requester. |
| **Notification** | Notify an owner or team when a relevant item is created or changed. |
| **Due-date reminder** | Run on a schedule, find due work, and send a reminder. |
| **Scheduled digest** | Group open or overdue work and send a periodic summary. |
| **Document approval** | Approve a file, set its content status, then route it to a folder. |
| **Request routing** | Assign work based on category, location, priority, or amount. |
| **Low-stock alert** | Compare quantity to reorder level and alert purchasing. |
| **Cross-system sync** | Use an approved connector, stable key, and idempotent upsert. |
| **File processing** | Read file metadata or content, classify it, and route the result. |

## Requests and approvals

| Use case | ID | App starter | Suggested automation |
|---|---|---|---|
| Laptop request | `laptop-request` | `equipment-request` | Approval |
| Monitor or phone request | `monitor-request` | `equipment-request` | Approval |
| Software license request | `software-license-request` | `equipment-request` | Approval |
| Office supplies request | `office-supplies-request` | `purchase-request` | Approval |
| Purchase request | `purchase-request-standard` | `purchase-request` | Approval |
| Business travel request | `travel-approval` | `travel-request` | Approval |
| Time off request | `time-off-request` | `leave-request` | Approval; check list permissions |
| Expense reimbursement | `expense-reimbursement` | `expense-claim` | Approval; check list permissions |
| Budget exception request | `budget-exception` | `purchase-request` | Approval |
| Overtime request | `overtime-request` | `leave-request` | Approval; check list permissions |

## People and workforce

| Use case | ID | App starter | Suggested automation |
|---|---|---|---|
| Employee onboarding | `employee-onboarding` | `onboarding-checklist` | Notification; check list permissions |
| Contractor onboarding | `contractor-onboarding` | `onboarding-checklist` | Notification; check list permissions |
| New manager transition | `manager-transition` | `onboarding-checklist` | Due-date reminder; check list permissions |
| Employee transfer checklist | `employee-transfer` | `onboarding-checklist` | Notification; check list permissions |
| Employee offboarding checklist | `employee-offboarding` | `onboarding-checklist` | Request routing; check list permissions |
| Policy acknowledgement campaign | `policy-acknowledgement-campaign` | `policy-acknowledgement` | Due-date reminder; check list permissions |
| Mandatory training assignment | `mandatory-training` | `training-tracker` | Due-date reminder; check list permissions |
| Certification renewal tracker | `certification-renewal` | `training-tracker` | Due-date reminder; check list permissions |
| Mentor and skills directory | `mentor-match` | `skills-directory` | No flow required |
| Employee skills profile | `employee-profile` | `skills-directory` | No flow required; check list permissions |

## IT and service support

| Use case | ID | App starter | Suggested automation |
|---|---|---|---|
| IT help desk | `it-help-desk` | `service-desk` | Request routing |
| System access request | `access-request` | `service-desk` | Approval; check list permissions |
| Hardware repair request | `hardware-repair` | `maintenance-work-order` | Request routing |
| Service outage report | `service-outage` | `issue-reporting` | Notification |
| Security incident intake | `security-incident` | `issue-reporting` | Request routing; check list permissions |
| Account provisioning request | `account-provisioning` | `service-desk` | Approval; check list permissions |
| Permission change request | `permission-change` | `service-desk` | Approval; check list permissions |
| Software license assignment | `license-assignment` | `inventory` | Low-stock alert |
| IT change request | `it-change-request` | `service-desk` | Approval |
| Business application support queue | `application-support` | `service-desk` | Request routing |

## Operations, inspections, and quality

| Use case | ID | App starter | Suggested automation |
|---|---|---|---|
| Vehicle inspection | `vehicle-inspection` | `inspection` | Request routing |
| Workplace safety report | `workplace-safety-report` | `issue-reporting` | Request routing; check list permissions |
| Product quality inspection | `quality-inspection` | `inspection` | Notification |
| Facility inspection | `facility-inspection` | `inspection` | Request routing |
| Corrective action tracker | `corrective-action` | `maintenance-work-order` | Due-date reminder |
| Store opening checklist | `store-opening-checklist` | `inspection` | No flow required |
| Shift handover log | `shift-handover` | `meeting-actions` | Scheduled digest |
| Site visit report | `site-visit-report` | `visitor-log` | Notification |
| Equipment maintenance work order | `equipment-maintenance` | `maintenance-work-order` | Request routing |
| Service level follow-up | `service-level-review` | `service-desk` | Scheduled digest |

## Assets, inventory, and facilities

| Use case | ID | App starter | Suggested automation |
|---|---|---|---|
| Asset register | `asset-register` | `asset-register` | No flow required |
| Equipment checkout | `equipment-checkout` | `asset-register` | Notification |
| Vehicle fleet register | `fleet-register` | `asset-register` | Due-date reminder |
| Warranty expiry tracker | `warranty-tracker` | `asset-register` | Due-date reminder |
| Stock register | `stock-register` | `inventory` | Low-stock alert |
| Low-stock reorder | `reorder-workflow` | `inventory` | Low-stock alert |
| Room booking requests | `room-booking-requests` | `meeting-actions` | Notification |
| Visitor registration | `visitor-registration` | `visitor-log` | Notification; check list permissions |
| Key or access card checkout | `key-card-checkout` | `asset-register` | Notification; check list permissions |
| Facilities service request | `facilities-service-request` | `maintenance-work-order` | Request routing |

## Projects, meetings, and planning

| Use case | ID | App starter | Suggested automation |
|---|---|---|---|
| Project tracker | `project-tracker` | `project-tracker` | Scheduled digest |
| Milestone tracker | `milestone-tracker` | `project-tracker` | Due-date reminder |
| Project issue log | `project-issue-log` | `issue-reporting` | Request routing |
| Project risk register | `project-risk-register` | `risk-register` | Scheduled digest |
| Meeting action log | `meeting-actions` | `meeting-actions` | Due-date reminder |
| Decision register | `decision-register` | `meeting-actions` | Notification |
| Project status reporting | `project-status-report` | `project-tracker` | Scheduled digest |
| Project dependency tracker | `dependency-tracker` | `project-tracker` | Due-date reminder |
| Product or service launch checklist | `launch-checklist` | `onboarding-checklist` | Due-date reminder |
| Project portfolio intake | `portfolio-intake` | `purchase-request` | Approval |

## Communications, events, and ideas

| Use case | ID | App starter | Suggested automation |
|---|---|---|---|
| Company bulletin board | `company-bulletin` | `bulletin-board` | Notification |
| Event planning calendar | `event-planning` | `event-calendar` | Due-date reminder |
| Editorial calendar | `editorial-calendar` | `event-calendar` | Approval |
| Internal campaign tracker | `campaign-tracker` | `event-calendar` | Scheduled digest |
| Employee idea portal | `employee-idea-portal` | `employee-ideas` | Notification |
| Topic Q&A board | `topic-qa` | `topic-qa` | Notification |
| Internal news approval | `news-approval` | `bulletin-board` | Approval |
| Event registration tracker | `event-registration` | `event-calendar` | Notification; check list permissions |
| Volunteer signup | `volunteer-signup` | `event-calendar` | Notification; check list permissions |
| Communications request queue | `communications-request` | `service-desk` | Request routing |

## Documents and knowledge

| Use case | ID | App starter | Suggested automation |
|---|---|---|---|
| Document approval | `document-approval` | `document-review` | Document approval |
| Contract review tracker | `contract-review` | `document-review` | Document approval; check list permissions |
| Policy review cycle | `policy-review` | `document-review` | Due-date reminder |
| Standard operating procedure review | `sop-review` | `document-review` | Due-date reminder |
| Document expiry tracker | `document-expiry` | `document-review` | Due-date reminder |
| Invoice intake and routing | `invoice-intake` | `purchase-request` | File processing; check list permissions |
| Document template request | `document-template-request` | `service-desk` | Request routing |
| Document metadata completion queue | `metadata-completion` | `document-review` | File processing |
| Knowledge article review | `knowledge-article-review` | `document-review` | Approval |
| Records review checklist | `records-review` | `document-review` | Due-date reminder; check list permissions |

## Finance and compliance

| Use case | ID | App starter | Suggested automation |
|---|---|---|---|
| Expense claim | `expense-claim` | `expense-claim` | Approval; check list permissions |
| Invoice approval | `invoice-approval` | `purchase-request` | Document approval; check list permissions |
| Budget tracker | `budget-tracker` | `purchase-request` | Scheduled digest; check list permissions |
| Vendor onboarding | `vendor-onboarding` | `onboarding-checklist` | Approval; check list permissions |
| Compliance attestation | `compliance-attestation` | `policy-acknowledgement` | Due-date reminder; check list permissions |
| Audit evidence tracker | `audit-evidence` | `document-review` | Due-date reminder; check list permissions |
| Policy exception request | `policy-exception` | `risk-register` | Approval; check list permissions |
| Compliance training completion | `compliance-training` | `training-tracker` | Due-date reminder; check list permissions |
| Grant or funding request | `grant-request` | `purchase-request` | Approval; check list permissions |
| Purchase order tracker | `purchase-order-tracker` | `purchase-request` | Cross-system sync; check list permissions |

## Customers, vendors, and partners

| Use case | ID | App starter | Suggested automation |
|---|---|---|---|
| Customer feedback intake | `customer-feedback` | `issue-reporting` | Request routing; check list permissions |
| Customer support queue | `customer-support-queue` | `service-desk` | Request routing; check list permissions |
| Product warranty claim | `warranty-claim` | `maintenance-work-order` | Request routing; check list permissions |
| Partner onboarding | `partner-onboarding` | `onboarding-checklist` | Approval; check list permissions |
| Supplier assessment | `supplier-assessment` | `risk-register` | Approval; check list permissions |
| Delivery issue report | `delivery-issue` | `issue-reporting` | Notification |
| Customer visit log | `customer-visit` | `visitor-log` | Notification; check list permissions |
| Customer service change request | `service-change-request` | `service-desk` | Approval; check list permissions |
| Product quality claim | `product-quality-claim` | `issue-reporting` | Request routing |
| Vendor document review | `vendor-document-review` | `document-review` | Document approval; check list permissions |

## What the starter covers

The 24 schema presets generate search, cached list data, and basic create/edit/delete forms using Fluent UI React v9. They cover text, multiline text, number, date, boolean, and single-choice columns.

Person and lookup columns need explicit mapping. Attachment controls, barcode and camera input, location capture, rich calendars/maps, and offline write queues are not generated by these presets. Use the focused API examples or add the required UI after generating.

## Safe, reliable flows

Use trigger conditions or idempotency markers when a flow edits the item that triggered it. Filter large lists on indexed columns, bound reads, and record reminder state to avoid duplicate notifications. Flows run as their configured connection, which can differ from the signed-in SPFx user. Never put a flow URL, token, or secret in client code.

Flow availability, licensing, permissions, and connector behavior depend on the target Power Platform environment. Review the relevant [scenario notes](SCENARIOS.md) and Microsoft's [SharePoint Power Automate guidance](https://learn.microsoft.com/en-us/power-automate/sharepoint-overview).
