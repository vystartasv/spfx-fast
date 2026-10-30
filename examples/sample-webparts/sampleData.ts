import type { GraphListItem } from "@spfx-fast/core";

export const sampleSiteId = "spfx-fast-demo-site";
export const sampleListIds = { requests: "sample-requests", tasks: "sample-tasks", approvals: "sample-approvals" };
export const sampleDriveId = "sample-drive";
export const sampleMarkdownItemId = "team-handbook";

export const sampleRequests: GraphListItem[] = [
  { id: "REQ-1001", fields: { Title: "Laptop for new starter", Requester: "Alex Morgan", Priority: "High", Status: "Pending", Amount: 1250 } },
  { id: "REQ-1002", fields: { Title: "Ergonomic chair for Finance", Requester: "Sam Patel", Priority: "Normal", Status: "Approved", Amount: 320 } },
  { id: "REQ-1003", fields: { Title: "Replace meeting room display", Requester: "Taylor Reed", Priority: "High", Status: "In review", Amount: 890 } },
];

export const sampleTasks: GraphListItem[] = [
  { id: "TASK-201", fields: { Title: "Review laptop supplier quote", Owner: "Alex Morgan", Priority: "High", Status: "Open" } },
  { id: "TASK-202", fields: { Title: "Confirm chair delivery date", Owner: "Sam Patel", Priority: "Normal", Status: "Open" } },
  { id: "TASK-203", fields: { Title: "Update West Wing asset register", Owner: "Taylor Reed", Priority: "Normal", Status: "Complete" } },
  { id: "TASK-204", fields: { Title: "Prepare meeting room display handover", Owner: "Alex Morgan", Priority: "High", Status: "Open" } },
];

export const sampleApprovals: GraphListItem[] = [
  { id: "APR-301", eTag: '"APR-301-v1"', fields: { Title: "Laptop for new starter", Requester: "Alex Morgan", Amount: 1250, Status: "Pending" } },
  { id: "APR-302", eTag: '"APR-302-v1"', fields: { Title: "Replace meeting room display", Requester: "Taylor Reed", Amount: 890, Status: "Pending" } },
  { id: "APR-303", eTag: '"APR-303-v1"', fields: { Title: "Ergonomic chair for Finance", Requester: "Sam Patel", Amount: 320, Status: "Approved" } },
];

export const sampleMarkdown = `**North Campus operations** · Equipment, meeting rooms, and building support.\n\n### Request equipment\n\nSubmit a request with the asset type, delivery location, and estimated cost. A manager reviews purchases over **£500**.\n\n### Meeting rooms\n\nRoom 04 seats eight and has a wall display. Return shared equipment to the cabinet after use.\n\n### Service desk\n\n| Day | Hours | Contact |\n| --- | --- | --- |\n| Monday–Friday | 08:00–18:00 | Extension 4400 |\n| Weekends | Emergency support | Facilities duty manager |\n\nFor urgent building issues, call the duty manager before submitting a standard request.\n`;
