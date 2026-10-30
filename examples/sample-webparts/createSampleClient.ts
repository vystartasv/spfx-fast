import { createClient, createMemoryStore, type GraphListItem } from "@spfx-fast/core";
import { sampleApprovals, sampleDriveId, sampleListIds, sampleMarkdown, sampleMarkdownItemId, sampleRequests, sampleTasks } from "./sampleData.js";

/** Local-only synthetic Graph adapter for demos and browser tests. It never calls SharePoint. */
export function createSampleClient() {
  const lists = new Map<string, GraphListItem[]>([
    [sampleListIds.requests, sampleRequests.map((item) => structuredClone(item))],
    [sampleListIds.tasks, sampleTasks.map((item) => structuredClone(item))],
    [sampleListIds.approvals, sampleApprovals.map((item) => structuredClone(item))],
  ]);
  let nextId = 1;
  const get = async (value: string) => {
    const url = new URL(value);
    const match = url.pathname.match(/\/lists\/([^/]+)\/items(?:\/delta)?$/);
    if (!match) throw new Error("The sample client supports only its sample lists.");
    const items = lists.get(decodeURIComponent(match[1]));
    if (!items) throw new Error("Unknown sample list.");
    let selected = items;
    const filter = url.searchParams.get("$filter");
    const matchFilter = filter?.match(/^fields\/Status eq '((?:[^']|'')*)'$/);
    if (matchFilter) selected = items.filter((item) => item.fields?.Status === matchFilter[1].replaceAll("''", "'"));
    return url.pathname.endsWith("/delta")
      ? { value: selected.map((item) => structuredClone(item)), "@odata.deltaLink": `${url.origin}${url.pathname}?token=sample` }
      : { value: selected.map((item) => structuredClone(item)) };
  };
  const request = async (value: string, options: { method?: string; body?: unknown; responseType?: string } = {}) => {
    const url = new URL(value);
    const method = options.method ?? "GET";
    if (url.pathname === `/v1.0/drives/${sampleDriveId}/items/${sampleMarkdownItemId}/content`) return sampleMarkdown;
    const match = url.pathname.match(/\/lists\/([^/]+)\/items(?:\/([^/]+))?$/);
    if (!match) throw new Error("The sample client supports only its sample Graph routes.");
    const items = lists.get(decodeURIComponent(match[1]));
    if (!items) throw new Error("Unknown sample list.");
    const itemId = match[2] ? decodeURIComponent(match[2]) : undefined;
    const body = (options.body ?? {}) as { fields?: Record<string, unknown> } & Record<string, unknown>;
    if (method === "POST" && !itemId) {
      const created = { id: `SAMPLE-${nextId++}`, eTag: '"1"', fields: { ...(body.fields ?? {}) } };
      items.unshift(created);
      return structuredClone(created);
    }
    if (method === "PATCH" && itemId) {
      const current = items.find((item) => item.id === itemId);
      if (!current) throw new Error("Sample item not found.");
      const updated = { ...current, eTag: '"2"', fields: { ...current.fields, ...(body.fields ?? {}) } };
      items[items.indexOf(current)] = updated;
      return structuredClone(updated);
    }
    if (method === "DELETE" && itemId) {
      const index = items.findIndex((item) => item.id === itemId);
      if (index >= 0) items.splice(index, 1);
      return undefined;
    }
    throw new Error(`Unsupported sample request: ${method}.`);
  };
  return createClient({ scope: "spfx-fast:local-sample", store: createMemoryStore(), get, request });
}
