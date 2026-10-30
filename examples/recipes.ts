import { createClient, type GraphListItem } from "@spfx-fast/core";
import { createLegacySharePointApi } from "@spfx-fast/core/sharepoint";
import type { SharePointRestTransport } from "@spfx-fast/core/spfx";

export type Client = ReturnType<typeof createClient>;

export async function discoverWorkList(client: Client, hostname: string, sitePath: string, listTitle: string) {
  const site = await client.getSite(hostname, sitePath) as { id: string; webUrl?: string };
  const list = await client.findList(site.id, listTitle) as { id: string; displayName?: string };
  const details = await client.inspectList(site.id, list.id);
  return { site, list, columns: details.columns };
}

export async function readNextAgentPage(client: Client, siteId: string, listId: string, cursor?: string) {
  return client.getItemPage(siteId, listId, { cursor, fields: ["Title", "Status"], pageSize: 50 });
}

export async function readAtMostFiveHundred(client: Client, siteId: string, listId: string) {
  const rows: GraphListItem[] = [];
  for await (const row of client.listItems(siteId, listId, { fields: ["Title", "Status"], maxItems: 500 })) rows.push(row);
  return rows;
}

export function readOpenTasks(client: Client, siteId: string, listId: string) {
  return client.listItems<GraphListItem>(siteId, listId, {
    fields: ["Title", "Status"],
    filter: { field: "Status", operator: "eq", value: "Open" },
    maxItems: 500,
  });
}

export async function keepLocalListFresh(client: Client, siteId: string, listId: string) {
  return client.syncList(siteId, listId, { fields: ["Title", "Status"], maxPages: 100 });
}

export async function createAndUpdateTask(client: Client, siteId: string, listId: string) {
  const created = await client.createItem(siteId, listId, { Title: "Review plan", Status: "Open" });
  return client.updateItem(siteId, listId, created.id, { Status: "In progress" });
}

export async function queryCachedTasks(client: Client, siteId: string, listId: string) {
  return client.queryItems(siteId, listId, {
    equals: { Status: "Open" },
    sort: { field: "Title", direction: "asc" },
    limit: 100,
  });
}

export async function readMarkdownData(client: Client, driveId: string, itemId: string) {
  const result = await client.readDriveItemContent(driveId, itemId, { as: "text" });
  return { markdown: String(result.data), source: result.source, refreshed: result.refresh };
}

export async function uploadReport(client: Client, driveId: string, folderId: string, file: File) {
  return client.uploadDriveItem(driveId, folderId, file.name, file, {
    onProgress: (uploaded, total) => console.info(`Upload ${Math.round(uploaded / total * 100)}%`),
  });
}

export async function readFolderPage(client: Client, driveId: string, folderId: string, cursor?: string) {
  return client.getDriveItemPage(driveId, folderId, { cursor, pageSize: 50 });
}

export async function runLegacyCaml(rest: SharePointRestTransport, listId: string) {
  const sharepoint = createLegacySharePointApi(rest);
  return sharepoint.getItemsByCaml(listId, "<View><RowLimit>100</RowLimit></View>");
}

export async function attachDocument(rest: SharePointRestTransport, listId: string, itemId: number, file: File) {
  const sharepoint = createLegacySharePointApi(rest);
  return sharepoint.addItemAttachment(listId, itemId, file.name, file);
}

export async function readSearchResults(rest: SharePointRestTransport) {
  return createLegacySharePointApi(rest).search('"Project Plan"', {
    rowLimit: 20,
    selectProperties: ["Title", "Path", "LastModifiedTime"],
    sortList: "LastModifiedTime:descending",
  });
}

export async function readRawGraph(client: Client, siteId: string, listId: string) {
  return client.requestGraph(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items?$expand=fields`);
}

export async function readNavigation(rest: SharePointRestTransport, scope: "web" | "site" = "web") {
  return createLegacySharePointApi(rest).getNavigationNodesPage({ scope, bar: "quickLaunch", includeChildren: true, pageSize: 50 });
}
