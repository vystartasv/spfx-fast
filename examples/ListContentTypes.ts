import type { GraphContentType, createClient } from "../src/index.js";

type Client = ReturnType<typeof createClient>;

/** Discover list content types before generating schema-specific forms or views. */
export async function listContentTypes(client: Client, siteId: string, listId: string): Promise<GraphContentType[]> {
  return client.getListContentTypes(siteId, listId);
}

// Microsoft Graph v1.0 exposes content types, but not SharePoint list views.
// Use the authenticated same-origin SharePoint REST client for view definitions.
