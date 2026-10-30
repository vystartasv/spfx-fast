import type { createClient } from "./index.js";

type Client = ReturnType<typeof createClient>;

export interface AgentTool {
  name: string;
  description: string;
  inputSchema: Record<string, unknown>;
  requiresConfirmation?: boolean;
  execute(input: Record<string, unknown>): Promise<unknown>;
}

export interface AgentToolContext {
  siteId: string;
  /** Optional allowlist. When present, tools cannot read or write other lists. */
  allowedListIds?: string[];
  maxReadItems?: number;
}

const objectSchema = (properties: Record<string, unknown>, required: string[]) => ({
  type: "object",
  properties,
  required,
  additionalProperties: false,
});

function stringInput(input: Record<string, unknown>, key: string): string {
  const value = input[key];
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${key} must be a non-empty string.`);
  return value;
}

function safeRecord(input: unknown, key: string): Record<string, unknown> {
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError(`${key} must be an object.`);
  return input as Record<string, unknown>;
}

/** Bounded, site-scoped tools for MCP or another agent host to expose. */
export function createAgentTools(client: Client, context: AgentToolContext): AgentTool[] {
  const limit = Math.min(Math.max(context.maxReadItems ?? 50, 1), 100);
  const assertList = (id: string) => {
    if (context.allowedListIds && !context.allowedListIds.includes(id)) throw new Error("This list is outside the tool's allowed list scope.");
  };

  return [
    {
      name: "sharepoint_list_lists",
      description: "List SharePoint lists available in the configured site. Use the returned list ID for inspection and item tools.",
      inputSchema: objectSchema({}, []),
      execute: async () => (await client.getLists(context.siteId)).map((entry) => {
        const list = entry as { id?: string; displayName?: string; name?: string; webUrl?: string; list?: { template?: string } };
        return { id: list.id, title: list.displayName ?? list.name, webUrl: list.webUrl, template: list.list?.template };
      }).filter((list) => !context.allowedListIds || context.allowedListIds.includes(String(list.id))),
    },
    {
      name: "sharepoint_inspect_list",
      description: "Return a SharePoint list's metadata and columns before generating an app or choosing field names.",
      inputSchema: objectSchema({ listId: { type: "string", description: "The list ID returned by sharepoint_list_lists." } }, ["listId"]),
      execute: async (input) => {
        const listId = stringInput(input, "listId");
        assertList(listId);
        return client.inspectList(context.siteId, listId);
      },
    },
    {
      name: "sharepoint_list_items_page",
      description: "Read one bounded page from a SharePoint list. Pass nextCursor from the previous response to continue; never guess a cursor.",
      inputSchema: objectSchema({
        listId: { type: "string" },
        nextCursor: { type: "string", description: "Opaque continuation value returned by the previous call." },
        fields: { type: "array", items: { type: "string" }, maxItems: 20 },
        pageSize: { type: "integer", minimum: 1, maximum: limit, default: Math.min(limit, 50) },
      }, ["listId"]),
      execute: async (input) => {
        const listId = stringInput(input, "listId");
        assertList(listId);
        const pageSize = typeof input.pageSize === "number" ? Math.min(Math.max(Math.floor(input.pageSize), 1), limit) : Math.min(limit, 50);
        return client.getItemPage(context.siteId, listId, {
          cursor: typeof input.nextCursor === "string" ? input.nextCursor : undefined,
          fields: Array.isArray(input.fields) ? input.fields.filter((field): field is string => typeof field === "string").slice(0, 20) : undefined,
          pageSize,
        });
      },
    },
    {
      name: "sharepoint_create_item",
      description: "Create one item in an allowed SharePoint list. Use inspected internal field names and only values supplied or confirmed by the user.",
      inputSchema: objectSchema({ listId: { type: "string" }, fields: { type: "object", additionalProperties: true } }, ["listId", "fields"]),
      execute: async (input) => {
        const listId = stringInput(input, "listId");
        assertList(listId);
        return client.createItem(context.siteId, listId, safeRecord(input.fields, "fields"));
      },
    },
    {
      name: "sharepoint_update_item",
      description: "Update selected fields on one SharePoint item. This is a write operation; confirm the intended changes with the user first.",
      inputSchema: objectSchema({ listId: { type: "string" }, itemId: { type: "string" }, fields: { type: "object", additionalProperties: true } }, ["listId", "itemId", "fields"]),
      requiresConfirmation: true,
      execute: async (input) => {
        const listId = stringInput(input, "listId");
        assertList(listId);
        return client.updateItem(context.siteId, listId, stringInput(input, "itemId"), safeRecord(input.fields, "fields"));
      },
    },
    {
      name: "sharepoint_delete_item",
      description: "Delete one SharePoint item. This is destructive; obtain explicit user confirmation immediately before execution.",
      inputSchema: objectSchema({ listId: { type: "string" }, itemId: { type: "string" }, confirmed: { type: "boolean", const: true } }, ["listId", "itemId", "confirmed"]),
      requiresConfirmation: true,
      execute: async (input) => {
        if (input.confirmed !== true) throw new Error("Explicit confirmation is required before deleting an item.");
        const listId = stringInput(input, "listId");
        assertList(listId);
        await client.deleteItem(context.siteId, listId, stringInput(input, "itemId"));
        return { deleted: true };
      },
    },
  ];
}
