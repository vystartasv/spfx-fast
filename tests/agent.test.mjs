import assert from "node:assert/strict";
import test from "node:test";
import { createAgentTools } from "../dist/agent.js";

function createFakeClient() {
  const calls = [];
  return {
    calls,
    client: {
      getLists: async (siteId) => { calls.push(["getLists", siteId]); return [
        { id: "allowed-list", displayName: "Requests", webUrl: "/lists/requests" },
        { id: "blocked-list", displayName: "Private", webUrl: "/lists/private" },
      ]; },
      inspectList: async (siteId, listId) => { calls.push(["inspectList", siteId, listId]); return { id: listId }; },
      getItemPage: async (siteId, listId, options) => { calls.push(["getItemPage", siteId, listId, options]); return { items: [], nextCursor: undefined }; },
      createItem: async (siteId, listId, fields) => { calls.push(["createItem", siteId, listId, fields]); return { id: "created" }; },
      updateItem: async (siteId, listId, itemId, fields) => { calls.push(["updateItem", siteId, listId, itemId, fields]); return { id: itemId }; },
      deleteItem: async (siteId, listId, itemId) => { calls.push(["deleteItem", siteId, listId, itemId]); },
    },
  };
}

test("agent tools enforce the configured site, list allowlist, and page bound", async () => {
  const fake = createFakeClient();
  const tools = createAgentTools(fake.client, { siteId: "site-1", allowedListIds: ["allowed-list"], maxReadItems: 500 });
  const byName = Object.fromEntries(tools.map((tool) => [tool.name, tool]));

  assert.deepEqual(await byName.sharepoint_list_lists.execute({}), [
    { id: "allowed-list", title: "Requests", webUrl: "/lists/requests", template: undefined },
  ]);
  await byName.sharepoint_inspect_list.execute({ listId: "allowed-list" });
  await assert.rejects(byName.sharepoint_inspect_list.execute({ listId: "blocked-list" }), /outside the tool's allowed list scope/);

  const fields = Array.from({ length: 25 }, (_, index) => `Field${index}`);
  await byName.sharepoint_list_items_page.execute({ listId: "allowed-list", nextCursor: "opaque", fields, pageSize: 500 });
  assert.deepEqual(fake.calls.at(-1), ["getItemPage", "site-1", "allowed-list", {
    cursor: "opaque", fields: fields.slice(0, 20), pageSize: 100,
  }]);
  await assert.rejects(byName.sharepoint_list_items_page.execute({ listId: "blocked-list" }), /outside the tool's allowed list scope/);
  assert.equal(byName.sharepoint_list_items_page.inputSchema.properties.pageSize.maximum, 100);
});

test("agent write tools mark confirmation and reject unconfirmed deletes", async () => {
  const fake = createFakeClient();
  const tools = Object.fromEntries(createAgentTools(fake.client, { siteId: "site-1", allowedListIds: ["allowed-list"] }).map((tool) => [tool.name, tool]));

  assert.equal(tools.sharepoint_update_item.requiresConfirmation, true);
  assert.equal(tools.sharepoint_delete_item.requiresConfirmation, true);
  await tools.sharepoint_update_item.execute({ listId: "allowed-list", itemId: "42", fields: { Status: "Done" } });
  await assert.rejects(tools.sharepoint_delete_item.execute({ listId: "allowed-list", itemId: "42" }), /Explicit confirmation is required/);
  assert.equal(fake.calls.some(([name]) => name === "deleteItem"), false);
  assert.deepEqual(await tools.sharepoint_delete_item.execute({ listId: "allowed-list", itemId: "42", confirmed: true }), { deleted: true });
  assert.equal(fake.calls.at(-1)[0], "deleteItem");
  await assert.rejects(tools.sharepoint_update_item.execute({ listId: "blocked-list", itemId: "42", fields: {} }), /outside the tool's allowed list scope/);
});
