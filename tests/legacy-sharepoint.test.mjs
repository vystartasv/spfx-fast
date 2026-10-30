import test from "node:test";
import assert from "node:assert/strict";

import { createLegacySharePointApi, syncListWithChangeToken } from "../dist/sharepoint.js";

const listId = "12345678-1234-4234-8234-123456789abc";

test("audience targeting uses the standard TargetTo field instead of a non-writable list property", async () => {
  const requests = [];
  let fields = [];
  const rest = {
    siteUrl: "https://contoso.sharepoint.com/sites/team",
    async get(path) { requests.push({ method: "GET", path }); return { value: fields }; },
    async request(path, options) {
      requests.push({ method: options.method, path, options });
      if (path.endsWith("/CreateFieldAsXml")) {
        fields = [{ Id: "61cbb965-1e04-4273-b658-eedaa662f48d", InternalName: "Target_x0020_Audiences", TypeAsString: "TargetTo" }];
        return fields[0];
      }
      if (options.method === "DELETE") fields = [];
      return undefined;
    },
    raw: async () => { throw new Error("unexpected raw request"); },
  };
  const api = createLegacySharePointApi(rest);

  await api.setAudienceTargeting(listId, true);
  assert.equal(requests[1].path, `/_api/web/lists(guid'${listId}')/fields/CreateFieldAsXml`);
  assert.match(requests[1].options.body.parameters.SchemaXml, /Type="TargetTo"/);
  assert.equal(requests[1].options.body.parameters.__metadata.type, "SP.XmlSchemaFieldCreationInformation");
  await api.setAudienceTargeting(listId, false);
  assert.equal(requests.at(-1).method, "DELETE");
  assert.match(requests.at(-1).path, /fields\(guid'61cbb965-1e04-4273-b658-eedaa662f48d'\)$/);
});

test("GetChanges serializes the SharePoint DeleteObject flag and Int64 FetchLimit", async () => {
  let captured;
  const api = createLegacySharePointApi({
    siteUrl: "https://contoso.sharepoint.com/sites/team",
    get: async () => ({}),
    raw: async () => { throw new Error("unexpected raw request"); },
    request: async (path, options) => { captured = { path, options }; return { value: [] }; },
  });
  await api.getChanges(listId, { Item: true, DeleteObject: true, FetchLimit: 100 });
  assert.equal(captured.path, `/_api/web/lists(guid'${listId}')/GetChanges`);
  assert.equal(captured.options.body.query.DeleteObject, true);
  assert.equal(captured.options.body.query.FetchLimit, "100");
  assert.throws(() => api.getChanges(listId, { FetchLimit: 0 }), /positive integer/);
});

test("change-token sync sends list-scoped query properties and saves the returned token", async () => {
  const previousParser = globalThis.DOMParser;
  globalThis.DOMParser = class {
    parseFromString() {
      const row = {
        attributes: [{ name: "ows_ID", value: "7" }, { name: "ows_Title", value: "Synthetic" }],
        getAttribute: (name) => name.toLowerCase() === "ows_id" ? "7" : null,
      };
      return {
        querySelector: () => null,
        getElementsByTagName(name) {
          if (name === "z:row") return [row];
          if (name === "Changes") return [{ getAttribute: (attribute) => attribute === "LastChangeToken" ? "token-next" : null }];
          return [];
        },
      };
    }
  };
  const calls = [];
  const store = {
    async getCursor() { return "token-before"; },
    async setCursor(_scope, _site, key, token) { calls.push({ type: "cursor", key, token }); },
    async putItems(_scope, _site, list, items) { calls.push({ type: "items", list, items }); },
    async deleteItems() {},
    async clearCursor() {},
    async clearItems() {},
  };
  const rest = {
    async raw(path, options) {
      calls.push({ type: "request", path, options });
      return { text: async () => "<rs:data><z:row ows_ID='7' /><rs:response><Changes LastChangeToken='token-next' /></rs:response></rs:data>" };
    },
  };
  try {
    const result = await syncListWithChangeToken({ rest, store, scope: "scope", siteId: "site", listId, fields: ["Title"], rowLimit: 10 });
    const request = calls.find((call) => call.type === "request");
    const query = request.options.body.query;
    assert.equal(request.path, `/_api/web/lists(guid'${listId}')/GetListItemChangesSinceToken`);
    assert.equal(query.ChangeToken, "token-before");
    assert.equal(query.ViewName, "");
    assert.equal(query.Query, "<Query />");
    assert.equal("ListName" in query, false);
    assert.equal(query.RowLimit, "10");
    assert.deepEqual(result, { pages: 1, changed: 1, token: "token-next" });
    assert.ok(calls.some((call) => call.type === "cursor" && call.token === "token-next"));
  } finally {
    if (previousParser === undefined) delete globalThis.DOMParser;
    else globalThis.DOMParser = previousParser;
  }
});

test("change-token sync clears an invalid token and retries a full refresh", async () => {
  const previousParser = globalThis.DOMParser;
  globalThis.DOMParser = class {
    parseFromString(xml) {
      const invalid = xml === "invalid";
      const row = { attributes: [{ name: "ows_ID", value: "9" }, { name: "ows_Title", value: "Refreshed" }], getAttribute: (name) => name.toLowerCase() === "ows_id" ? "9" : null };
      return {
        querySelector: () => null,
        getElementsByTagName(name) {
          if (name === "z:row") return invalid ? [] : [row];
          if (name === "Changes") return [{ getAttribute: (attribute) => attribute === "LastChangeToken" ? "token-refreshed" : null }];
          if (name === "Id") return invalid ? [{ getAttribute: (attribute) => attribute === "ChangeType" ? "InvalidToken" : null }] : [];
          return [];
        },
      };
    }
  };
  const calls = [];
  const writes = [];
  let requestCount = 0;
  const store = {
    async getCursor() { return "stale-token"; },
    async setCursor(_scope, _site, key, token) { writes.push({ type: "cursor", key, token }); },
    async clearCursor(_scope, _site, key) { writes.push({ type: "clear-cursor", key }); },
    async clearItems(_scope, _site, key) { writes.push({ type: "clear-items", key }); },
    async putItems(_scope, _site, key, items) { writes.push({ type: "items", key, items }); },
    async deleteItems() {},
  };
  const rest = {
    async raw(path, options) {
      calls.push({ path, query: options.body.query });
      return { text: async () => (++requestCount === 1 ? "invalid" : "fresh") };
    },
  };
  try {
    const result = await syncListWithChangeToken({ rest, store, scope: "scope", siteId: "site", listId, fields: ["Title"] });
    assert.equal(calls.length, 2);
    assert.equal(calls[0].query.ChangeToken, "stale-token");
    assert.equal(calls[1].query.ChangeToken, null, "the retry starts a full refresh without the stale token");
    assert.ok(writes.some((write) => write.type === "clear-cursor"));
    assert.ok(writes.some((write) => write.type === "clear-items"));
    assert.deepEqual(result, { pages: 1, changed: 1, token: "token-refreshed" });
    assert.ok(writes.some((write) => write.type === "items" && write.items[0].fields.Title === "Refreshed"));
    assert.ok(writes.some((write) => write.type === "cursor" && write.token === "token-refreshed"));
  } finally {
    if (previousParser === undefined) delete globalThis.DOMParser;
    else globalThis.DOMParser = previousParser;
  }
});
