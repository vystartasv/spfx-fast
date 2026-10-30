import "fake-indexeddb/auto";
import assert from "node:assert/strict";
import test from "node:test";
import { applyItemQuery, createClient, createIndexedDbStore } from "../dist/index.js";

const dbName = () => `spfx-fast-test-${Math.random().toString(36).slice(2)}`;

test("local sorting keeps stable order for equal values", () => {
  const items = [
    { id: "a-first", fields: { Title: "A" } },
    { id: "b", fields: { Title: "B" } },
    { id: "a-second", fields: { Title: "A" } },
  ];

  assert.deepEqual(applyItemQuery(items, { sort: { field: "Title" }, limit: 2 }).map(({ id }) => id), ["a-first", "a-second"]);
  assert.deepEqual(applyItemQuery(items, { sort: { field: "Title", direction: "desc" }, limit: 2 }).map(({ id }) => id), ["b", "a-first"]);
  assert.deepEqual(applyItemQuery(items, { limit: 0 }), []);
});

test("cached Graph pages are served first and refreshed into IndexedDB", async () => {
  const scope = "tenant:user";
  const url = "https://graph.microsoft.com/v1.0/sites/site/lists/list/items";
  const store = createIndexedDbStore(dbName());
  let calls = 0;
  const client = createClient({
    scope,
    store,
    get: async () => ({ value: [{ id: String(++calls) }] }),
  });

  const cold = await client.readPage(url);
  assert.equal(cold.source, "network");
  await cold.refresh;
  await new Promise((resolve) => setTimeout(resolve, 0));

  const warm = await client.readPage(url);
  assert.equal(warm.source, "cache");
  assert.equal(warm.data.value[0].id, "1");
  await warm.refresh;
  await new Promise((resolve) => setTimeout(resolve, 0));

  const refreshed = await client.readPage(url);
  assert.equal(refreshed.data.value[0].id, "2");
});

test("IndexedDB local queries match the shared filter, sort, and limit behavior", async () => {
  const scope = "tenant:user";
  const store = createIndexedDbStore(dbName());
  const rows = Array.from({ length: 120 }, (_, index) => ({
    id: String(index),
    fields: { Title: `Task ${String(index).padStart(3, "0")}`, Status: index % 3 ? "Open" : "Closed" },
  }));
  await store.putItems(scope, "site", "list", rows);
  const client = createClient({ scope, store, get: async () => ({ value: [] }) });
  for (const query of [
    { equals: { Status: "Open" }, limit: 10 },
    { contains: { field: "Title", value: "1" }, sort: { field: "Title", direction: "desc" }, limit: 7 },
  ]) {
    const stored = await store.getItems(scope, "site", "list");
    assert.deepEqual(await client.queryItems("site", "list", query), applyItemQuery(stored, query));
  }
});

test("delta sync rebuilds after an expired cursor and applies tombstones", async () => {
  const scope = "tenant:user";
  const siteId = "site";
  const listId = "list";
  const stale = "https://graph.microsoft.com/v1.0/sites/site/lists/list/items/delta?token=stale";
  const fresh = "https://graph.microsoft.com/v1.0/sites/site/lists/list/items/delta?token=fresh";
  const store = createIndexedDbStore(dbName());
  await store.setCursor(scope, siteId, listId, stale);
  await store.putItems(scope, siteId, listId, [{ id: "old", fields: { Title: "Old" } }]);
  let freshSyncs = 0;
  const client = createClient({
    scope,
    store,
    get: async (url) => {
      if (url === stale) throw Object.assign(new Error("Expired delta link"), { status: 410 });
      if (url.startsWith("https://graph.microsoft.com/v1.0/sites/site/lists/list/items/delta")) {
        freshSyncs++;
        return {
          value: freshSyncs === 1
            ? [{ id: "new", fields: { Title: "New" } }]
            : [{ id: "new", deleted: { state: "deleted" } }],
          "@odata.deltaLink": fresh,
        };
      }
      throw new Error(`Unexpected URL ${url}`);
    },
  });

  await client.syncList(siteId, listId);
  assert.deepEqual((await client.queryItems(siteId, listId)).map((item) => item.id), ["new"]);
  assert.equal(await store.getCursor(scope, siteId, listId), fresh);

  await client.syncList(siteId, listId);
  assert.deepEqual(await client.queryItems(siteId, listId), []);
});

test("malformed Graph delta cursors do not clear cached items or replace the checkpoint", async () => {
  const scope = "tenant:user";
  const siteId = "site";
  const listId = "list";
  const malformed = "https://graph.microsoft.com/v1.0/sites/site/lists/list/items/delta?token=malformed";
  const store = createIndexedDbStore(dbName());
  await store.setCursor(scope, siteId, listId, malformed);
  await store.putItems(scope, siteId, listId, [{ id: "cached", fields: { Title: "Keep until a full sync succeeds" } }]);
  const client = createClient({
    scope,
    store,
    get: async (url) => {
      if (url === malformed) throw Object.assign(new Error("Invalid delta request"), { status: 400, code: "invalidRequest" });
      throw new Error(`Unexpected URL ${url}`);
    },
  });

  await assert.rejects(client.syncList(siteId, listId), { status: 400, code: "invalidRequest" });
  assert.deepEqual((await store.getItems(scope, siteId, listId)).map((item) => item.id), ["cached"]);
  assert.equal(await store.getCursor(scope, siteId, listId), malformed);
});

test("drive delta state and cursor persist in IndexedDB and invalidate cached folder pages", async () => {
  const scope = "tenant:user";
  const driveId = "drive";
  const driveStoreSiteId = "__spfx_fast_drive_items__";
  const db = dbName();
  const store = createIndexedDbStore(db);
  const folderPage = "https://graph.microsoft.com/v1.0/drives/drive/items/folder/children?$top=50";
  const oldDelta = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=old";
  const newDelta = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=new";
  let pageCalls = 0;
  await store.set(scope, folderPage, { value: [{ id: "stale" }] });
  await store.setCursor(scope, driveStoreSiteId, driveId, oldDelta);
  await store.putItems(scope, driveStoreSiteId, driveId, [{ id: "prior", name: "prior.txt", file: {} }]);
  const client = createClient({
    scope,
    store,
    get: async (url) => {
      if (url === oldDelta) return { value: [{ id: "fresh", name: "fresh.txt", file: {} }], "@odata.deltaLink": newDelta };
      if (url === folderPage) return { value: [{ id: `page-${++pageCalls}` }] };
      throw new Error(`Unexpected URL ${url}`);
    },
  });

  await client.syncDrive(driveId);
  assert.deepEqual((await client.getDriveItems(driveId)).map(({ id }) => id).sort(), ["fresh", "prior"]);
  assert.equal(await store.getCursor(scope, driveStoreSiteId, driveId), newDelta);
  assert.deepEqual((await client.readPage(folderPage)).data.value, [{ id: "page-1" }]);
  const reopened = createIndexedDbStore(db);
  assert.equal(await reopened.getCursor(scope, driveStoreSiteId, driveId), newDelta);
  assert.deepEqual((await reopened.getItems(scope, driveStoreSiteId, driveId)).map(({ id }) => id).sort(), ["fresh", "prior"]);
});

test("optimistic item writes preserve ETags and roll back failed updates", async () => {
  const store = createIndexedDbStore(dbName());
  await store.putItems("tenant:user", "site", "list", [{ id: "1", eTag: '"v1"', fields: { Title: "Before" } }]);
  let requestOptions;
  const client = createClient({
    scope: "tenant:user",
    store,
    get: async () => ({ value: [] }),
    request: async (_url, options) => {
      requestOptions = options;
      throw Object.assign(new Error("ETag conflict"), { status: 412 });
    },
  });

  await assert.rejects(client.updateItem("site", "list", "1", { Title: "After" }), /ETag conflict/);

  assert.equal(requestOptions.headers["If-Match"], '"v1"');
  assert.equal((await client.queryItems("site", "list"))[0].fields.Title, "Before");
});

test("paged items can be updated before delta sync with their returned ETag", async () => {
  const store = createIndexedDbStore(dbName());
  let requestOptions;
  const client = createClient({
    scope: "tenant:user",
    store,
    get: async () => ({ value: [{ id: "1", eTag: '"v1"', fields: { Title: "Laptop", Status: "Pending" } }] }),
    request: async (_url, options) => {
      requestOptions = options;
      return { id: "1", eTag: '"v2"', fields: { Title: "Laptop", Status: "Approved" } };
    },
  });

  const page = await client.getItemPage("site", "list", { fields: ["Title", "Status"] });
  await page.refresh;
  await assert.rejects(client.updateItem("site", "list", page.items[0].id, { Status: "Approved" }), /pass its ETag/);
  const updated = await client.updateItem("site", "list", page.items[0].id, { Status: "Approved" }, { eTag: page.items[0].eTag });

  assert.equal(requestOptions.headers["If-Match"], '"v1"');
  assert.equal(updated.fields.Status, "Approved");
  assert.equal((await client.queryItems("site", "list"))[0].fields.Status, "Approved");
});
