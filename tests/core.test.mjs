import assert from "node:assert/strict";
import test from "node:test";
import { createClient, createMemoryStore, DriveUploadInterruptedError, isDriveUploadInterruptedError } from "../dist/index.js";
import { createSPFxGraphTransport, createSPFxRestTransport } from "../dist/spfx.js";

const driveStoreSiteId = "__spfx_fast_drive_items__";

test("uses volatile storage to keep delta sync working when IndexedDB is unavailable", async () => {
  const storageErrors = [];
  const client = createClient({
    scope: "tenant:user",
    get: async (url) => url.includes("/delta")
      ? { value: [{ id: "1", fields: { Title: "Cached offline" } }], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/sites/site/lists/list/items/delta?token=next" }
      : { value: [] },
    onStorageError: (error) => storageErrors.push(error),
  });

  await client.syncList("site", "list");

  assert.equal((await client.queryItems("site", "list"))[0]?.fields?.Title, "Cached offline");
  assert.ok(storageErrors.length > 0);
});

test("list delta rejects continuation and delta links for another Graph resource", async () => {
  const escapedLinks = [
    { "@odata.nextLink": "https://graph.microsoft.com/v1.0/sites/site/lists/other/items/delta?token=next" },
    { "@odata.deltaLink": "https://graph.microsoft.com/v1.0/sites/other/lists/list/items/delta?token=done" },
    { "@odata.nextLink": "https://graph.microsoft.com/v1.0/sites/site/lists/list%2Fother/items/delta?token=escape" },
    { "@odata.nextLink": "https://evil.example/v1.0/sites/site/lists/list/items/delta?token=next" },
  ];
  for (const links of escapedLinks) {
    const client = createClient({ scope: "tenant:user", get: async () => ({ value: [], ...links }) });
    await assert.rejects(client.syncList("site", "list"), /requested list items delta endpoint/);
  }
});

test("list delta accepts equivalent URL encoding while pinning its resource path", async () => {
  const siteId = "tenant.sharepoint.com,site-id,web-id";
  const next = "https://graph.microsoft.com/v1.0/sites/tenant.sharepoint.com,site-id,web-id/lists/list-id/items/delta?token=next";
  const delta = "https://graph.microsoft.com/v1.0/sites/tenant.sharepoint.com%2Csite-id%2Cweb-id/lists/list-id/items/delta?token=done";
  const client = createClient({
    scope: "tenant:user",
    get: async (url) => url === next
      ? { value: [], "@odata.deltaLink": delta }
      : { value: [], "@odata.nextLink": next },
  });
  const result = await client.syncList(siteId, "list-id");
  assert.equal(result.pages, 2);
  assert.equal(result.deltaLink, delta);
});

test("list delta rebuilds cached items after an expired cursor", async () => {
  const scope = "tenant:user";
  const store = createMemoryStore();
  const stale = "https://graph.microsoft.com/v1.0/sites/site/lists/list/items/delta?token=stale";
  const fresh = "https://graph.microsoft.com/v1.0/sites/site/lists/list/items/delta?token=fresh";
  await store.setCursor(scope, "site", "list", stale);
  await store.putItems(scope, "site", "list", [{ id: "old", fields: { Title: "Old item" } }]);
  const requests = [];
  const client = createClient({
    scope,
    store,
    get: async (url) => {
      requests.push(url);
      if (url === stale) throw Object.assign(new Error("Expired list cursor"), { status: 410 });
      const target = new URL(url);
      assert.equal(target.pathname, "/v1.0/sites/site/lists/list/items/delta");
      assert.equal(target.searchParams.has("token"), false);
      return { value: [{ id: "current", fields: { Title: "Current item" } }], "@odata.deltaLink": fresh };
    },
  });

  await client.syncList("site", "list");

  assert.equal(requests.length, 2);
  assert.deepEqual((await store.getItems(scope, "site", "list")).map(({ id }) => id), ["current"]);
  assert.equal(await store.getCursor(scope, "site", "list"), fresh);
});

test("list delta restarts from a same-resource Graph 410 Location link", async () => {
  const scope = "tenant:user";
  const store = createMemoryStore();
  const stale = "https://graph.microsoft.com/v1.0/sites/site/lists/list/items/delta?token=stale";
  const restart = "https://graph.microsoft.com/v1.0/sites/site/lists/list/items/delta?token=restart";
  const fresh = "https://graph.microsoft.com/v1.0/sites/site/lists/list/items/delta?token=fresh";
  await store.setCursor(scope, "site", "list", stale);
  await store.putItems(scope, "site", "list", [{ id: "old", fields: { Title: "Old item" } }]);
  const requests = [];
  const client = createClient({
    scope,
    store,
    get: async (url) => {
      requests.push(url);
      if (url === stale) throw Object.assign(new Error("Resync required"), { statusCode: 410, headers: new Headers({ Location: restart }) });
      assert.equal(url, restart, "the service-provided restart link is followed instead of replaying the stale cursor");
      return { value: [{ id: "current", fields: { Title: "Current item" } }], "@odata.deltaLink": fresh };
    },
  });
  await client.syncList("site", "list");
  assert.deepEqual(requests, [stale, restart]);
  assert.deepEqual((await store.getItems(scope, "site", "list")).map(({ id }) => id), ["current"]);
  assert.equal(await store.getCursor(scope, "site", "list"), fresh);
});

test("list delta ignores an unsafe Graph 410 Location and falls back to its original endpoint", async () => {
  const scope = "tenant:user";
  const store = createMemoryStore();
  const stale = "https://graph.microsoft.com/v1.0/sites/site/lists/list/items/delta?token=stale";
  const unsafe = "https://evil.example/v1.0/sites/site/lists/list/items/delta?token=escape";
  await store.setCursor(scope, "site", "list", stale);
  const requests = [];
  const client = createClient({
    scope,
    store,
    get: async (url) => {
      requests.push(url);
      if (url === stale) throw Object.assign(new Error("Resync required"), { status: 410, headers: new Headers({ Location: unsafe }) });
      return { value: [{ id: "current", fields: { Title: "Current item" } }], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/sites/site/lists/list/items/delta?token=fresh" };
    },
  });
  await client.syncList("site", "list");
  assert.equal(requests.length, 2);
  assert.equal(new URL(requests[1]).origin, "https://graph.microsoft.com");
  assert.equal(new URL(requests[1]).pathname, "/v1.0/sites/site/lists/list/items/delta");
  assert.equal(requests.includes(unsafe), false);
});

test("memory store notifies subscribers after item updates", async () => {
  const store = createMemoryStore();
  let changes = 0;
  const unsubscribe = store.subscribe("tenant:user", "site", "list", () => changes++);

  await store.putItems("tenant:user", "site", "list", [{ id: "1", fields: { Title: "Task" } }]);
  await store.deleteItems("tenant:user", "site", "list", ["1"]);
  unsubscribe();

  assert.equal(changes, 2);
  assert.deepEqual(await store.getItems("tenant:user", "site", "list"), []);
});

test("Graph batch uses version-relative resource paths and rejects malformed routes", async () => {
  const requests = [];
  const client = createClient({
    scope: "tenant:user",
    get: async () => ({}),
    request: async (url, options) => {
      requests.push({ url, options });
      return { responses: [{ id: "site", status: 200, body: { id: "site-id" } }] };
    },
  });
  const response = await client.batch([{ id: "site", method: "GET", url: "/sites/site-id?$select=id" }]);
  assert.equal(response[0].status, 200);
  assert.equal(requests[0].url, "https://graph.microsoft.com/v1.0/$batch");
  assert.equal(requests[0].options.body.requests[0].url, "/sites/site-id?$select=id");
  await client.batch([{ id: "create", method: "POST", url: "/sites/site-id/lists/list-id/items", body: { fields: { Title: "Synthetic" } } }]);
  assert.equal(requests[1].options.body.requests[0].headers["Content-Type"], "application/json");
  await client.batch([{ id: "custom", method: "PATCH", url: "/sites/site-id/lists/list-id/items/item-id/fields", headers: { "content-type": "application/custom+json" }, body: { Title: "Updated" } }]);
  assert.equal(requests[2].options.body.requests[0].headers["content-type"], "application/custom+json");
  await assert.rejects(() => client.batch([{ id: "version", method: "GET", url: "/v1.0/sites/site-id" }]), /omit the version/);
  await assert.rejects(() => client.batch([{ id: "external", method: "GET", url: "//evil.example/path" }]), /Graph resource paths/);
  await assert.rejects(() => client.batch([{ id: "case", method: "GET", url: "/sites/site-id" }, { id: "CASE", method: "GET", url: "/sites/site-id" }]), /unique/);
  await assert.rejects(() => client.batch([{ id: "invalid", method: "GET", url: 7 }]), /Graph resource paths/);
  await assert.rejects(() => client.batch(Array.from({ length: 21 }, (_, index) => ({ id: String(index), method: "GET", url: "/sites/site-id" }))), /between 1 and 20/);
});

test("bounded local query results are cached briefly and invalidated by list changes", async () => {
  const store = createMemoryStore();
  await store.putItems("tenant:user", "site", "list", [{ id: "1", fields: { Title: "Original", Status: "Open" } }]);
  const queryItems = store.queryItems.bind(store);
  let queryRuns = 0;
  store.queryItems = async (...args) => { queryRuns++; return queryItems(...args); };
  const metrics = [];
  const client = createClient({ scope: "tenant:user", store, get: async () => ({ value: [] }), onMetric: (event) => metrics.push(event) });
  const query = { equals: { Status: "Open" }, sort: { field: "Title" }, limit: 20 };

  const first = await client.queryItems("site", "list", query);
  first[0].fields.Title = "Caller edit";
  const cached = await client.queryItems("site", "list", query);
  assert.equal(queryRuns, 1, "the second request reuses the bounded query result");
  assert.equal(cached[0].fields.Title, "Original", "callers receive a copy of cached items");
  assert.equal(metrics.at(-1).type, "cache-hit");

  await new Promise((resolve) => setTimeout(resolve, 1010));
  await client.queryItems("site", "list", query);
  assert.equal(queryRuns, 2, "expired query results are refreshed");

  await store.putItems("tenant:user", "site", "list", [{ id: "2", fields: { Title: "Another", Status: "Open" } }]);
  const updated = await client.queryItems("site", "list", query);
  assert.equal(queryRuns, 3, "a list update invalidates cached query results");
  assert.equal(updated.length, 2);
});

test("drive item content is cached and returned before its background refresh", async () => {
  let requests = 0;
  const client = createClient({
    scope: "tenant:user",
    store: createMemoryStore(),
    get: async () => ({ value: [] }),
    request: async (_url, options) => { requests++; return options?.responseType === "text" ? "# Team notes" : new Blob(["file"]); },
  });
  const first = await client.readDriveItemContent("drive-id", "item-id", { as: "text" });
  assert.equal(first.source, "network");
  assert.equal(first.data, "# Team notes");
  const second = await client.readDriveItemContent("drive-id", "item-id", { as: "text" });
  assert.equal(second.source, "cache");
  assert.equal(second.data, "# Team notes");
  await second.refresh;
  assert.equal(requests, 2);
});

test("downloadDriveItem returns fresh authenticated bytes without using persistent cache", async () => {
  let requests = 0;
  const store = createMemoryStore();
  let writes = 0;
  const originalSet = store.set.bind(store);
  store.set = (...args) => { writes++; return originalSet(...args); };
  const client = createClient({
    scope: "tenant:user",
    store,
    get: async () => ({ value: [] }),
    request: async (url, options) => {
      requests++;
      assert.match(url, /\/drives\/drive-id\/items\/item-id\/content$/);
      assert.equal(options?.responseType, "blob");
      return new Blob(["current bytes"]);
    },
  });
  const file = await client.downloadDriveItem("drive-id", "item-id");
  assert.equal(await file.text(), "current bytes");
  assert.equal(requests, 1);
  assert.equal(writes, 0);
});

test("drive children cursors are pinned to the requested drive item", async () => {
  const client = createClient({ scope: "tenant:user", get: async () => ({ value: [] }) });
  await assert.rejects(
    client.getDriveItemPage("drive-a", "folder-a", { cursor: "https://graph.microsoft.com/v1.0/drives/drive-b/items/folder-b/children" }),
    /requested drive children path/,
  );
});

test("drive item permission reads page through authenticated Graph without caching", async () => {
  const next = "https://graph.microsoft.com/v1.0/drives/drive-a/items/item-a/permissions?$skiptoken=next";
  const urls = [];
  const store = createMemoryStore();
  let writes = 0;
  const originalSet = store.set.bind(store);
  store.set = (...args) => { writes++; return originalSet(...args); };
  const client = createClient({
    scope: "tenant:user",
    store,
    get: async (url) => {
      urls.push(url);
      return url === next
        ? { value: [{ id: "p2", roles: ["read"] }] }
        : { value: [{ id: "p1", roles: ["write"], grantedToV2: { user: { displayName: "Ada" } } }], "@odata.nextLink": next };
    },
  });
  const permissions = await client.getDriveItemPermissions("drive-a", "item-a");
  assert.deepEqual(permissions.map(({ id }) => id), ["p1", "p2"]);
  assert.equal(new URL(urls[0]).pathname, "/v1.0/drives/drive-a/items/item-a/permissions");
  assert.equal(urls.length, 2);
  assert.equal(writes, 0);
});

test("sharing links require an explicit audience and preserve inherited permissions", async () => {
  const calls = [];
  const client = createClient({
    scope: "tenant:user",
    get: async () => ({ value: [] }),
    request: async (url, options) => {
      calls.push({ url, options });
      return { id: "permission-1", link: { type: "view", scope: "organization", webUrl: "https://contoso.sharepoint.com/:w:/s/site/file" } };
    },
  });
  const permission = await client.createDriveItemSharingLink("drive", "item", {
    type: "view", scope: "organization", expirationDateTime: "2026-12-31T23:59:59Z",
  });
  assert.equal(permission.id, "permission-1");
  assert.equal(calls[0].url, "https://graph.microsoft.com/v1.0/drives/drive/items/item/createLink");
  assert.deepEqual(calls[0].options.body, {
    type: "view", scope: "organization", retainInheritedPermissions: true,
    expirationDateTime: "2026-12-31T23:59:59Z",
  });
  await assert.rejects(client.createDriveItemSharingLink("drive", "item", { type: "view", scope: "anonymous", expirationDateTime: "not-a-date" }), /ISO date-time/);
});

test("permission removal re-reads access and refuses inherited grants", async () => {
  const calls = [];
  let inherited = true;
  const client = createClient({
    scope: "tenant:user",
    get: async () => ({ value: [] }),
    request: async (url, options) => {
      calls.push({ url, options });
      return options?.method === "DELETE" ? undefined : {
        id: "permission/1",
        inheritedFrom: inherited ? { id: "parent" } : null,
      };
    },
  });
  await assert.rejects(client.deleteDriveItemPermission("drive", "item", "permission/1"), /inherited permission/);
  assert.equal(calls.length, 1, "an inherited grant is never deleted");
  inherited = false;
  await client.deleteDriveItemPermission("drive", "item", "permission/1");
  assert.equal(calls.length, 3, "the direct permission is re-read immediately before deletion");
  assert.equal(calls[2].options.method, "DELETE");
  assert.match(calls[0].url, /permissions\/permission%2F1$/);
});

test("drive copy returns a validated monitor URL and honors version-history constraints", async () => {
  const calls = [];
  const response = { status: 202, headers: { get: (name) => name.toLowerCase() === "location" ? "https://contoso.sharepoint.com/_api/v2.0/monitor/job-1" : null } };
  const client = createClient({
    scope: "tenant:user",
    get: async () => ({ value: [] }),
    request: async (url, options) => { calls.push({ url, options }); return response; },
  });
  const operation = await client.copyDriveItem("source", "file", {
    targetDriveId: "target", targetFolderId: "folder", conflictBehavior: "rename", includeAllVersionHistory: true,
  });
  assert.equal(operation.monitorUrl, "https://contoso.sharepoint.com/_api/v2.0/monitor/job-1");
  assert.equal(new URL(calls[0].url).searchParams.get("@microsoft.graph.conflictBehavior"), "rename");
  assert.deepEqual(calls[0].options.body, { parentReference: { driveId: "target", id: "folder" }, includeAllVersionHistory: true });
  assert.equal(calls[0].options.responseType, "raw");
  await assert.rejects(client.copyDriveItem("source", "file", { targetDriveId: "target", targetFolderId: "folder", name: "renamed.docx", includeAllVersionHistory: true }), /without a new name/);
});

test("drive copy rejects non-accepted responses and unsafe monitor URLs", async () => {
  const clientWith = (response) => createClient({ scope: "tenant:user", get: async () => ({ value: [] }), request: async () => response });
  await assert.rejects(clientWith({ status: 400 }).copyDriveItem("source", "file", { targetDriveId: "target", targetFolderId: "folder" }), /HTTP 400/);
  await assert.rejects(clientWith({ status: 202, headers: { get: () => "https://contoso.sharepoint.com/_api/v2.0/not-a-monitor/job" } }).copyDriveItem("source", "file", { targetDriveId: "target", targetFolderId: "folder" }), /invalid copy operation monitor URL/);
});

test("calendar view pages stay on the selected date window and return a resumable cursor", async () => {
  const startDateTime = "2026-10-29T00:00:00Z";
  const endDateTime = "2026-10-30T00:00:00Z";
  const nextLink = `https://graph.microsoft.com/v1.0/me/calendar/calendarView?startDateTime=${encodeURIComponent(startDateTime)}&endDateTime=${encodeURIComponent(endDateTime)}&$skip=1`;
  const finalLink = `https://graph.microsoft.com/v1.0/me/calendar/calendarView?startDateTime=${encodeURIComponent(startDateTime)}&endDateTime=${encodeURIComponent(endDateTime)}&$skip=3`;
  const calls = [];
  const client = createClient({ scope: "tenant:user", get: async () => ({ value: [] }), request: async (url, options) => {
    calls.push({ url, options });
    const skip = new URL(url).searchParams.get("$skip");
    if (skip === "1") return { value: [{ id: "two" }, { id: "three" }], "@odata.nextLink": finalLink };
    if (skip === "3") return { value: [{ id: "four" }] };
    return { value: [{ id: "one" }], "@odata.nextLink": nextLink };
  } });
  const partial = await client.getCalendarView(startDateTime, endDateTime, { maxItems: 2 });
  assert.deepEqual(partial.items.map((item) => item.id), ["one"]);
  assert.equal(partial.complete, false);
  assert.equal(new URL(calls[0].url).pathname, "/v1.0/me/calendar/calendarView");
  assert.equal(new URL(calls[0].url).searchParams.get("$top"), "2");
  assert.equal(calls[0].options.signal, undefined);
  const resumed = await client.getCalendarView(startDateTime, endDateTime, { cursor: partial.nextCursor, maxItems: 2 });
  assert.deepEqual(resumed.items.map((item) => item.id), ["two", "three"]);
  assert.equal(resumed.complete, false);
  const final = await client.getCalendarView(startDateTime, endDateTime, { cursor: resumed.nextCursor, maxItems: 2 });
  assert.deepEqual(final.items.map((item) => item.id), ["four"]);
  assert.equal(final.complete, true);
  assert.equal(calls.length, 4);
  await assert.rejects(client.getCalendarView("2026-10-29T00:00:00", endDateTime), /timezone/);
  await assert.rejects(client.getCalendarView(startDateTime, endDateTime, { cursor: "https://graph.microsoft.com/v1.0/me/messages" }), /requested calendar window/);
});

test("calendar view uses the selected calendar route", async () => {
  let requestedUrl = "";
  const client = createClient({
    scope: "tenant:user",
    get: async () => ({ value: [] }),
    request: async (url) => { requestedUrl = url; return { value: [] }; },
  });
  await client.getCalendarView("2026-10-29T00:00:00Z", "2026-10-30T00:00:00Z", { calendarId: "calendar-id" });
  assert.equal(new URL(requestedUrl).pathname, "/v1.0/me/calendars/calendar-id/calendarView");
});

test("drive item permission pagination rejects escaped cursors and excessive pages", async () => {
  const client = createClient({
    scope: "tenant:user",
    get: async () => ({ value: [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/drives/drive-b/items/item-b/permissions" }),
  });
  await assert.rejects(client.getDriveItemPermissions("drive-a", "item-a"), /requested drive item permissions path/);
  await assert.rejects(client.getDriveItemPermissions("drive-a", "item-a", { maxPages: 0 }), /maxPages/);
  const loop = createClient({ scope: "tenant:user", get: async () => ({ value: [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/drives/drive-a/items/item-a/permissions?next=1" }) });
  await assert.rejects(loop.getDriveItemPermissions("drive-a", "item-a", { maxPages: 1 }), /exceeded 1 pages/);
});

test("large file upload uses the documented parent path, sequential 320 KiB ranges, and no credentials on upload URL", async () => {
  const originalFetch = globalThis.fetch;
  const ranges = [];
  let graphCalls = 0;
  let sessionPath = "";
  let sessionBody;
  globalThis.fetch = async (_url, init = {}) => {
    assert.equal(init.credentials, "omit");
    assert.equal(init.headers?.Authorization, undefined);
    ranges.push(init.headers?.["Content-Range"]);
    return ranges.length < 2
      ? new Response(JSON.stringify({ nextExpectedRanges: [`${10 * 1024 * 1024}-`] }), { status: 202, headers: { "Content-Type": "application/json" } })
      : new Response(JSON.stringify({ id: "uploaded-id", name: "report.bin", size: 700000 }), { status: 201, headers: { "Content-Type": "application/json" } });
  };
  try {
    const client = createClient({
      scope: "tenant:user",
      store: createMemoryStore(),
      get: async () => ({ value: [] }),
      request: async (url, options = {}) => {
        graphCalls++;
        if (new URL(url).pathname.endsWith("/root")) return { id: "root-item-id" };
        sessionPath = new URL(url).pathname;
        sessionBody = options.body;
        return { uploadUrl: "https://upload.example.test/session-secret" };
      },
    });
    const size = 10 * 1024 * 1024 + 1;
    const result = await client.uploadDriveItem("drive-id", "root", "report.bin", new Blob([new Uint8Array(size)]), { chunkSize: 10 * 1024 * 1024 });
    assert.equal(result.id, "uploaded-id");
    assert.equal(graphCalls, 2);
    assert.equal(sessionPath, "/v1.0/drives/drive-id/items/root-item-id:/report.bin:/createUploadSession");
    assert.equal(sessionBody, undefined, "Graph's default fail behavior needs no request body");
    assert.deepEqual(ranges, [`bytes 0-${10 * 1024 * 1024 - 1}/${size}`, `bytes ${10 * 1024 * 1024}-${size - 1}/${size}`]);
  } finally { globalThis.fetch = originalFetch; }
});

test("large file upload resumes at the server-reported range after a network interruption", async () => {
  const originalFetch = globalThis.fetch;
  const chunkSize = 10 * 1024 * 1024;
  const size = chunkSize + 1;
  const requests = [];
  globalThis.fetch = async (_url, init = {}) => {
    assert.equal(init.credentials, "omit");
    const method = init.method ?? "GET";
    requests.push({ method, range: init.headers?.["Content-Range"] });
    if (method === "GET") return new Response(JSON.stringify({ nextExpectedRanges: [`${chunkSize}-`] }), { status: 200, headers: { "Content-Type": "application/json" } });
    if (requests.filter((request) => request.method === "PUT").length === 1) throw new TypeError("simulated connection loss after server accepted the first chunk");
    return new Response(JSON.stringify({ id: "resumed-upload", name: "report.bin", size }), { status: 201, headers: { "Content-Type": "application/json" } });
  };
  try {
    const client = createClient({
      scope: "tenant:user",
      store: createMemoryStore(),
      get: async () => ({ value: [] }),
      request: async () => ({ uploadUrl: "https://upload.example.test/session-secret" }),
    });
    const result = await client.uploadDriveItem("drive-id", "folder-id", "report.bin", new Blob([new Uint8Array(size)]), { chunkSize });
    assert.equal(result.id, "resumed-upload");
    assert.deepEqual(requests, [
      { method: "PUT", range: `bytes 0-${chunkSize - 1}/${size}` },
      { method: "GET", range: undefined },
      { method: "PUT", range: `bytes ${chunkSize}-${size - 1}/${size}` },
    ]);
  } finally { globalThis.fetch = originalFetch; }
});

test("large file upload retries transient offline failures while checking the session range", async () => {
  const originalFetch = globalThis.fetch;
  const chunkSize = 10 * 1024 * 1024;
  const size = chunkSize + 1;
  const requests = [];
  let statusChecks = 0;
  globalThis.fetch = async (_url, init = {}) => {
    const method = init.method ?? "GET";
    requests.push(method);
    if (method === "GET" && ++statusChecks === 1) throw new TypeError("browser is still offline");
    if (method === "GET") return new Response(JSON.stringify({ nextExpectedRanges: [`${chunkSize}-`] }), { status: 200, headers: { "Content-Type": "application/json" } });
    if (requests.filter((request) => request === "PUT").length === 1) throw new TypeError("chunk accepted but response lost");
    return new Response(JSON.stringify({ id: "resumed-after-reconnect", name: "report.bin", size }), { status: 201, headers: { "Content-Type": "application/json" } });
  };
  try {
    const client = createClient({
      scope: "tenant:user",
      store: createMemoryStore(),
      get: async () => ({ value: [] }),
      request: async () => ({ uploadUrl: "https://upload.example.test/session-secret" }),
    });
    const result = await client.uploadDriveItem("drive-id", "folder-id", "report.bin", new Blob([new Uint8Array(size)]), { chunkSize });
    assert.equal(result.id, "resumed-after-reconnect");
    assert.deepEqual(requests, ["PUT", "GET", "GET", "PUT"]);
  } finally { globalThis.fetch = originalFetch; }
});

test("large upload returns a private resume handle after its retry window and resumes the same session", async () => {
  const originalFetch = globalThis.fetch;
  const originalSetTimeout = globalThis.setTimeout;
  const chunkSize = 10 * 1024 * 1024;
  const size = chunkSize + 1;
  const ranges = [];
  let offline = true;
  let statusChecks = 0;
  let sessionCreates = 0;
  let cancels = 0;
  globalThis.setTimeout = (callback) => { queueMicrotask(callback); return 0; };
  globalThis.fetch = async (_url, init = {}) => {
    const method = init.method ?? "GET";
    if (method === "DELETE") { cancels++; return new Response(null, { status: 204 }); }
    if (method === "PUT") ranges.push(init.headers?.["Content-Range"]);
    if (offline) {
      if (method === "GET") statusChecks++;
      throw new TypeError("simulated network outage");
    }
    if (method === "GET") return new Response(JSON.stringify({ nextExpectedRanges: [`${chunkSize}-`] }), { status: 200, headers: { "Content-Type": "application/json" } });
    return new Response(JSON.stringify({ id: "resumed-upload", name: "report.bin", size }), { status: 201, headers: { "Content-Type": "application/json" } });
  };
  try {
    const client = createClient({
      scope: "tenant:user",
      store: createMemoryStore(),
      get: async () => ({ value: [] }),
      request: async () => { sessionCreates++; return { uploadUrl: "https://upload.example.test/session-secret" }; },
    });
    let interrupted;
    try {
      await client.uploadDriveItem("drive-id", "folder-id", "report.bin", new Blob([new Uint8Array(size)]), { chunkSize });
    } catch (error) { interrupted = error; }
    assert.ok(interrupted instanceof DriveUploadInterruptedError);
    assert.ok(isDriveUploadInterruptedError(interrupted));
    assert.equal(statusChecks, 6);
    assert.equal(cancels, 0, "do not delete a session whose accepted byte range is unknown");
    assert.match(interrupted.message, /Resume upload/);

    offline = false;
    const result = await interrupted.resume();
    assert.equal(result.id, "resumed-upload");
    assert.equal(sessionCreates, 1, "resume the original session instead of creating a second upload");
    assert.deepEqual(ranges, [
      `bytes 0-${chunkSize - 1}/${size}`,
      `bytes ${chunkSize}-${size - 1}/${size}`,
    ]);
    assert.equal(cancels, 0);
  } finally {
    globalThis.fetch = originalFetch;
    globalThis.setTimeout = originalSetTimeout;
  }
});

test("large file upload verifies the target when the final chunk response is lost", async () => {
  const originalFetch = globalThis.fetch;
  const chunkSize = 10 * 1024 * 1024;
  const file = new Blob([new Uint8Array(chunkSize), new Uint8Array([42])]);
  const calls = [];
  let puts = 0;
  globalThis.fetch = async (_url, init = {}) => {
    const method = init.method ?? "GET";
    calls.push(method);
    if (method === "GET") return new Response("expired", { status: 404 });
    if (++puts === 1) return new Response(JSON.stringify({ nextExpectedRanges: [`${chunkSize}-`] }), { status: 202, headers: { "Content-Type": "application/json" } });
    throw new TypeError("final chunk was committed but its response was lost");
  };
  try {
    const client = createClient({
      scope: "tenant:user",
      store: createMemoryStore(),
      get: async () => ({ value: [] }),
      request: async (url) => {
        if (url.endsWith("/createUploadSession")) return { uploadUrl: "https://upload.example.test/session-secret" };
        if (url.endsWith("/content")) return file;
        return { id: "committed-item", name: "report.bin", size: file.size };
      },
    });
    const result = await client.uploadDriveItem("drive-id", "folder-id", "report.bin", file, { chunkSize });
    assert.equal(result.id, "committed-item");
    assert.deepEqual(calls, ["PUT", "PUT", "GET"]);
  } finally { globalThis.fetch = originalFetch; }
});

test("large replace upload reconciles a lost final response against the destination bytes", async () => {
  const originalFetch = globalThis.fetch;
  const chunkSize = 10 * 1024 * 1024;
  const file = new Blob([new Uint8Array(chunkSize), new Uint8Array([73])]);
  const requests = [];
  let puts = 0;
  globalThis.fetch = async (_url, init = {}) => {
    const method = init.method ?? "GET";
    requests.push(method);
    if (method === "GET") return new Response("expired", { status: 404 });
    if (++puts === 1) return new Response(JSON.stringify({ nextExpectedRanges: [`${chunkSize}-`] }), { status: 202, headers: { "Content-Type": "application/json" } });
    throw new TypeError("final response was lost");
  };
  try {
    const client = createClient({
      scope: "tenant:user",
      store: createMemoryStore(),
      get: async () => ({ value: [] }),
      request: async (url) => {
        if (url.endsWith("/createUploadSession")) return { uploadUrl: "https://upload.example.test/session-secret" };
        if (url.endsWith("/content")) return file;
        return { id: "replaced-item", name: "report.bin", size: file.size };
      },
    });

    const result = await client.uploadDriveItem("drive-id", "folder-id", "report.bin", file, { chunkSize, conflictBehavior: "replace" });

    assert.equal(result.id, "replaced-item");
    assert.deepEqual(requests, ["PUT", "PUT", "GET"]);
  } finally { globalThis.fetch = originalFetch; }
});

test("large replace upload creates its session against an existing drive item ID", async () => {
  const originalFetch = globalThis.fetch;
  const chunkSize = 10 * 1024 * 1024;
  const size = chunkSize + 1;
  const graphRequests = [];
  let puts = 0;
  globalThis.fetch = async (_url, init = {}) => {
    if (++puts === 1) return new Response(JSON.stringify({ nextExpectedRanges: [`${chunkSize}-`] }), { status: 202, headers: { "Content-Type": "application/json" } });
    return new Response(JSON.stringify({ id: "existing-item", name: "report.bin", size }), { status: 201, headers: { "Content-Type": "application/json" } });
  };
  try {
    const client = createClient({
      scope: "tenant:user",
      store: createMemoryStore(),
      get: async () => ({ value: [] }),
      request: async (url, config = {}) => {
        graphRequests.push({ url, config });
        if (url.includes("/items/folder-id:/report.bin:")) return { id: "existing-item", name: "report.bin", file: {} };
        if (url.endsWith("/createUploadSession")) return { uploadUrl: "https://upload.example.test/session-secret" };
        throw new Error(`Unexpected Graph request: ${url}`);
      },
    });

    const result = await client.uploadDriveItem("drive-id", "folder-id", "report.bin", new Blob([new Uint8Array(size)]), { chunkSize, conflictBehavior: "replace" });

    assert.equal(result.id, "existing-item");
    assert.match(graphRequests[0].url, /\/items\/folder-id:\/report\.bin:\?\$select=id,name,folder$/);
    assert.equal(graphRequests[1].url, "https://graph.microsoft.com/v1.0/drives/drive-id/items/existing-item/createUploadSession");
    assert.equal(graphRequests[1].config.method, "POST");
    assert.equal(graphRequests[1].config.body, undefined);
  } finally { globalThis.fetch = originalFetch; }
});

test("large replace upload creates a name-based session when the destination does not exist", async () => {
  const originalFetch = globalThis.fetch;
  const chunkSize = 10 * 1024 * 1024;
  const size = chunkSize + 1;
  const graphRequests = [];
  let puts = 0;
  globalThis.fetch = async () => ++puts === 1
    ? new Response(JSON.stringify({ nextExpectedRanges: [`${chunkSize}-`] }), { status: 202, headers: { "Content-Type": "application/json" } })
    : new Response(JSON.stringify({ id: "new-item", name: "report.bin", size }), { status: 201, headers: { "Content-Type": "application/json" } });
  try {
    const client = createClient({
      scope: "tenant:user",
      store: createMemoryStore(),
      get: async () => ({ value: [] }),
      request: async (url, config = {}) => {
        graphRequests.push({ url, config });
        if (url.includes("/items/folder-id:/report.bin:?$select=")) throw Object.assign(new Error("Not found"), { status: 404 });
        if (url.endsWith("/createUploadSession")) return { uploadUrl: "https://upload.example.test/session-secret" };
        throw new Error(`Unexpected Graph request: ${url}`);
      },
    });

    const result = await client.uploadDriveItem("drive-id", "folder-id", "report.bin", new Blob([new Uint8Array(size)]), { chunkSize, conflictBehavior: "replace" });

    assert.equal(result.id, "new-item");
    assert.equal(graphRequests[1].url, "https://graph.microsoft.com/v1.0/drives/drive-id/items/folder-id:/report.bin:/createUploadSession");
    assert.deepEqual(graphRequests[1].config.body, { item: { "@microsoft.graph.conflictBehavior": "replace" } });
  } finally { globalThis.fetch = originalFetch; }
});

test("large rename upload reconciles a lost final response through bounded name probes and exact bytes", async () => {
  const originalFetch = globalThis.fetch;
  const chunkSize = 10 * 1024 * 1024;
  const file = new Blob([new Uint8Array(chunkSize), new Uint8Array([73])]);
  const destinationReads = [];
  let puts = 0;
  let cancels = 0;
  globalThis.fetch = async (_url, init = {}) => {
    const method = init.method ?? "GET";
    if (method === "GET") return new Response("expired", { status: 404 });
    if (method === "DELETE") { cancels++; return new Response(null, { status: 204 }); }
    if (++puts === 1) return new Response(JSON.stringify({ nextExpectedRanges: [`${chunkSize}-`] }), { status: 202, headers: { "Content-Type": "application/json" } });
    throw new TypeError("final response was lost");
  };
  try {
    const client = createClient({
      scope: "tenant:user",
      store: createMemoryStore(),
      get: async () => ({ value: [] }),
      request: async (url) => {
        if (url.endsWith("/createUploadSession")) return { uploadUrl: "https://upload.example.test/session-secret" };
        if (url.endsWith("/content")) return file;
        destinationReads.push(url);
        if (url.includes("/items/folder-id:/report%201.bin:")) return { id: "renamed-item", name: "report 1.bin", size: file.size, file: {}, parentReference: { id: "folder-id" } };
        throw Object.assign(new Error("Not found"), { status: 404 });
      },
    });

    const item = await client.uploadDriveItem("drive-id", "folder-id", "report.bin", file, { chunkSize, conflictBehavior: "rename" });

    assert.equal(item.id, "renamed-item");
    assert.equal(item.name, "report 1.bin");
    assert.equal(puts, 2);
    assert.equal(cancels, 0, "do not cancel a completed renamed upload");
    assert.ok(destinationReads[0].includes("/items/folder-id:/report%201.bin:"));
  } finally { globalThis.fetch = originalFetch; }
});

test("large rename upload does not reconcile a same-size candidate with different bytes", async () => {
  const originalFetch = globalThis.fetch;
  const originalSetTimeout = globalThis.setTimeout;
  const chunkSize = 10 * 1024 * 1024;
  const file = new Blob([new Uint8Array(chunkSize), new Uint8Array([73])]);
  let puts = 0;
  let cancels = 0;
  globalThis.setTimeout = (callback) => { queueMicrotask(callback); return 0; };
  globalThis.fetch = async (_url, init = {}) => {
    const method = init.method ?? "GET";
    if (method === "GET") return new Response("expired", { status: 404 });
    if (method === "DELETE") { cancels++; return new Response(null, { status: 204 }); }
    if (++puts === 1) return new Response(JSON.stringify({ nextExpectedRanges: [`${chunkSize}-`] }), { status: 202, headers: { "Content-Type": "application/json" } });
    throw new TypeError("final response was lost");
  };
  try {
    const client = createClient({
      scope: "tenant:user",
      store: createMemoryStore(),
      get: async () => ({ value: [] }),
      request: async (url) => {
        if (url.endsWith("/createUploadSession")) return { uploadUrl: "https://upload.example.test/session-secret" };
        if (url.includes("/items/folder-id:/report%201.bin:")) return { id: "different-file", name: "report 1.bin", size: file.size, file: {}, parentReference: { id: "folder-id" } };
        if (url.endsWith("/content")) return new Blob(["different bytes"]);
        throw Object.assign(new Error("Not found"), { status: 404 });
      },
    });

    await assert.rejects(
      client.uploadDriveItem("drive-id", "folder-id", "report.bin", file, { chunkSize, conflictBehavior: "rename" }),
      /Upload may have completed, but its final response was lost and the session expired\./,
    );
    assert.equal(cancels, 1);
  } finally { globalThis.fetch = originalFetch; globalThis.setTimeout = originalSetTimeout; }
});

test("large replace upload verifies committed bytes after session-status retries are exhausted", async () => {
  const originalFetch = globalThis.fetch;
  const originalSetTimeout = globalThis.setTimeout;
  const chunkSize = 10 * 1024 * 1024;
  const file = new Blob([new Uint8Array(chunkSize), new Uint8Array([73])]);
  let puts = 0;
  let statusReads = 0;
  let cancels = 0;
  globalThis.setTimeout = (callback) => { queueMicrotask(callback); return 0; };
  globalThis.fetch = async (_url, init = {}) => {
    const method = init.method ?? "GET";
    if (method === "GET") { statusReads++; throw new TypeError("still offline"); }
    if (method === "DELETE") { cancels++; return new Response(null, { status: 204 }); }
    if (++puts === 1) return new Response(JSON.stringify({ nextExpectedRanges: [`${chunkSize}-`] }), { status: 202, headers: { "Content-Type": "application/json" } });
    throw new TypeError("final response was lost");
  };
  try {
    const client = createClient({
      scope: "tenant:user",
      store: createMemoryStore(),
      get: async () => ({ value: [] }),
      request: async (url) => {
        if (url.endsWith("/createUploadSession")) return { uploadUrl: "https://upload.example.test/session-secret" };
        if (url.endsWith("/content")) return file;
        return { id: "replaced-item", name: "report.bin", size: file.size };
      },
    });

    const result = await client.uploadDriveItem("drive-id", "folder-id", "report.bin", file, { chunkSize, conflictBehavior: "replace" });

    assert.equal(result.id, "replaced-item");
    assert.equal(puts, 2);
    assert.equal(statusReads, 6);
    assert.equal(cancels, 0, "leave the uncertain session alone so Graph can finish processing");
  } finally {
    globalThis.fetch = originalFetch;
    globalThis.setTimeout = originalSetTimeout;
  }
});

test("small root-drive uploads use the documented binary content route", async () => {
  let requestUrl = "";
  let requestOptions;
  const content = new Blob(["x"], { type: "text/plain" });
  const client = createClient({
    scope: "tenant:user",
    get: async () => ({ value: [] }),
    request: async (url, options) => { requestUrl = url; requestOptions = options; return { id: "root-uploaded-id", name: "check.txt", size: 1 }; },
  });
  const result = await client.uploadDriveItem("drive-id", "root", "check.txt", content);
  assert.equal(result.id, "root-uploaded-id");
  assert.equal(new URL(requestUrl).pathname, "/v1.0/drives/drive-id/items/root:/check.txt:/content");
  assert.equal(new URL(requestUrl).searchParams.get("@microsoft.graph.conflictBehavior"), "fail");
  assert.equal(requestOptions.method, "PUT");
  assert.equal(requestOptions.headers["Content-Type"], "text/plain");
  assert.equal(requestOptions.rawBody, content);
  await client.uploadDriveItem("drive-id", "root", "check.txt", content, { conflictBehavior: "rename" });
  assert.equal(new URL(requestUrl).searchParams.get("@microsoft.graph.conflictBehavior"), "rename");
});

test("failed upload sessions are cancelled without credentials", async () => {
  const originalFetch = globalThis.fetch;
  const methods = [];
  globalThis.fetch = async (_url, options = {}) => {
    methods.push(options.method ?? "GET");
    assert.equal(options.credentials, "omit");
    return options.method === "DELETE" ? new Response(null, { status: 204 }) : new Response("failed", { status: 400 });
  };
  try {
    const client = createClient({
      scope: "tenant:user",
      get: async () => ({ value: [] }),
      request: async () => ({ uploadUrl: "https://upload.example.test/session-secret" }),
    });
    await assert.rejects(client.uploadDriveItem("drive-id", "folder-id", "report.bin", new Blob([new Uint8Array(10 * 1024 * 1024 + 1)])), /HTTP 400/);
    await new Promise((resolve) => setTimeout(resolve, 0));
    assert.deepEqual(methods, ["PUT", "DELETE"]);
  } finally { globalThis.fetch = originalFetch; }
});

test("agent page cursors cannot escape the requested list path", async () => {
  const client = createClient({ scope: "tenant:user", get: async () => ({ value: [] }) });
  await assert.rejects(
    client.getItemPage("site", "allowed-list", { cursor: "https://graph.microsoft.com/v1.0/sites/site/lists/other-list/items" }),
    /Continuation cursors must belong to the requested/,
  );
});

test("list content types use Graph v1.0 and follow every continuation page", async () => {
  const next = "https://graph.microsoft.com/v1.0/sites/site-id/lists/list-id/contentTypes?$skiptoken=next";
  const urls = [];
  const client = createClient({
    scope: "tenant:user",
    get: async (url) => {
      urls.push(url);
      return url === next
        ? { value: [{ id: "ct-task", name: "Task" }] }
        : { value: [{ id: "ct-item", name: "Item" }], "@odata.nextLink": next };
    },
  });
  const types = await client.getListContentTypes("site-id", "list-id");
  assert.deepEqual(types.map(({ id }) => id), ["ct-item", "ct-task"]);
  assert.equal(new URL(urls[0]).pathname, "/v1.0/sites/site-id/lists/list-id/contentTypes");
  assert.equal(urls.length, 2);
});

test("list-item and drive-item version reads are typed, paged, bounded, and uncached", async () => {
  const listNext = "https://graph.microsoft.com/v1.0/sites/site/lists/list/items/item/versions?$skiptoken=next";
  const urls = [];
  const client = createClient({
    scope: "tenant:user",
    store: createMemoryStore(),
    get: async (url) => {
      urls.push(url);
      if (url.includes("/lists/")) return url === listNext
        ? { value: [{ id: "1.0", fields: { Title: "old" } }, { id: "0.9" }] }
        : { value: [{ id: "2.0", fields: { Title: "new" } }], "@odata.nextLink": listNext };
      return { value: [{ id: "3.0", size: 30 }] };
    },
  });
  const listVersions = await client.getListItemVersions("site", "list", "item", { maxItems: 2 });
  const driveVersions = await client.getDriveItemVersions("drive", "drive item");
  await client.getListItemVersions("site", "list", "item", { maxItems: 1 });
  assert.deepEqual(listVersions.map(({ id }) => id), ["2.0", "1.0"]);
  assert.deepEqual(driveVersions, [{ id: "3.0", size: 30 }]);
  assert.equal(new URL(urls[0]).pathname, "/v1.0/sites/site/lists/list/items/item/versions");
  assert.equal(new URL(urls[1]).pathname, "/v1.0/sites/site/lists/list/items/item/versions");
  assert.equal(new URL(urls[2]).pathname, "/v1.0/drives/drive/items/drive%20item/versions");
  assert.equal(urls.length, 4, "repeated version reads fetch the Graph endpoint again instead of using cached pages");
});

test("version reads reject escaped continuation links and invalid bounds", async () => {
  for (const nextLink of [
    "https://user@graph.microsoft.com/v1.0/sites/site/lists/list/items/item/versions?token=escape",
    "https://graph.microsoft.com/v1.0/sites/other/lists/list/items/item/versions?token=escape",
  ]) {
    const client = createClient({
      scope: "tenant:user",
      get: async () => ({ value: [{ id: "1.0" }], "@odata.nextLink": nextLink }),
    });
    await assert.rejects(client.getListItemVersions("site", "list", "item"), /requested item versions endpoint/);
  }
  const client = createClient({ scope: "tenant:user", get: async () => ({ value: [] }) });
  await assert.rejects(client.getDriveItemVersions("drive", "item", { maxItems: 0 }), /maxItems/);
  await assert.rejects(client.getDriveItemVersions("drive", "item", { maxPages: 101 }), /maxPages/);
});

test("previous file-version downloads are uncached and restore uses the explicit Graph action", async () => {
  const calls = [];
  const client = createClient({
    scope: "tenant:user",
    store: createMemoryStore(),
    get: async () => ({ value: [] }),
    request: async (url, options) => {
      calls.push({ url, options });
      return options?.responseType === "blob" ? new Blob(["previous bytes"]) : undefined;
    },
  });

  const bytes = await client.downloadDriveItemVersion("drive", "item", "2/1");
  await client.restoreDriveItemVersion("drive", "item", "2/1");
  assert.equal(await bytes.text(), "previous bytes");
  assert.match(calls[0].url, /\/versions\/2%2F1\/content$/);
  assert.equal(calls[0].options.responseType, "blob");
  assert.equal(calls[1].url, "https://graph.microsoft.com/v1.0/drives/drive/items/item/versions/2%2F1/restoreVersion");
  assert.equal(calls[1].options.method, "POST");
  assert.equal(calls[1].options.body, undefined);
});

test("drive delta sync caches the hierarchy, resumes bounded pages, and applies deletions", async () => {
  const scope = "tenant:user";
  const store = createMemoryStore();
  const next = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=page-2";
  const delta = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=delta-1";
  const delta2 = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=delta-2";
  const urls = [];
  const client = createClient({
    scope,
    store,
    get: async (url) => {
      urls.push(url);
      if (url === next) return { value: [{ id: "file", name: "notes.md", parentReference: { id: "folder" }, file: {} }], "@odata.deltaLink": delta };
      if (url === delta) return { value: [{ id: "file", deleted: { state: "deleted" } }], "@odata.deltaLink": delta2 };
      return { value: [{ id: "folder", name: "Team", folder: {} }, { id: "old", name: "old.txt", file: {} }], "@odata.nextLink": next };
    },
  });

  const partial = await client.syncDrive("drive", { maxPages: 1 });
  assert.equal(partial.complete, false);
  assert.equal(partial.nextCursor, next);
  assert.deepEqual((await client.getDriveItems("drive")).map(({ id }) => id), ["folder", "old"]);
  assert.equal(await store.getCursor(scope, driveStoreSiteId, "drive"), undefined);

  const complete = await client.syncDrive("drive");
  assert.equal(complete.complete, true);
  assert.equal(complete.deltaLink, delta);
  assert.deepEqual((await client.getDriveItems("drive")).map(({ id }) => id).sort(), ["file", "folder", "old"]);
  assert.equal(await store.getCursor(scope, driveStoreSiteId, "drive"), delta);

  await client.syncDrive("drive");
  assert.deepEqual((await client.getDriveItems("drive")).map(({ id }) => id).sort(), ["folder", "old"]);
  assert.equal(urls[1], next, "the next invocation resumes from the saved page checkpoint");
  assert.equal(urls[2], delta, "later syncs use the final delta link");
});

test("drive delta sync keeps deleted folders until children are processed", async () => {
  const scope = "tenant:user";
  const store = createMemoryStore();
  const prior = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=prior";
  const next = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=children";
  const fresh = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=fresh";
  await store.putItems(scope, driveStoreSiteId, "drive", [
    { id: "folder", folder: {} },
    { id: "child", name: "child.txt", parentReference: { id: "folder" }, file: {} },
  ]);
  await store.setCursor(scope, driveStoreSiteId, "drive", prior);
  const client = createClient({
    scope,
    store,
    get: async (url) => {
      if (url === next) return { value: [{ id: "child", deleted: { state: "deleted" } }], "@odata.deltaLink": fresh };
      assert.equal(url, prior);
      return { value: [
        { id: "folder", deleted: { state: "deleted" } },
        { id: "child", deleted: { state: "deleted" } },
      ], "@odata.nextLink": next };
    },
  });

  const partial = await client.syncDrive("drive", { maxPages: 1 });
  assert.equal(partial.complete, false);
  assert.deepEqual((await client.getDriveItems("drive")).map(({ id }) => id), ["folder"]);
  await client.syncDrive("drive");
  assert.deepEqual(await client.getDriveItems("drive"), []);
  assert.equal(await store.getCursor(scope, driveStoreSiteId, "drive"), fresh);
});

test("drive delta sync rebuilds after an expired cursor", async () => {
  const scope = "tenant:user";
  const store = createMemoryStore();
  const stale = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=stale";
  const fresh = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=fresh";
  await store.setCursor(scope, driveStoreSiteId, "drive", stale);
  await store.putItems(scope, driveStoreSiteId, "drive", [{ id: "old", name: "Old item", file: {} }]);
  const client = createClient({
    scope,
    store,
    get: async (url) => {
      if (url === stale) throw Object.assign(new Error("Expired drive cursor"), { status: 410 });
      return { value: [{ id: "current", name: "Current item", file: {} }], "@odata.deltaLink": fresh };
    },
  });
  await client.syncDrive("drive");
  assert.deepEqual((await client.getDriveItems("drive")).map(({ id }) => id), ["current"]);
  assert.equal(await store.getCursor(scope, driveStoreSiteId, "drive"), fresh);
});

test("drive delta restarts from a same-drive Graph 410 Location link", async () => {
  const scope = "tenant:user";
  const store = createMemoryStore();
  const stale = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=stale";
  const restart = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=restart";
  const fresh = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=fresh";
  await store.setCursor(scope, driveStoreSiteId, "drive", stale);
  await store.putItems(scope, driveStoreSiteId, "drive", [{ id: "old", name: "Old item", file: {} }]);
  const requests = [];
  const client = createClient({
    scope,
    store,
    get: async (url) => {
      requests.push(url);
      if (url === stale) throw Object.assign(new Error("Resync required"), { status: 410, headers: new Headers({ Location: restart }) });
      assert.equal(url, restart, "the service-provided restart link is followed instead of replaying the stale cursor");
      return { value: [{ id: "current", name: "Current item", file: {} }], "@odata.deltaLink": fresh };
    },
  });
  await client.syncDrive("drive");
  assert.deepEqual(requests, [stale, restart]);
  assert.deepEqual((await client.getDriveItems("drive")).map(({ id }) => id), ["current"]);
  assert.equal(await store.getCursor(scope, driveStoreSiteId, "drive"), fresh);
});

test("drive delta ignores an unsafe Graph 410 Location and falls back to its original endpoint", async () => {
  const scope = "tenant:user";
  const store = createMemoryStore();
  const stale = "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=stale";
  const unsafe = "https://evil.example/v1.0/drives/drive/root/delta?token=escape";
  await store.setCursor(scope, driveStoreSiteId, "drive", stale);
  const requests = [];
  const client = createClient({
    scope,
    store,
    get: async (url) => {
      requests.push(url);
      if (url === stale) throw Object.assign(new Error("Resync required"), { status: 410, headers: new Headers({ Location: unsafe }) });
      return { value: [{ id: "current", name: "Current item", file: {} }], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/drives/drive/root/delta?token=fresh" };
    },
  });
  await client.syncDrive("drive");
  assert.equal(requests.length, 2);
  assert.equal(new URL(requests[1]).origin, "https://graph.microsoft.com");
  assert.equal(new URL(requests[1]).pathname, "/v1.0/drives/drive/root/delta");
  assert.equal(requests.includes(unsafe), false);
});

test("drive delta sync rejects continuation links outside the requested drive endpoint", async () => {
  const client = createClient({
    scope: "tenant:user",
    get: async () => ({ value: [{ id: "folder", folder: {} }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/drives/other/root/delta?token=escape" }),
  });
  await assert.rejects(client.syncDrive("drive"), /requested drive root delta endpoint/);
  await assert.rejects(client.syncDrive("drive", { maxPages: 1001 }), /maxPages/);
});

test("drive folder creation, same-drive moves, and recycle-bin deletion use authenticated Graph and update local state", async () => {
  const scope = "tenant:user";
  const driveId = "drive";
  const driveStoreSiteId = "__spfx_fast_drive_items__";
  const store = createMemoryStore();
  await store.putItems(scope, driveStoreSiteId, driveId, [
    { id: "folder", folder: {}, eTag: '"folder-v1"' },
    { id: "child", name: "child.txt", file: {}, parentReference: { id: "folder" } },
    { id: "moving", name: "old.txt", file: {}, eTag: '"item-v1"' },
  ]);
  const calls = [];
  const client = createClient({
    scope,
    store,
    get: async () => ({ value: [] }),
    request: async (url, config) => {
      calls.push({ url, ...config });
      if (config.method === "POST") return { id: "new-folder", name: "New", folder: {} };
      if (config.method === "PATCH") return { id: "moving", name: "renamed.txt", file: {}, parentReference: { id: "new-folder" } };
      return undefined;
    },
  });

  const folder = await client.createDriveFolder(driveId, "root-id", "New");
  assert.equal(folder.id, "new-folder");
  assert.deepEqual(calls[0], {
    url: "https://graph.microsoft.com/v1.0/drives/drive/items/root-id/children",
    method: "POST",
    body: { name: "New", folder: {}, "@microsoft.graph.conflictBehavior": "fail" },
    signal: undefined,
  });

  await client.moveDriveItem(driveId, "moving", "new-folder", { name: "renamed.txt" });
  assert.equal(calls[1].method, "PATCH");
  assert.equal(calls[1].headers["If-Match"], '"item-v1"');
  assert.deepEqual(calls[1].body, { parentReference: { id: "new-folder" }, name: "renamed.txt" });

  await client.deleteDriveItem(driveId, "folder");
  assert.equal(calls[2].method, "DELETE");
  assert.equal(calls[2].headers["If-Match"], '"folder-v1"');
  assert.deepEqual((await client.getDriveItems(driveId)).map(({ id }) => id).sort(), ["moving", "new-folder"]);
  await assert.rejects(client.createDriveFolder(driveId, "root-id", "../unsafe"), /leaf name/);
  await assert.rejects(client.moveDriveItem(driveId, "moving", "folder", { name: "nested/name" }), /leaf name/);
});

test("taxonomy helpers enumerate term store groups, sets, and hierarchical children", async () => {
  const urls = [];
  const client = createClient({
    scope: "tenant:user",
    get: async (url) => { urls.push(url); return { value: [{ id: "term-1", labels: [{ name: "Operations", languageTag: "en-US" }] }] }; },
  });
  const [groups, sets, children] = await Promise.all([
    client.getTermStoreGroups("site-id"),
    client.getTermStoreSets("site-id", "group-id"),
    client.getTermChildren("site-id", "set-id", "parent-id"),
  ]);
  assert.equal(groups[0]?.id, "term-1");
  assert.equal(sets[0]?.id, "term-1");
  assert.equal(children[0]?.labels?.[0]?.name, "Operations");
  assert.deepEqual(urls.map((url) => new URL(url).pathname), [
    "/v1.0/sites/site-id/termStore/groups",
    "/v1.0/sites/site-id/termStore/groups/group-id/sets",
    "/v1.0/sites/site-id/termStore/sets/set-id/terms/parent-id/children",
  ]);
});

test("Graph list filtering escapes literals and keeps continuation cursors on the same filter", async () => {
  const cursor = "https://graph.microsoft.com/v1.0/sites/site/lists/list/items?$skiptoken=next&$filter=fields%2FStatus%20eq%20%27O%27%27Brien%27";
  const urls = [];
  const client = createClient({
    scope: "tenant:user",
    store: createMemoryStore(),
    get: async (url) => { urls.push(url); return { value: [], "@odata.nextLink": cursor }; },
  });
  const filter = { field: "Status", operator: "eq", value: "O'Brien" };
  const first = await client.getItemPage("site", "list", { filter });
  assert.equal(new URL(urls[0]).searchParams.get("$filter"), "fields/Status eq 'O''Brien'");
  await client.getItemPage("site", "list", { cursor: first.nextCursor, filter });
  await assert.rejects(client.getItemPage("site", "list", { cursor: first.nextCursor }), /same Graph filter/);
  await assert.rejects(client.getItemPage("site", "list", { filter: { field: "Status eq true", operator: "eq", value: true } }), /internal field name/);
});

test("bounded listItems uses the server filter before following pages", async () => {
  let requested;
  const client = createClient({
    scope: "tenant:user",
    store: createMemoryStore(),
    get: async (url) => { requested = url; return { value: [{ id: "1" }] }; },
  });
  const rows = [];
  for await (const item of client.listItems("site", "list", {
    filter: { field: "Status", operator: "eq", value: "Open" },
    maxItems: 1,
  })) rows.push(item);
  assert.deepEqual(rows.map((item) => item.id), ["1"]);
  assert.equal(new URL(requested).searchParams.get("$filter"), "fields/Status eq 'Open'");
});

test("listItems follows next links and stops at maxItems", async () => {
  const next = "https://graph.microsoft.com/v1.0/sites/site/lists/list/items?$skiptoken=next";
  const urls = [];
  const client = createClient({
    scope: "tenant:user",
    get: async (url) => {
      urls.push(url);
      return url === next
        ? { value: [{ id: "2" }, { id: "3" }] }
        : { value: [{ id: "1" }], "@odata.nextLink": next };
    },
  });
  const items = [];
  for await (const item of client.listItems("site", "list", { pageSize: 1, maxItems: 2 })) items.push(item.id);

  assert.deepEqual(items, ["1", "2"]);
  assert.equal(urls.length, 2);
});

test("SPFx Graph transport supports PUT binary bodies and raw responses", async () => {
  const calls = [];
  const builder = {
    header() { return this; },
    responseType(value) { calls.push(["responseType", value]); return this; },
    get: async () => ({}),
    post: async () => new Response(null, { status: 202 }),
    put: async (body) => { calls.push(["put", body]); return new Blob(["ok"]); },
    patch: async () => ({}),
    delete: async () => undefined,
  };
  const transport = createSPFxGraphTransport({ api: () => builder });
  const body = new Blob(["file contents"]);

  const response = await transport.request("https://graph.microsoft.com/v1.0/me/drive/root:/file.txt:/content", {
    method: "PUT",
    rawBody: body,
    responseType: "blob",
  });

  assert.ok(response instanceof Blob);
  assert.equal(calls[0][0], "responseType");
  assert.equal(calls[1][0], "put");
  assert.equal(calls[1][1], body);
  const raw = await transport.request("https://graph.microsoft.com/v1.0/drives/source/items/file/copy", { method: "POST", responseType: "raw" });
  assert.equal(raw.status, 202);
  assert.deepEqual(calls.filter(([name]) => name === "responseType").map(([, value]) => value), ["blob", "raw"]);
});

test("SPFx Graph transport retries recognized network failures on reads", async () => {
  let attempts = 0;
  const retries = [];
  const transport = createSPFxGraphTransport({ api: () => ({
    get: async () => {
      attempts++;
      if (attempts === 1) throw new TypeError("Failed to fetch");
      return { value: [] };
    },
    post: async () => ({}),
    put: async () => ({}),
    patch: async () => ({}),
    delete: async () => undefined,
  }) }, { maxRetries: 1, maxBackoffMs: 0, onRetry: (delayMs, error) => retries.push({ delayMs, error }) });
  assert.deepEqual(await transport.get("https://graph.microsoft.com/v1.0/sites/site/lists"), { value: [] });
  assert.equal(attempts, 2);
  assert.equal(retries.length, 1);
  assert.equal(retries[0].delayMs, 0);
});

test("SPFx REST transport preserves binary bodies and reports retry delays", async () => {
  const binary = new Blob(["upload"]);
  const metrics = [];
  const requests = [];
  let readAttempt = 0;
  const transport = createSPFxRestTransport({
    fetch: async (url, _configuration, options) => {
      requests.push({ url, options });
      if (options.method === "GET" && ++readAttempt === 1) {
        return { ok: false, status: 503, headers: { get: () => "0" }, json: async () => ({}), text: async () => "", blob: async () => new Blob() };
      }
      return { ok: true, status: 200, headers: { get: () => null }, json: async () => ({ value: [] }), text: async () => "", blob: async () => new Blob(["download"]) };
    },
  }, {}, "https://contoso.sharepoint.com/sites/team", { maxRetries: 1, maxBackoffMs: 0, onMetric: (metric) => metrics.push(metric) });

  const uploaded = await transport.raw("/_api/web/GetFolderByServerRelativeUrl('Docs')/Files/add(url='a.bin',overwrite=true)", {
    method: "POST",
    rawBody: binary,
    headers: { "Content-Type": "application/octet-stream" },
  });
  const listed = await transport.get("/_api/web/lists");

  assert.equal(requests[0].options.body, binary);
  assert.equal(requests[0].options.headers["Content-Type"], "application/octet-stream");
  assert.equal(new URL(requests[0].url).pathname, "/sites/team/_api/web/GetFolderByServerRelativeUrl('Docs')/Files/add(url='a.bin',overwrite=true)");
  assert.equal(new URL(requests[1].url).pathname, "/sites/team/_api/web/lists");
  assert.equal((await uploaded.blob()).size, 8);
  assert.deepEqual(listed, { value: [] });
  assert.equal(metrics[0].type, "retry");
  assert.equal(metrics[0].delayMs, 0);
});
