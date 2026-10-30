import assert from "node:assert/strict";
import test from "node:test";
import { createLegacySharePointApi } from "../dist/sharepoint.js";

const listId = "00000000-0000-0000-0000-000000000123";
const viewId = "10000000-0000-0000-0000-000000000123";
const recycleBinItemId = "20000000-0000-0000-0000-000000000123";

test("recycle-bin helpers page both response formats and validate restore and purge routes", async () => {
  const calls = [];
  const next = "https://contoso.sharepoint.com/sites/team/_api/web/RecycleBin?$skiptoken=next";
  const item = { Id: recycleBinItemId, LeafName: "report.docx", DirName: "Shared Documents", ItemType: 1, Size: 1024 };
  const api = createLegacySharePointApi({
    siteUrl: "https://contoso.sharepoint.com/sites/team",
    get: async (url) => {
      calls.push({ method: "GET", url });
      return url.includes("$skiptoken")
        ? { d: { results: [item] } }
        : { value: [{ ...item, LeafName: "notes.md" }], "@odata.nextLink": next };
    },
    request: async (url, options) => { calls.push({ method: options?.method, url, options }); return undefined; },
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  });

  const first = await api.getRecycleBinItemsPage({ pageSize: 500 });
  assert.equal(first.items[0].LeafName, "notes.md");
  assert.equal(new URL(calls[0].url).searchParams.get("$top"), "100");
  assert.equal(new URL(calls[0].url).searchParams.get("$select"), "Id,Title,ItemType,ItemState,LeafName,DirName,DeletedDate,DeletedByName,AuthorName,Size");
  const second = await api.getRecycleBinItemsPage({ cursor: first.nextCursor });
  assert.deepEqual(second.items, [item]);

  await api.restoreRecycleBinItem(recycleBinItemId);
  await api.permanentlyDeleteRecycleBinItem(recycleBinItemId, { scope: "site" });
  assert.equal(calls[2].url, `/_api/web/RecycleBin('${recycleBinItemId}')/restore()`);
  assert.equal(calls[2].method, "POST");
  assert.equal(calls[3].url, `/_api/site/RecycleBin('${recycleBinItemId}')/deleteObject()`);
  assert.equal(calls[3].method, "POST");
  assert.throws(() => api.restoreRecycleBinItem("bad-id"), /IDs must be GUIDs/);
});

test("regional settings helper reads modern and verbose REST response shapes", async () => {
  const calls = [];
  const settings = { LocaleId: 1033, CalendarType: 1, Time24: false, FirstDayOfWeek: 0 };
  const api = createLegacySharePointApi({
    get: async (url) => { calls.push(url); return calls.length === 1 ? settings : { d: settings }; },
    request: async () => undefined,
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  });
  assert.deepEqual(await api.getRegionalSettings(), settings);
  assert.deepEqual(await api.getRegionalSettings(), settings);
  assert.deepEqual(calls, ["/_api/web/RegionalSettings", "/_api/web/RegionalSettings"]);
  const malformed = createLegacySharePointApi({
    get: async () => ({ d: null }), request: async () => undefined,
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  });
  await assert.rejects(malformed.getRegionalSettings(), /must be an object/);
});

test("typed list attachment helpers preserve binary bodies and use documented REST routes", async () => {
  const calls = [];
  const content = new Blob(["attachment data"]);
  const response = {
    status: 200,
    json: async () => ({ Name: "brief.txt" }),
    blob: async () => content,
  };
  const rest = {
    get: async (url) => { calls.push({ method: "GET", url }); return { value: [] }; },
    request: async (url, options) => { calls.push({ method: options?.method, url, options }); return undefined; },
    raw: async (url, options) => { calls.push({ method: options?.method ?? "GET", url, options }); return response; },
  };
  const api = createLegacySharePointApi(rest);

  await api.listItemAttachments(listId, 42);
  assert.deepEqual(await api.readItemAttachment(listId, 42, "brief.txt"), content);
  await api.addItemAttachment(listId, 42, "brief.txt", content);
  await api.updateItemAttachment(listId, 42, "brief.txt", content, "etag-1");
  await api.deleteItemAttachment(listId, 42, "brief.txt", "etag-2");

  assert.match(calls[0].url, /items\(42\)\/AttachmentFiles$/);
  assert.match(calls[1].url, /getByFileName\('brief\.txt'\)\/\$value$/);
  assert.equal(calls[2].options.rawBody, content);
  assert.match(calls[2].url, /add\(FileName='brief\.txt'\)$/);
  assert.equal(calls[3].options.headers["X-HTTP-Method"], "PUT");
  assert.equal(calls[3].options.headers["If-Match"], "etag-1");
  assert.equal(calls[4].method, "DELETE");
  assert.equal(calls[4].options.headers["If-Match"], "etag-2");
});

test("typed attachment helpers reject non-numeric item IDs and path-like file names", async () => {
  const api = createLegacySharePointApi({ get: async () => ({}), request: async () => undefined, raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }) });
  assert.throws(() => api.listItemAttachments("list", "1)\u002fitems(2"), /numeric/);
  await assert.rejects(api.readItemAttachment("list", 1, "../secret.txt"), /leaf name/);
});

test("typed Search REST builder quotes query values, includes selected options, and bounds paging", async () => {
  let requested;
  const api = createLegacySharePointApi({
    get: async (url) => { requested = url; return { d: { query: {} } }; },
    request: async () => undefined,
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  });
  await api.search(`Title:O'Brien`, { rowLimit: 30, startRow: 20, selectProperties: ["Title", "Path"], refiners: ["FileExtension"], sortList: "Rank:descending", trimDuplicates: false });
  const query = new URL(requested, "https://contoso.sharepoint.com").searchParams;
  assert.equal(query.get("querytext"), "'Title:O''Brien'");
  assert.equal(query.get("rowlimit"), "30");
  assert.equal(query.get("startrow"), "20");
  assert.equal(query.get("selectproperties"), "'Title,Path'");
  assert.equal(query.get("refiners"), "'FileExtension'");
  assert.equal(query.get("sortlist"), "'Rank:descending'");
  assert.equal(query.get("trimduplicates"), "false");
  assert.throws(() => api.search("   "), /must not be empty/);
  assert.throws(() => api.search("anything", { rowLimit: 501 }), /rowLimit/);
  assert.throws(() => api.search("anything", { startRow: 50_001 }), /startRow/);
});

test("read-only list permission helpers return bindings and inheritance metadata with the authenticated REST routes", async () => {
  const calls = [];
  const assignment = { PrincipalId: 7, Member: { Id: 7, Title: "Operations Owners", PrincipalType: 8 }, RoleDefinitionBindings: [{ Id: 1073741829, Name: "Full Control" }] };
  const api = createLegacySharePointApi({
    get: async (url) => {
      calls.push(url);
      return url.includes("roleassignments")
        ? { value: [assignment] }
        : { Id: listId, Title: "Operations", HasUniqueRoleAssignments: false };
    },
    request: async () => undefined,
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  });
  assert.deepEqual(await api.getListRoleAssignments(listId), [assignment]);
  assert.deepEqual(await api.getListPermissionScope(listId), { Id: listId, Title: "Operations", HasUniqueRoleAssignments: false });
  const assignmentUrl = new URL(calls[0], "https://contoso.sharepoint.com");
  assert.equal(assignmentUrl.pathname, `/_api/web/lists(guid'${listId}')/roleassignments`);
  assert.equal(assignmentUrl.searchParams.get("$top"), "100");
  assert.equal(assignmentUrl.searchParams.get("$expand"), "Member,RoleDefinitionBindings");
  assert.equal(calls[1], `/_api/web/lists(guid'${listId}')?$select=Id,Title,HasUniqueRoleAssignments`);
});

test("list role assignment reads follow safe same-endpoint pages and enforce a page cap", async () => {
  const next = `https://contoso.sharepoint.com/sites/team/_api/web/lists(guid'${listId}')/roleassignments?$skiptoken=next`;
  const calls = [];
  const api = createLegacySharePointApi({
    siteUrl: "https://contoso.sharepoint.com/sites/team",
    get: async (url) => {
      calls.push(url);
      return url.includes("$skiptoken")
        ? { value: [{ PrincipalId: 8, Member: { Id: 8, Title: "Readers" } }] }
        : { value: [{ PrincipalId: 7, Member: { Id: 7, Title: "Owners" } }], "@odata.nextLink": next };
    },
    request: async () => undefined,
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  });
  const assignments = await api.getListRoleAssignments(listId);
  assert.deepEqual(assignments.map((item) => item.PrincipalId), [7, 8]);
  assert.equal(calls.length, 2);
  assert.equal((await api.getListRoleAssignmentsPage(listId, { pageSize: 1 })).nextCursor, next);
  await assert.rejects(api.getListRoleAssignments(listId, { maxPages: 1 }), /exceeded maxPages/);
  await assert.rejects(api.getListRoleAssignments(listId, { maxPages: 1001 }), /maxPages/);
});

test("role assignment writes use SharePoint REST routes and require explicit confirmation", async () => {
  const calls = [];
  const signal = new AbortController().signal;
  const api = createLegacySharePointApi({
    get: async (url) => { calls.push({ method: "GET", url }); return { value: [{ Id: 1073741827, Name: "Read", RoleTypeKind: 2 }] }; },
    request: async (url, options) => { calls.push({ method: options?.method, url, options }); },
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  });
  assert.deepEqual((await api.getRoleDefinitionsPage()).items, [{ Id: 1073741827, Name: "Read", RoleTypeKind: 2 }]);
  assert.equal(calls[0].url, "/_api/web/roledefinitions?%24top=50&%24select=Id%2CName%2CDescription%2CRoleTypeKind");
  await assert.rejects(api.addListRoleAssignment(listId, 7, 1073741827, { confirmed: false }), /Confirm this SharePoint permission change/);
  await assert.rejects(api.breakListRoleInheritance(listId, { confirmed: false }), /Confirm this SharePoint permission change/);
  await assert.rejects(api.resetListRoleInheritance(listId, { confirmed: false }), /Confirm this SharePoint permission change/);
  await assert.rejects(api.removeListRoleAssignment(listId, 7, { confirmed: false }), /Confirm this SharePoint permission change/);
  assert.equal(calls.length, 1, "unconfirmed permission writes must not reach SharePoint");

  await api.breakListRoleInheritance(listId, { confirmed: true, signal });
  await api.addListRoleAssignment(listId, 7, 1073741827, { confirmed: true, signal });
  await api.removeListRoleAssignment(listId, 7, { confirmed: true, signal });
  await api.resetListRoleInheritance(listId, { confirmed: true, signal });
  assert.deepEqual(calls.slice(1).map(({ method, url }) => [method, url]), [
    ["POST", `/_api/web/lists(guid'${listId}')/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`],
    ["POST", `/_api/web/lists(guid'${listId}')/roleassignments/addroleassignment(principalid=7,roledefid=1073741827)`],
    ["POST", `/_api/web/lists(guid'${listId}')/roleassignments/getbyprincipalid(7)`],
    ["POST", `/_api/web/lists(guid'${listId}')/resetroleinheritance()`],
  ]);
  assert.deepEqual(calls[3].options.headers, { "X-HTTP-Method": "DELETE" });
  assert.equal(calls[2].options.signal, signal);
  await assert.rejects(api.addListRoleAssignment(listId, 0, 1, { confirmed: true }), /principal ID/);
  await assert.rejects(api.addListRoleAssignment(listId, 1, 0, { confirmed: true }), /role definition ID/);
});

test("list permission helpers reject unexpected response shapes", async () => {
  const api = createLegacySharePointApi({ get: async () => ({}), request: async () => undefined, raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }) });
  await assert.rejects(api.getListRoleAssignments(listId), /value collection/);
  await assert.rejects(api.getListPermissionScope(listId), /HasUniqueRoleAssignments/);
});

test("list permission reads reject list path injection instead of building an OData route", async () => {
  const api = createLegacySharePointApi({ get: async () => ({ value: [] }), request: async () => undefined, raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }) });
  await assert.rejects(api.getListRoleAssignments("x')/web/siteusers"), /GUIDs/);
});

test("navigation helper pages web and site bars and normalizes one level of children", async () => {
  const calls = [];
  const next = "https://contoso.sharepoint.com/sites/team/_api/web/Navigation/QuickLaunch?$skiptoken=next";
  const api = createLegacySharePointApi({
    siteUrl: "https://contoso.sharepoint.com/sites/team",
    get: async (url) => {
      calls.push(url);
      if (url.includes("$skiptoken")) return { d: { results: [{ Id: 4, Title: "Second", Url: "/sites/team/second", IsExternal: false, IsVisible: true }] } };
      if (url.includes("RootWeb")) return { value: [{ Id: 5, Title: "Root link", Url: "/sites/team/root" }] };
      if (url.includes("TopNavigationBar")) return { value: [] };
      return { d: { results: [{
        Id: 1, Title: "Documents", Url: "/sites/team/shared", IsExternal: false, IsVisible: true,
        Children: { results: [{ Id: 2, Title: "Reports", Url: "/sites/team/reports", IsExternal: false, IsVisible: true, Children: { __deferred: { uri: "ignored" } } }] },
      }], __next: next } };
    },
    request: async () => undefined,
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  });

  const first = await api.getNavigationNodesPage({ includeChildren: true, pageSize: 500 });
  assert.equal(first.items[0].Title, "Documents");
  assert.deepEqual(first.items[0].Children, [{ Id: 2, Title: "Reports", Url: "/sites/team/reports", IsExternal: false, IsVisible: true }]);
  assert.equal(new URL(calls[0]).searchParams.get("$top"), "100");
  assert.equal(new URL(calls[0]).searchParams.get("$expand"), "Children");
  assert.match(new URL(calls[0]).searchParams.get("$select"), /Children\/Title/);

  const second = await api.getNavigationNodesPage({ cursor: first.nextCursor, includeChildren: true });
  assert.deepEqual(second.items.map((node) => node.Title), ["Second"]);
  assert.equal(new URL(calls[1]).search, "?$skiptoken=next");

  const top = await api.getNavigationNodesPage({ bar: "topNavigationBar" });
  assert.deepEqual(top.items, []);
  assert.match(calls[2], /_api\/web\/Navigation\/TopNavigationBar/);
  const site = await api.getNavigationNodesPage({ scope: "site" });
  assert.equal(site.items[0].Title, "Root link");
  assert.match(calls[3], /_api\/site\/RootWeb\/Navigation\/QuickLaunch/);

  const noExpand = await api.getNavigationNodesPage();
  assert.equal(Object.hasOwn(noExpand.items[0], "Children"), false);

  await assert.rejects(api.getNavigationNodesPage({ scope: "tenant" }), /scope must be web or site/);
  await assert.rejects(api.getNavigationNodesPage({ bar: "quicklaunch" }), /bar must be quickLaunch or topNavigationBar/);
});

test("navigation helper rejects malformed collections, node IDs, and unsafe continuation links", async () => {
  const make = (get) => createLegacySharePointApi({ siteUrl: "https://contoso.sharepoint.com/sites/team", get, request: async () => undefined, raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }) });
  await assert.rejects(make(async () => ({ value: [{ Title: "No ID" }] })).getNavigationNodesPage(), /numeric Id/);
  await assert.rejects(make(async () => ({})).getNavigationNodesPage(), /collection response/);
  const api = make(async () => ({ value: [], "@odata.nextLink": "https://evil.example/_api/web/Navigation/QuickLaunch?skip=1" }));
  const page = await api.getNavigationNodesPage();
  await assert.rejects(api.getNavigationNodesPage({ cursor: page.nextCursor }), /same collection endpoint and web/);
});

test("site people and groups use bounded REST pages and preserve opaque same-endpoint continuation links", async () => {
  const calls = [];
  const rest = {
    siteUrl: "https://contoso.sharepoint.com/sites/team",
    get: async (url) => {
      calls.push(url);
      if (url.includes("siteusers")) return { value: [{ Id: 1, Title: "Ada" }], "@odata.nextLink": "https://contoso.sharepoint.com/sites/team/_api/web/siteusers?$skiptoken=next" };
      if (url.includes("sitegroups(9)")) return { d: { results: [{ Id: 2, Title: "Lin" }] } };
      return { value: [{ Id: 9, Title: "Owners" }] };
    },
    request: async () => undefined,
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  };
  const api = createLegacySharePointApi(rest);
  const users = await api.getSiteUsersPage({ pageSize: 500 });
  assert.deepEqual(users.items, [{ Id: 1, Title: "Ada" }]);
  assert.equal(new URL(calls[0]).searchParams.get("$top"), "100");
  const next = await api.getSiteUsersPage({ cursor: users.nextCursor });
  assert.deepEqual(next.items, [{ Id: 1, Title: "Ada" }]);
  const groups = await api.getSiteGroupsPage();
  const members = await api.getSiteGroupMembersPage(9);
  assert.deepEqual(groups.items, [{ Id: 9, Title: "Owners" }]);
  assert.deepEqual(members.items, [{ Id: 2, Title: "Lin" }]);
  assert.equal(new URL(calls[3]).pathname, "/sites/team/_api/web/sitegroups(9)/users");
  await assert.rejects(api.getSiteUsersPage({ cursor: "https://attacker.example/_api/web/siteusers?x=1" }), /same collection endpoint/);
  await assert.rejects(api.getSiteUsersPage({ cursor: "https://contoso.sharepoint.com/sites/team/_api/web/sitegroups?x=1" }), /same collection endpoint/);
  await assert.rejects(api.getSiteUsersPage({ pageSize: Number.NaN }), /positive integer/);
  assert.throws(() => api.getSiteGroupMembersPage(0), /positive integers/);
});

test("current user helper reads the authenticated hosting web user", async () => {
  let requested;
  const api = createLegacySharePointApi({
    get: async (url) => { requested = url; return { Id: 12, Title: "Ada", LoginName: "i:0#.f|membership|ada@example.com" }; },
    request: async () => undefined,
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  });
  assert.equal((await api.getCurrentUser()).Id, 12);
  assert.equal(requested, "/_api/web/currentuser?$select=Id,Title,Email,LoginName,PrincipalType,IsSiteAdmin");
});

test("PeopleManager profile reads encode account aliases and validate response shapes", async () => {
  const calls = [];
  const api = createLegacySharePointApi({
    get: async (url) => {
      calls.push(url);
      if (url.endsWith("GetMyProperties")) return { d: { AccountName: "i:0#.f|membership|ada@example.com", DisplayName: "Ada" } };
      if (url.includes("GetUserProfilePropertyFor")) return { d: "Engineering" };
      return { AccountName: "i:0#.f|membership|lee@example.com", UserProfileProperties: [{ Key: "Department", Value: "Research" }] };
    },
    request: async () => undefined,
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  });
  assert.equal((await api.getMyProperties()).DisplayName, "Ada");
  const profile = await api.getPropertiesFor("i:0#.f|membership|lee@example.com");
  assert.equal(profile.UserProfileProperties?.[0].Value, "Research");
  assert.equal(await api.getUserProfilePropertyFor("i:0#.f|membership|lee@example.com", "SPS-Department"), "Engineering");
  const alias = new URLSearchParams(new URL(calls[1], "https://contoso.sharepoint.com").search);
  assert.equal(alias.get("@v"), "'i:0#.f|membership|lee@example.com'");
  assert.match(decodeURIComponent(calls[2]), /propertyName='SPS-Department'/);
  await assert.rejects(api.getPropertiesFor("  "), /account name/);
  await assert.rejects(api.getUserProfilePropertyFor("ada", "bad/property"), /property name/);

  const malformed = createLegacySharePointApi({ get: async () => [], request: async () => undefined, raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }) });
  await assert.rejects(malformed.getMyProperties(), /must be an object/);
  await assert.rejects(malformed.getUserProfilePropertyFor("ada", "Title"), /must be a string/);
});

test("list view helpers use bounded subweb REST routes and validate their continuations", async () => {
  const calls = [];
  const view = { Id: viewId, Title: "Open tasks", ViewType: "HTML", DefaultView: true, ViewQuery: "<Where />", RowLimit: 50, Paged: true };
  const api = createLegacySharePointApi({
    siteUrl: "https://contoso.sharepoint.com/sites/team",
    get: async (url) => {
      calls.push(url);
      if (url.includes("/ViewFields")) return { Items: ["Title", "Status", "DueDate"] };
      if (url.includes("$skiptoken")) return { d: { results: [view] } };
      return { value: [view], "@odata.nextLink": "https://contoso.sharepoint.com/sites/team/_api/web/lists(guid'00000000-0000-0000-0000-000000000123')/views?$skiptoken=next" };
    },
    request: async () => undefined,
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  });

  const first = await api.getListViewsPage(listId, { pageSize: 500 });
  assert.deepEqual(first.items, [view]);
  assert.equal(new URL(calls[0]).pathname, `/sites/team/_api/web/lists(guid'${listId}')/views`);
  assert.equal(new URL(calls[0]).searchParams.get("$top"), "100");
  assert.equal(new URL(calls[0]).searchParams.get("$select"), "Id,Title,ViewType,PersonalView,DefaultView,Hidden,ServerRelativeUrl,ViewQuery,RowLimit,Paged");
  const next = await api.getListViewsPage(listId, { cursor: first.nextCursor });
  assert.deepEqual(next.items, [view]);
  assert.deepEqual(await api.getListViewFields(listId, viewId), ["Title", "Status", "DueDate"]);
  assert.equal(calls[2], `/_api/web/lists(guid'${listId}')/views(guid'${viewId}')/ViewFields`);
  await assert.rejects(api.getListViewsPage(listId, { cursor: "https://attacker.example/_api/web/lists/views?x=1" }), /same collection endpoint/);
  await assert.rejects(api.getListViewsPage(listId, { pageSize: 0 }), /positive integer/);
  await assert.rejects(api.getListViewFields(listId, "not-a-guid"), /view IDs must be GUIDs/);

  const verbose = createLegacySharePointApi({
    get: async () => ({ d: { Items: ["Title", "ID"] } }),
    request: async () => undefined,
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  });
  assert.deepEqual(await verbose.getListViewFields(listId, viewId), ["Title", "ID"]);
  const malformed = createLegacySharePointApi({
    get: async () => ({ Items: ["Title", 42] }),
    request: async () => undefined,
    raw: async () => ({ status: 204, json: async () => undefined, blob: async () => new Blob() }),
  });
  await assert.rejects(malformed.getListViewFields(listId, viewId), /array of field names/);
});
