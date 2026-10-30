#!/usr/bin/env node
// Live, disposable checks for legacy SharePoint list REST endpoints.
import { existsSync } from "node:fs";
import { resolve } from "node:path";
import { chromium } from "playwright";
import { acquireProfileLock } from "../packages/playwright/src/profile-lock.js";

const options = {};
for (let i = 2; i < process.argv.length; i++) {
  if (process.argv[i] === "--help" || process.argv[i] === "-h") {
    console.log("Usage: node scripts/sharepoint-legacy-api-batch.mjs --site <site-url> --profile-dir <signed-in-profile>");
    process.exit(0);
  }
  const key = process.argv[i].replace(/^--/, "");
  if (!key || !process.argv[i].startsWith("--") || !process.argv[i + 1] || process.argv[i + 1].startsWith("--")) throw new Error("Expected --site and --profile-dir options.");
  options[key] = process.argv[++i];
}
if (!options.site || !options["profile-dir"]) throw new Error("Expected --site and --profile-dir options.");
const site = new URL(options.site);
if (site.protocol !== "https:") throw new Error("The SharePoint site must use HTTPS.");
const profileDir = resolve(options["profile-dir"]);
if (!existsSync(profileDir)) throw new Error("The signed-in browser profile does not exist.");

const releaseProfile = acquireProfileLock(profileDir);
let context;
try {
  context = await chromium.launchPersistentContext(profileDir, { headless: true });
  const page = context.pages()[0] ?? await context.newPage();
  await page.goto(site.href, { waitUntil: "domcontentloaded", timeout: 60000 });
  if (page.url().includes("login.microsoftonline.com")) throw new Error("The browser profile is signed out.");
  const results = await page.evaluate(async ({ sitePath }) => {
    const api = `${sitePath.replace(/\/$/, "")}/_api`;
    const accept = "application/json;odata=verbose";
    const rows = [];
    let listId;
let invalidTokenListId;
    let digest;
    const record = (name, detail = "ok") => rows.push({ name, status: "pass", detail });
    const fail = (name, error) => rows.push({ name, status: "fail", detail: String(error?.message ?? error).slice(0, 180) });
    const request = async (path, { method = "GET", body, headers = {} } = {}) => {
      const h = { Accept: accept, ...headers };
      if (body !== undefined) h["Content-Type"] ??= "application/json;odata=verbose";
      if (method !== "GET" && method !== "HEAD") h["X-RequestDigest"] ??= digest;
      const response = await fetch(`${api}${path}`, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body) });
      if (!response.ok) {
        const bodyText = await response.text().catch(() => "");
        let detail = "";
        try {
          const payload = JSON.parse(bodyText);
          const apiError = payload?.error ?? payload?.["odata.error"] ?? payload?.d?.error;
          detail = apiError?.message?.value ?? apiError?.message ?? apiError?.code ?? "";
        } catch { /* Keep server response bodies out of logs. */ }
        if (!detail && bodyText) detail = bodyText.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").replace(/[0-9a-f]{8}-[0-9a-f-]{27,}/gi, "[id]").trim();
        const error = new Error(`HTTP ${response.status}${detail ? `: ${String(detail).slice(0, 140)}` : ""}`);
        error.status = response.status;
        throw error;
      }
      if (response.status === 204) return undefined;
      const text = await response.text();
      if (!text) return undefined;
      if ((response.headers.get("content-type") ?? "").includes("xml") || text.trimStart().startsWith("<?xml")) return text;
      try { return JSON.parse(text); } catch { return text; }
    };
    const entity = (response) => response?.d ?? response;
    const collection = (response) => entity(response)?.results ?? response?.value ?? [];
    const listTitle = `SPFxFastLegacyBatch_${crypto.randomUUID().slice(0, 8)}`;
    try {
      const contextInfo = await request("/contextinfo", { method: "POST" });
      digest = entity(contextInfo)?.GetContextWebInformation?.FormDigestValue;
      if (!digest) throw new Error("contextinfo did not return a request digest");
      record("Request digest");

      const webNavigation = collection(await request("/web/Navigation/QuickLaunch?$select=Id,Title,Url,IsExternal,IsVisible,Children/Id,Children/Title,Children/Url,Children/IsExternal,Children/IsVisible&$expand=Children&$top=100"));
      let navigationChildren = 0;
      for (const node of webNavigation) {
        if (!Number.isSafeInteger(node?.Id)) throw new Error("Quick Launch navigation returned a node without a numeric Id");
        if (node.Children != null) {
          const children = node.Children.results ?? node.Children.value;
          if (!Array.isArray(children)) throw new Error("Quick Launch child expansion returned an unexpected collection shape");
          navigationChildren += children.length;
        }
      }
      record("Web Quick Launch navigation", `${webNavigation.length} nodes; ${navigationChildren} one-level children`);

      const topNavigation = collection(await request("/web/Navigation/TopNavigationBar?$select=Id,Title,Url,IsExternal,IsVisible&$top=100"));
      if (topNavigation.some((node) => !Number.isSafeInteger(node?.Id))) throw new Error("Top navigation returned a node without a numeric Id");
      record("Web top navigation", `${topNavigation.length} nodes`);

      const siteNavigation = collection(await request("/site/RootWeb/Navigation/QuickLaunch?$select=Id,Title,Url,IsExternal,IsVisible&$top=100"));
      if (siteNavigation.some((node) => !Number.isSafeInteger(node?.Id))) throw new Error("Site-root Quick Launch returned a node without a numeric Id");
      record("Site-root Quick Launch", `${siteNavigation.length} nodes`);

      const created = entity(await request("/web/lists", { method: "POST", body: {
        __metadata: { type: "SP.List" }, BaseTemplate: 100, Title: listTitle,
      } }));
      listId = created?.Id;
      if (!listId) throw new Error("temporary list creation returned no ID");
      record("Create disposable list");

      const list = entity(await request(`/web/lists(guid'${listId}')?$select=Id,Title,ListItemEntityTypeFullName,EnableAudienceTargeting`));
      const listType = list.ListItemEntityTypeFullName;
      if (!listType) throw new Error("list metadata omitted ListItemEntityTypeFullName");
      record("Read list metadata");

      const synthetic = `SPFx legacy API batch ${crypto.randomUUID()}`;
      for (let index = 0; index < 3; index++) await request(`/web/lists(guid'${listId}')/items`, { method: "POST", body: { __metadata: { type: listType }, Title: `${synthetic} ${index + 1}` } });
      record("Create 3 synthetic items");

      const viewXml = `<View><Query><Where><BeginsWith><FieldRef Name="Title"/><Value Type="Text">${synthetic}</Value></BeginsWith></Where><OrderBy><FieldRef Name="ID" Ascending="TRUE"/></OrderBy></Query><ViewFields><FieldRef Name="ID"/><FieldRef Name="Title"/></ViewFields><RowLimit>10</RowLimit></View>`;
      const camlResponse = await request(`/web/lists(guid'${listId}')/GetItems`, {
        method: "POST",
        body: { query: { __metadata: { type: "SP.CamlQuery" }, ViewXml: viewXml } },
      });
      const caml = entity(camlResponse);
      const camlItems = collection(caml);
      if (camlItems.length !== 3) throw new Error(`CAML returned ${camlItems.length} rows; expected 3`);
      record("CAML GetItems", `${camlItems.length} rows`);

      const streamResponse = await request(`/web/lists(guid'${listId}')/RenderListDataAsStream`, {
        method: "POST",
        body: { parameters: { ViewXml: viewXml } },
      });
      const stream = entity(streamResponse);
      const streamRows = stream?.ListData?.Row ?? stream?.Row ?? [];
      if (!streamRows.length) throw new Error("RenderListDataAsStream returned no rows");
      record("RenderListDataAsStream", `${streamRows.length} rows`);

      let changedRows = 0;
      let changePages = 0;
      let changeTokenFound = false;
      let lastChangeToken;
      let pagingPosition;
      do {
        if (++changePages > 10) throw new Error("change-token paging exceeded 10 pages");
        const queryOptions = pagingPosition
          ? `<QueryOptions><Paging ListItemCollectionPositionNext="${pagingPosition}" /></QueryOptions>`
          : "<QueryOptions><DateInUtc>TRUE</DateInUtc><IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns></QueryOptions>";
        const changesXml = await request(`/web/lists(guid'${listId}')/GetListItemChangesSinceToken`, { method: "POST", headers: { Accept: "application/xml" }, body: { query: {
          __metadata: { type: "SP.ChangeLogItemQuery" }, ChangeToken: null, ViewName: "",
          Query: "<Query><OrderBy><FieldRef Name=\"ID\" Ascending=\"TRUE\"/></OrderBy></Query>",
          ViewFields: "<ViewFields><FieldRef Name=\"ID\"/><FieldRef Name=\"Title\"/></ViewFields>", RowLimit: "1",
          QueryOptions: queryOptions,
        } } });
        changedRows += (changesXml.match(/<z:row\b/g) ?? []).length;
        const tokenMatch = changesXml.match(/LastChangeToken="([^"]+)"/i);
        if (tokenMatch) lastChangeToken = tokenMatch[1];
        changeTokenFound ||= Boolean(tokenMatch);
        pagingPosition = changesXml.match(/ListItemCollectionPositionNext="([^"]+)"/i)?.[1];
      } while (pagingPosition);
      if (changedRows !== 3 || !changeTokenFound) throw new Error(`change-token paging returned ${changedRows} rows; expected 3 and a token`);
      record("GetListItemChangesSinceToken", `${changedRows} rows across ${changePages} page(s); token present`);

      const invalidList = entity(await request("/web/lists", { method: "POST", body: {
        __metadata: { type: "SP.List" }, BaseTemplate: 100, Title: `SPFxFastInvalidToken_${crypto.randomUUID().slice(0, 8)}`,
      } }));
      invalidTokenListId = invalidList?.Id;
      if (!invalidTokenListId || !lastChangeToken) throw new Error("Could not prepare the invalid change-token probe.");
      const invalidListType = entity(await request(`/web/lists(guid'${invalidTokenListId}')?$select=ListItemEntityTypeFullName`))?.ListItemEntityTypeFullName;
      await request(`/web/lists(guid'${invalidTokenListId}')/items`, { method: "POST", body: {
        __metadata: { type: invalidListType }, Title: `SPFx Fast invalid-token ${crypto.randomUUID()}`,
      } });
      const invalidTokenXml = await request(`/web/lists(guid'${invalidTokenListId}')/GetListItemChangesSinceToken`, { method: "POST", headers: { Accept: "application/xml" }, body: { query: {
        __metadata: { type: "SP.ChangeLogItemQuery" }, ChangeToken: lastChangeToken, ViewName: "",
        Query: "<Query />", ViewFields: "<ViewFields><FieldRef Name=\"ID\"/><FieldRef Name=\"Title\"/></ViewFields>", RowLimit: "20",
        QueryOptions: "<QueryOptions><DateInUtc>TRUE</DateInUtc></QueryOptions>",
      } } });
      if (!/<Id\b[^>]*ChangeType=["']InvalidToken["']/i.test(invalidTokenXml)) throw new Error("SharePoint did not return the expected InvalidToken XML marker.");
      record("Invalid change-token response", "A token from another disposable list returned the InvalidToken marker used by the reset helper.");

      const changes = entity(await request(`/web/lists(guid'${listId}')/GetChanges`, { method: "POST", body: { query: {
        __metadata: { type: "SP.ChangeQuery" }, Item: true, Add: true, Update: true, DeleteObject: true, FetchLimit: "100",
      } } }));
      const changeRows = collection(changes);
      if (!Array.isArray(changeRows)) throw new Error("GetChanges returned an unexpected shape");
      record("GetChanges", `${changeRows.length} entries`);

      const views = collection(await request(`/web/lists(guid'${listId}')/views?$select=Id,Title,DefaultView&$top=10`));
      if (!views.length) throw new Error("the temporary list returned no views");
      record("List views", `${views.length} views`);

      const fieldPath = `/web/lists(guid'${listId}')/fields`;
      const audienceField = "Target_x0020_Audiences";
      const existingFields = collection(await request(`${fieldPath}?$select=Id,InternalName,TypeAsString&$filter=InternalName eq '${audienceField}'`));
      if (existingFields.length) throw new Error("temporary list unexpectedly already has an audience field");
      const schema = `<Field ID="{61CBB965-1E04-4273-B658-EEDAA662F48D}" Type="TargetTo" Name="${audienceField}" StaticName="${audienceField}" DisplayName="Target Audiences" Required="FALSE" />`;
      const createdField = entity(await request(`${fieldPath}/CreateFieldAsXml`, { method: "POST", body: { parameters: {
        __metadata: { type: "SP.XmlSchemaFieldCreationInformation" }, SchemaXml: schema, Options: 0,
      } } }));
      const enabledFields = collection(await request(`${fieldPath}?$select=Id,InternalName,TypeAsString&$filter=InternalName eq '${audienceField}'`));
      if (enabledFields.length !== 1 || enabledFields[0].TypeAsString !== "TargetTo") throw new Error("Target Audiences field was not created as TargetTo");
      record("Audience targeting enable via TargetTo field");
      const createdId = createdField?.Id ?? enabledFields[0].Id;
      await request(`${fieldPath}(guid'${createdId}')`, { method: "DELETE", headers: { "If-Match": "*" } });
      const disabledFields = collection(await request(`${fieldPath}?$select=Id&$filter=InternalName eq '${audienceField}'`));
      if (disabledFields.length) throw new Error("Target Audiences field remained after disable");
      record("Audience targeting disable and read-back");
    } catch (error) {
      fail("Batch stopped", error);
    } finally {
      if (invalidTokenListId) {
        try { await request(`/web/lists(guid'${invalidTokenListId}')`, { method: "DELETE", headers: { "If-Match": "*" } }); record("Cleanup invalid-token list"); }
        catch (error) { fail("Cleanup invalid-token list", error); }
      }
      if (listId) {
        try { await request(`/web/lists(guid'${listId}')`, { method: "DELETE", headers: { "If-Match": "*" } }); record("Cleanup temporary list"); }
        catch (error) { fail("Cleanup temporary list", error); }
      }
    }
    return rows;
  }, { sitePath: site.pathname });
  const passed = results.filter((result) => result.status === "pass").length;
  const failed = results.filter((result) => result.status === "fail");
  console.log(JSON.stringify({ passed, failed: failed.length, results }, null, 2));
  if (failed.length) process.exitCode = 1;
} finally {
  await context?.close();
  releaseProfile();
}
