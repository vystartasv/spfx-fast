import assert from "node:assert/strict";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import test from "node:test";

const capabilities = JSON.parse(readFileSync(new URL("../agent/capabilities.json", import.meta.url), "utf8"));
const packageJson = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8"));
const { getMissingRequiredFieldNames, toGraphFieldValues } = await import("../dist/react.js");

function markdownFiles(directory) {
  const skipped = new Set([".git", "node_modules", "dist", "lib", "lib-dts", "temp", "test-results", ".heft"]);
  const files = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const path = resolve(directory, entry.name);
    if (entry.isDirectory()) {
      if (!skipped.has(entry.name)) files.push(...markdownFiles(path));
    } else if (entry.name.endsWith(".md")) files.push(path);
  }
  return files;
}

function markdownAnchors(contents) {
  const counts = new Map();
  return [...contents.matchAll(/^#{1,6} (.+)$/gm)].map(([, heading]) => {
    const base = heading.toLowerCase().replace(/[^a-z0-9 -]/g, "").trim().replace(/\s+/g, "-");
    const count = counts.get(base) ?? 0;
    counts.set(base, count + 1);
    return count ? `${base}-${count}` : base;
  });
}

test("multi-choice fields treat an empty selection as missing", () => {
  const fields = [{ name: "Categories", title: "Categories", type: "multiChoice", required: true, choices: ["A", "B"] }];
  assert.deepEqual(getMissingRequiredFieldNames(fields, { Categories: [] }), ["Categories"]);
  assert.deepEqual(getMissingRequiredFieldNames(fields, { Categories: ["A", "B"] }), []);
});

test("Person and Lookup form values use Graph LookupId field names", () => {
  const fields = [
    { name: "Owner", title: "Owner", type: "person" },
    { name: "Project", title: "Project", type: "lookup" },
    { name: "Title", title: "Title", type: "text" },
  ];
  assert.deepEqual(toGraphFieldValues(fields, { Owner: 13, Project: "42", Title: "Request" }), {
    OwnerLookupId: 13,
    ProjectLookupId: "42",
    Title: "Request",
  });
});

test("agent capability names are unique and every capability has a status", () => {
  const names = capabilities.capabilities.map((entry) => entry.name);
  assert.equal(new Set(names).size, names.length);
  assert.ok(capabilities.capabilities.every((entry) => typeof entry.status === "string"));
});

test("every advertised capability links to a shipped page and valid section", () => {
  const names = capabilities.capabilities.map((entry) => entry.name).sort();
  const links = capabilities.capabilityDocs ?? {};
  assert.deepEqual(Object.keys(links).sort(), names, "each capability must have exactly one documentation target");
  for (const [name, target] of Object.entries(links)) {
    const [path, anchor] = target.split("#");
    assert.ok(path && anchor, `${name} documentation target must include a page and anchor`);
    const contents = readFileSync(new URL(`../${path}`, import.meta.url), "utf8");
    const slugs = [...contents.matchAll(/^#{1,6} (.+)$/gm)].map(([, heading]) => heading.toLowerCase().replace(/[^a-z0-9 -]/g, "").trim().replace(/\s+/g, "-"));
    assert.ok(slugs.includes(anchor), `${name} points to missing heading ${target}`);
  }
});

test("local Markdown links and heading anchors resolve", () => {
  const root = resolve(new URL("..", import.meta.url).pathname);
  const files = markdownFiles(root);
  for (const file of files) {
    const contents = readFileSync(file, "utf8");
    for (const [, destination] of contents.matchAll(/\[[^\]]*\]\(([^)]+)\)/g)) {
      const target = destination.trim().split(/\s+/)[0].replace(/^<|>$/g, "");
      if (!target || /^(https?:|mailto:|data:)/i.test(target)) continue;
      const [relativePath, rawAnchor] = target.split("#");
      const targetFile = relativePath ? resolve(dirname(file), decodeURIComponent(relativePath)) : file;
      assert.ok(existsSync(targetFile), `${file.slice(root.length + 1)} links to missing ${target}`);
      if (rawAnchor) {
        const anchors = markdownAnchors(readFileSync(targetFile, "utf8"));
        assert.ok(anchors.includes(decodeURIComponent(rawAnchor)), `${file.slice(root.length + 1)} links to missing anchor ${target}`);
      }
    }
  }
});

test("agent recipe catalog exists and its declared recipe count matches", () => {
  const recipes = readFileSync(new URL("../docs/EXAMPLES.md", import.meta.url), "utf8");
  assert.equal(capabilities.exampleCatalog.recipeCount, (recipes.match(/^\|\s*\d+\s*\|/gm) ?? []).length);
  const readme = readFileSync(new URL("../README.md", import.meta.url), "utf8");
  assert.match(readme, new RegExp(`\\[${capabilities.exampleCatalog.recipeCount} copyable recipes\\]`));
  assert.match(readme, new RegExp(`\\*\\*Ready:\\*\\*[^\\n]*${capabilities.exampleCatalog.recipeCount} recipes`));
});

test("business scenario presets are unique, typed, and satisfy the generated form contract", () => {
  const catalog = JSON.parse(readFileSync(new URL("../examples/scenarios.json", import.meta.url), "utf8"));
  assert.equal(catalog.scenarios.length, capabilities.exampleCatalog.scenarioCount);
  const ids = catalog.scenarios.map(({ id }) => id);
  assert.equal(new Set(ids).size, ids.length);
  for (const scenario of catalog.scenarios) {
    assert.ok(scenario.id && scenario.name && scenario.description, `scenario ${scenario.id} needs discoverable metadata`);
    assert.ok(scenario.fields.some((field) => field.name === "Title"), `${scenario.id} needs the standard Title field`);
    for (const field of scenario.fields) {
      assert.ok(/^[A-Za-z0-9_]{1,128}$/.test(field.name), `${scenario.id}.${field.name} needs a safe internal field name`);
      assert.ok(["text", "multiline", "number", "date", "boolean", "choice", "multiChoice", "person", "lookup"].includes(field.type), `${scenario.id}.${field.name} has unsupported form type`);
      if (field.type === "choice" || field.type === "multiChoice") assert.ok(Array.isArray(field.choices) && field.choices.length > 0, `${scenario.id}.${field.name} needs choices`);
    }
  }
});

test("100 use cases map to valid app presets and flow patterns and are documented", () => {
  const scenarios = JSON.parse(readFileSync(new URL("../examples/scenarios.json", import.meta.url), "utf8"));
  const catalog = JSON.parse(readFileSync(new URL("../examples/use-cases.json", import.meta.url), "utf8"));
  const useCases = catalog.categories.flatMap((category) => category.useCases);
  const ids = useCases.map(({ id }) => id);
  assert.equal(useCases.length, capabilities.exampleCatalog.useCaseCount);
  assert.equal(new Set(ids).size, ids.length);
  const presetIds = new Set(scenarios.scenarios.map(({ id }) => id));
  const flowIds = new Set(catalog.flowPatterns.map(({ id }) => id));
  const docs = readFileSync(new URL("../docs/USE-CASES.md", import.meta.url), "utf8");
  for (const item of useCases) {
    assert.ok(presetIds.has(item.preset), `${item.id} references missing app preset ${item.preset}`);
    assert.ok(flowIds.has(item.flow), `${item.id} references missing flow pattern ${item.flow}`);
    assert.ok(docs.includes(item.id), `${item.id} is absent from the human-readable use-case catalog`);
  }
});

test("newly promoted APIs have numbered recipes and matching sections", () => {
  const recipes = readFileSync(new URL("../docs/EXAMPLES.md", import.meta.url), "utf8");
  const headingSlugs = [...recipes.matchAll(/^## (.+)$/gm)].map(([, heading]) => heading.toLowerCase().replace(/[^a-z0-9 -]/g, "").trim().replace(/\s+/g, "-"));
  for (const [capability, anchor] of [
    ["client.retry-policy", "configure-sharepoint-read-retries"],
    ["client.sync-policy", "keep-local-list-data-fresh"],
    ["item.delete", "delete-after-explicit-confirmation"],
    ["sharepoint.renderListDataAsStream", "render-a-sharepoint-list-view"],
    ["sharepoint.files-folders", "read-a-sharepoint-file-with-rest"],
    ["sharepoint.views-content-types", "inspect-sharepoint-list-views"],
    ["sharepoint.user-profiles", "read-sharepoint-user-profile-properties"],
    ["sharepoint.versions", "read-list-and-file-version-history"],
    ["sharepoint.drive-items", "sync-a-document-library-with-delta"],
    ["sharepoint.drive-items", "create-move-and-delete-drive-items"],
  ]) {
    assert.ok(recipes.split("\n").some((line) => line.startsWith("|") && line.includes(`#${anchor}`) && line.includes(`\`${capability}\``)), `${capability} must have a numbered recipe pointing to #${anchor}`);
    assert.ok(headingSlugs.includes(anchor), `recipe section #${anchor} must exist`);
  }
});

test("package and generated UI target Fluent UI React v9 only", () => {
  assert.equal(packageJson.peerDependencies["@fluentui/react-components"].startsWith("^9."), true);
  assert.deepEqual(Object.keys(packageJson.peerDependencies).sort(), ["@fluentui/react-components", "@types/react", "react"]);
  assert.ok(Object.values(packageJson.peerDependenciesMeta).every((peer) => peer.optional));
  assert.equal(capabilities.defaults.ui.majorVersion, 9);
});

test("SPFx host fixture targets 1.24.0 with Fluent UI v9 and least declared Graph access", () => {
  const fixtureRoot = new URL("../examples/fixtures/spfx-fast-dashboard/", import.meta.url);
  const fixturePackage = JSON.parse(readFileSync(new URL("package.json", fixtureRoot), "utf8"));
  const solution = JSON.parse(readFileSync(new URL("config/package-solution.json", fixtureRoot), "utf8"));
  assert.equal(fixturePackage.dependencies["@microsoft/sp-webpart-base"], "1.24.0");
  assert.equal(fixturePackage.dependencies["@fluentui/react-components"].startsWith("^9."), true);
  assert.equal(fixturePackage.dependencies["@spfx-fast/core"], "file:../../..");
  assert.equal(fixturePackage.dependencies["@fluentui/react"], undefined);
  assert.ok(existsSync(new URL("src/webparts/spfxFastDashboard/components/DocumentLibraryPanel.tsx", fixtureRoot)));
  assert.ok(existsSync(new URL("src/webparts/spfxFastDashboard/components/CalendarAgenda.tsx", fixtureRoot)));
  assert.ok(existsSync(new URL("src/webparts/spfxFastDashboard/components/PerformanceProbe.tsx", fixtureRoot)));
  assert.deepEqual(solution.solution.webApiPermissionRequests, [
    { resource: "Microsoft Graph", scope: "Sites.ReadWrite.All" },
    { resource: "Microsoft Graph", scope: "Calendars.ReadBasic" },
  ]);
});

test("Copilot UX fixture targets SPFx 1.24.0 and Fluent UI v9 only", () => {
  const fixtureRoot = new URL("../examples/fixtures/spfx-fast-copilot/", import.meta.url);
  const fixturePackage = JSON.parse(readFileSync(new URL("package.json", fixtureRoot), "utf8"));
  const solution = JSON.parse(readFileSync(new URL("config/package-solution.json", fixtureRoot), "utf8"));
  const serve = JSON.parse(readFileSync(new URL("config/serve.json", fixtureRoot), "utf8"));
  const manifest = JSON.parse(readFileSync(new URL("copilot/declarativeAgent.json", fixtureRoot), "utf8"));
  const readme = readFileSync(new URL("README.md", fixtureRoot), "utf8");
  assert.equal(fixturePackage.dependencies["@microsoft/sp-copilot-component"], "1.24.0");
  assert.equal(fixturePackage.dependencies["@fluentui/react-components"].startsWith("^9."), true);
  assert.equal(fixturePackage.dependencies["@fluentui/react"], undefined);
  assert.equal(fixturePackage.dependencies.zod, "4.6.7");
  assert.equal(fixturePackage.dependencies["zod-to-json-schema"], "3.25.2");
  assert.equal(fixturePackage.dependencies["@spfx-fast/core"], "file:../../..");
  assert.deepEqual(solution.solution.webApiPermissionRequests, [{ resource: "Microsoft Graph", scope: "Sites.Read.All" }]);
  assert.equal(serve.initialPage, "https://{tenantDomain}/_layouts/15/copilotworkbench.aspx");
  assert.match(readme, /SPFX_SERVE_TENANT_DOMAIN=contoso\.sharepoint\.com pnpm start --nobrowser/);
  assert.match(readme, /SPFX_SERVE_TENANT_DOMAIN=contoso\.sharepoint\.com npm start -- --nobrowser/);
  assert.equal(manifest.version, "v1.8");
  assert.ok(existsSync(new URL("copilot/instruction.txt", fixtureRoot)));
});

test("package manifest is publishable and export targets exist", () => {
  assert.equal(packageJson.private, undefined);
  assert.equal(packageJson.publishConfig.access, "public");
  assert.equal(packageJson.license, "MIT");
  assert.ok(packageJson.files.includes("docs"));
  for (const entry of Object.values(packageJson.exports)) {
    for (const target of typeof entry === "string" ? [entry] : Object.values(entry)) {
      assert.ok(existsSync(new URL(`../${target.replace(/^\.\//, "")}`, import.meta.url)), `${target} exists`);
    }
  }
});

test("SPFx authenticated service factory is exported with its supporting adapters", async () => {
  const spfx = await import("../dist/spfx.js");
  assert.equal(typeof spfx.createSPFxServices, "function");
  const graph = { api: () => ({ get: async () => ({ value: [] }), post: async () => ({}), put: async () => ({}), patch: async () => ({}), delete: async () => undefined }) };
  const restRequests = [];
  const restClient = { fetch: async (url) => { restRequests.push(url); return { ok: true, status: 200, headers: { get: () => null }, json: async () => ({ Id: 12, Title: "Ada" }), text: async () => "", blob: async () => new Blob() }; } };
  const services = await spfx.createSPFxServices({
    msGraphClientFactory: { getClient: async () => graph },
    spHttpClient: restClient,
    pageContext: { site: { absoluteUrl: "https://contoso.sharepoint.com/sites/team" }, web: { absoluteUrl: "https://contoso.sharepoint.com/sites/team" }, user: { loginName: "user@contoso.com" } },
  }, {});
  assert.equal(services.scope, "contoso.sharepoint.com:user@contoso.com");
  assert.equal(typeof services.client.getLists, "function");
  assert.equal(typeof services.sharepoint.setAudienceTargeting, "function");
  assert.equal((await services.sharepoint.getCurrentUser()).Id, 12);
  assert.equal(new URL(restRequests[0]).pathname, "/sites/team/_api/web/currentuser");
  assert.equal(typeof services.sharepoint.getSiteUsersPage, "function");
});

test("SPFx services poll same-origin drive-copy operations and surface progress", async () => {
  const spfx = await import("../dist/spfx.js");
  const states = [
    { status: "inProgress", percentComplete: 35 },
    { status: "completed", percentComplete: 100, resourceId: "copied-file" },
  ];
  const requests = [];
  const graph = { api: () => ({ get: async () => ({}), post: async () => ({}), put: async () => ({}), patch: async () => ({}), delete: async () => undefined }) };
  const restClient = {
    fetch: async (url) => {
      requests.push(url);
      return { ok: true, status: 200, headers: { get: () => null }, json: async () => states.shift(), text: async () => "", blob: async () => new Blob() };
    },
  };
  const services = await spfx.createSPFxServices({
    msGraphClientFactory: { getClient: async () => graph },
    spHttpClient: restClient,
    pageContext: { site: { absoluteUrl: "https://contoso.sharepoint.com/sites/team" }, web: { absoluteUrl: "https://contoso.sharepoint.com/sites/team" }, user: { loginName: "user@contoso.com" } },
  }, {});
  const progress = [];
  const result = await services.waitForDriveCopy({ monitorUrl: "https://contoso.sharepoint.com/_api/v2.0/monitor/job" }, {
    pollIntervalMs: 50, timeoutMs: 1000, onProgress: (status) => progress.push(status.percentComplete),
  });
  assert.equal(result.resourceId, "copied-file");
  assert.deepEqual(progress, [35, 100]);
  assert.equal(requests.length, 2);
  assert.equal(new URL(requests[0]).pathname, "/_api/v2.0/monitor/job");
  await assert.rejects(services.waitForDriveCopy({ monitorUrl: "https://attacker.example/_api/v2.0/monitor/job" }), /same-origin/);
  await assert.rejects(services.waitForDriveCopy({ monitorUrl: "https://user@contoso.sharepoint.com/_api/v2.0/monitor/job" }), /same-origin/);
});

test("React entrypoint exports cached file and server-filtered list hooks", async () => {
  const react = await import("../dist/react.js");
  assert.equal(typeof react.useMarkdownFile, "function");
  assert.equal(typeof react.useListItems, "function");
  assert.equal(typeof react.DriveFloorPlanMap, "function");
  assert.equal(typeof react.FloorPlanMap, "function");
  assert.deepEqual(react.getMissingRequiredFieldNames([
    { name: "Title", title: "Title", required: true },
    { name: "Priority", title: "Priority", type: "choice", required: true },
    { name: "Enabled", title: "Enabled", type: "boolean", required: true },
    { name: "Count", title: "Count", type: "number", required: true },
  ], { Title: "   ", Priority: "", Enabled: false, Count: 0 }), ["Title", "Priority"]);
});

test("custom element entrypoint stays import-safe outside browsers", async () => {
  const elements = await import("../dist/elements.js");
  assert.equal(typeof elements.defineSPFxFastListElement, "function");
  assert.throws(() => elements.defineSPFxFastListElement(), /available only in a browser/);
});
