#!/usr/bin/env node
import { existsSync, mkdirSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";
import { acquireProfileLock } from "../packages/playwright/src/profile-lock.js";

const help = `Usage:
  node scripts/sharepoint-validation-page.mjs --site <site-url> --profile-dir <signed-in-profile> --task-list-id <guid> --project-list-id <guid> [--folder-item-id root] [--artifacts test-results/spfx-fast-validation]

Creates and publishes one SPFx Fast dashboard validation page, then reads it back and checks that the web part, delta check, and current bundle load. Re-running updates nothing when the page already has the requested configuration.`;
const args = process.argv.slice(2);
if (args.includes("--help") || args.includes("-h")) { console.log(help); process.exit(0); }
const options = {};
for (let i = 0; i < args.length; i++) {
  if (!args[i].startsWith("--") || !args[i + 1] || args[i + 1].startsWith("--")) throw new Error(`Invalid option: ${args[i]}\n${help}`);
  options[args[i].slice(2)] = args[++i];
}
for (const key of ["site", "profile-dir", "task-list-id", "project-list-id"]) if (!options[key]) throw new Error(`Missing --${key}\n${help}`);
const site = new URL(options.site);
if (site.protocol !== "https:") throw new Error("--site must use HTTPS.");
const apiRoot = `${site.pathname.replace(/\/+$/, "")}/_api`;
const profileDir = resolve(options["profile-dir"]);
if (!existsSync(profileDir)) throw new Error("The signed-in browser profile does not exist.");
const listIdPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
for (const key of ["task-list-id", "project-list-id"]) if (!listIdPattern.test(options[key])) throw new Error(`--${key} must be a SharePoint list GUID.`);
const componentId = "7a1ec127-4cf7-4dfe-872a-0ab52f0cb56c";
const manifestPath = resolve(fileURLToPath(new URL(`../examples/fixtures/spfx-fast-dashboard/dist/${componentId}.manifest.json`, import.meta.url)));
if (!existsSync(manifestPath)) throw new Error("Build the dashboard fixture first with pnpm run test:spfx-fixture.");
const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
const expectedBundle = manifest.loaderConfig?.scriptResources?.["spfx-fast-dashboard-web-part"]?.path;
if (!expectedBundle) throw new Error("The built dashboard manifest does not name its web-part bundle.");
const title = "SPFx Fast Validation";
const fileName = "SPFxFastValidation.aspx";
const properties = {
  taskListId: options["task-list-id"],
  projectListId: options["project-list-id"],
  folderItemId: options["folder-item-id"] ?? "root",
};
const artifacts = resolve(options.artifacts ?? "test-results/spfx-fast-validation");
mkdirSync(artifacts, { recursive: true });
const releaseProfile = acquireProfileLock(profileDir);
let context;
try {
  context = await chromium.launchPersistentContext(profileDir, { headless: true, viewport: { width: 1440, height: 960 } });
  const page = context.pages()[0] ?? await context.newPage();
  await page.goto(site.href.replace(/\/$/, ""), { waitUntil: "domcontentloaded", timeout: 60000 });
  for (let i = 0; i < 15 && page.url().includes("login.microsoftonline.com"); i++) await page.waitForTimeout(1000);
  if (page.url().includes("login.microsoftonline.com")) throw new Error("The profile is signed out. Sign in through the normal SharePoint browser flow, then retry.");

  const provision = await page.evaluate(async ({ apiRoot, componentId, title, fileName, properties }) => {
    const headers = { Accept: "application/json;odata=nometadata" };
    const call = async (path, method = "GET", body) => {
      const requestHeaders = { ...headers };
      if (method !== "GET") {
        const digestResponse = await fetch(`${apiRoot}/contextinfo`, { method: "POST", headers });
        if (!digestResponse.ok) throw new Error(`SharePoint request digest failed (HTTP ${digestResponse.status}).`);
        requestHeaders["X-RequestDigest"] = (await digestResponse.json()).FormDigestValue;
        requestHeaders["Content-Type"] = "application/json;odata=nometadata";
      }
      const response = await fetch(path, { method, headers: requestHeaders, body: body === undefined ? undefined : JSON.stringify(body) });
      if (!response.ok) throw new Error(`SharePoint page API failed (HTTP ${response.status}).`);
      return response.status === 204 ? undefined : response.json();
    };
    const components = await call(`${apiRoot}/web/GetClientSideWebParts`);
    const registered = (components.value ?? []).some((entry) => {
      const id = String(entry.Id ?? "").replace(/[{}]/g, "").toLowerCase();
      return id === componentId.toLowerCase();
    });
    if (!registered) throw new Error("The dashboard web-part component is not registered on this site. Deploy and install its SPFx package first.");

    const filter = encodeURIComponent(`FileLeafRef eq '${fileName}'`);
    const existing = await call(`${apiRoot}/web/lists/getbytitle('Site Pages')/items?$select=Id,Title,FileLeafRef&$filter=${filter}&$top=1`);
    if (existing.value?.length) {
      const current = existing.value[0];
      const page = await call(`${apiRoot}/sitepages/pages(${current.Id})?$select=Id,Title,FileName,CanvasContent1`);
      const canvas = JSON.parse(page.CanvasContent1 || "[]");
      const found = canvas.find((control) => String(control.webPartId ?? "").replace(/[{}]/g, "").toLowerCase() === componentId.toLowerCase());
      if (!found || JSON.stringify(found.webPartData?.properties ?? {}) !== JSON.stringify(properties)) {
        throw new Error("A page with the reserved validation filename exists but has different content. Review it before changing it.");
      }
      return { created: false, pageId: current.Id, fileName, title: page.Title };
    }

    const created = await call(`${apiRoot}/sitepages/pages`, "POST", { Title: title, PageLayoutType: "Article", Name: fileName });
    if (!created?.Id) throw new Error("SharePoint created no page ID.");
    const instanceId = crypto.randomUUID();
    const canvas = [
      {
        controlType: 3,
        position: { controlIndex: 1, sectionIndex: 1, sectionFactor: 12, zoneIndex: 1, layoutIndex: 1 },
        webPartId: componentId,
        id: instanceId,
        addedFromPersistedData: true,
        webPartData: {
          id: componentId,
          instanceId,
          title: "SPFx Fast Dashboard",
          description: "Authenticated SPFx Fast validation dashboard",
          dataVersion: "1.0",
          properties,
          serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} },
        },
        emphasis: {},
      },
      { controlType: 0, pageSettingsSlice: { isDefaultDescription: true, isDefaultThumbnail: true } },
    ];
    await call(`${apiRoot}/sitepages/pages(${created.Id})/checkoutpage`, "POST");
    await call(`${apiRoot}/sitepages/pages(${created.Id})/SavePageAsDraft`, "POST", { CanvasContent1: JSON.stringify(canvas), Title: title });
    await call(`${apiRoot}/sitepages/pages(${created.Id})/Publish`, "POST");
    const saved = await call(`${apiRoot}/sitepages/pages(${created.Id})?$select=Id,Title,FileName,CanvasContent1`);
    if (!String(saved.CanvasContent1 ?? "").includes(componentId)) throw new Error("Published page read-back did not include the dashboard component.");
    return { created: true, pageId: created.Id, fileName: saved.FileName ?? fileName, title: saved.Title };
  }, { apiRoot, componentId, title, fileName, properties });

  const pageUrl = new URL(`SitePages/${encodeURIComponent(provision.fileName)}`, `${site.href.replace(/\/$/, "")}/`).href;
  const bundles = [];
  page.on("response", (response) => {
    const path = new URL(response.url()).pathname;
    const match = path.match(/spfx-fast-dashboard-web-part_[a-f0-9]+\.js/i);
    if (match) bundles.push(match[0]);
  });
  await page.goto(pageUrl, { waitUntil: "domcontentloaded", timeout: 60000 });
  await page.getByRole("button", { name: "Run signed-in checks" }).waitFor({ state: "visible", timeout: 90000 });
  await page.getByRole("button", { name: "Test drive delta (one page)" }).waitFor({ state: "visible", timeout: 30000 });
  await page.screenshot({ path: `${artifacts}/dashboard.png`, fullPage: true });
  if (!bundles.includes(expectedBundle)) throw new Error(`The published page did not load the built bundle ${expectedBundle}; loaded: ${[...new Set(bundles)].join(", ") || "none"}.`);
  console.log(JSON.stringify({ status: "passed", ...provision, pageUrl, componentId, expectedBundle, loadedBundles: [...new Set(bundles)], screenshot: `${artifacts}/dashboard.png` }, null, 2));
} finally {
  await context?.close();
  releaseProfile();
}
