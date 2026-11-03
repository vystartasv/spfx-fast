#!/usr/bin/env node
// SharePoint SPFx ALM through the current signed-in Playwright profile.
// Verifies deployments by reading back catalog and site state.
import { readFileSync, existsSync } from "node:fs";
import { resolve } from "node:path";
import { chromium } from "playwright";
import { acquireProfileLock } from "../packages/playwright/src/profile-lock.js";

const usage = `Usage:
  node scripts/sharepoint-app-lifecycle.mjs status --web <site-or-catalog-url> --config <package-solution.json> --profile-dir <path> [--scope tenant|sitecollection]
  node scripts/sharepoint-app-lifecycle.mjs deploy --catalog <catalog-url> --package <file.sppkg> --config <package-solution.json> --profile-dir <path> [--scope tenant|sitecollection]
  node scripts/sharepoint-app-lifecycle.mjs install --site <site-url> --config <package-solution.json> --profile-dir <path> [--scope tenant|sitecollection]
  node scripts/sharepoint-app-lifecycle.mjs reinstall --site <site-url> --config <package-solution.json> --profile-dir <path> [--scope tenant|sitecollection]
  node scripts/sharepoint-app-lifecycle.mjs verify --page <published-page-url> --bundle <served-bundle-filename> --profile-dir <path>

The profile must already be signed in. No tokens or cookies are read or copied.`;

const argv = process.argv.slice(2);
if (argv.includes("--help") || argv.includes("-h")) { console.log(usage); process.exit(0); }
const [command, ...args] = argv;
const options = {};
for (let i = 0; i < args.length; i++) {
  if (args[i] === "--help" || args[i] === "-h") { console.log(usage); process.exit(0); }
  if (!args[i].startsWith("--") || !args[i + 1] || args[i + 1].startsWith("--")) throw new Error(`Invalid option: ${args[i]}\n${usage}`);
  options[args[i].slice(2)] = args[++i];
}
if (!command || !["status", "deploy", "install", "reinstall", "verify"].includes(command)) throw new Error(usage);
const required = command === "verify" ? ["page", "bundle", "profile-dir"] : [command === "deploy" ? "catalog" : ["install", "reinstall"].includes(command) ? "site" : "web", "config", "profile-dir"];
for (const key of required) if (!options[key]) throw new Error(`Missing --${key}\n${usage}`);
let solution = {};
if (command !== "verify") {
  const configPath = resolve(options.config);
  if (!existsSync(configPath)) throw new Error(`Package config does not exist: ${configPath}`);
  solution = JSON.parse(readFileSync(configPath, "utf8")).solution;
  if (!solution?.name || !solution?.id || !solution?.version || typeof solution.skipFeatureDeployment !== "boolean") {
    throw new Error("package-solution.json must define solution.name, solution.id, solution.version, and solution.skipFeatureDeployment");
  }
}
const scope = options.scope ?? "tenant";
if (!(["tenant", "sitecollection"].includes(scope))) throw new Error("--scope must be tenant or sitecollection");
const baseUrl = new URL(options.page ?? options.catalog ?? options.site ?? options.web);
if (baseUrl.protocol !== "https:") throw new Error("SharePoint URL must use HTTPS");
const webPath = baseUrl.pathname.replace(/\/+$/, "");
const profileDir = resolve(options["profile-dir"]);
if (!existsSync(profileDir)) throw new Error(`Browser profile does not exist: ${profileDir}`);
const apiScope = `${scope}appcatalog`;
const accept = "application/json;odata=nometadata";

let packageBase64;
let packageName;
let expectedVersion;
let skipFeatureDeployment;
if (command === "deploy") {
  expectedVersion = solution.version;
  skipFeatureDeployment = solution.skipFeatureDeployment;
  const packagePath = resolve(options.package);
  if (!existsSync(packagePath)) throw new Error(`Package does not exist: ${packagePath}`);
  packageName = packagePath.split(/[\\/]/).at(-1);
  packageBase64 = readFileSync(packagePath).toString("base64");
}

const unlock = acquireProfileLock(profileDir);
let context;
try {
  context = await chromium.launchPersistentContext(profileDir, { headless: true });
  const page = context.pages()[0] ?? await context.newPage();
  await page.goto(baseUrl.href.replace(/\/$/, ""), { waitUntil: "domcontentloaded", timeout: 60000 });
  for (let i = 0; i < 15 && page.url().includes("login.microsoftonline.com"); i++) await page.waitForTimeout(1000);
  if (page.url().includes("login.microsoftonline.com")) throw new Error("This profile is signed out. Sign in with the normal browser flow, then retry.");

  if (command === "verify") {
    const expected = options.bundle;
    const served = await page.waitForFunction((bundle) => performance.getEntriesByType("resource").some((entry) => new URL(entry.name).pathname.endsWith(bundle)), expected, { timeout: 30000 })
      .then(() => true).catch(() => false);
    const resources = await page.evaluate((bundle) => {
      const assets = performance.getEntriesByType("resource").flatMap((entry) => {
        const pathname = new URL(entry.name).pathname;
        const asset = pathname.split("/").at(-1);
        return asset?.startsWith("library-demo-web-part_") && asset.endsWith(".js")
          ? [{ asset, transferSize: entry.transferSize, durationMs: Math.round(entry.duration * 10) / 10 }]
          : [];
      });
      return { expected: assets.filter((entry) => entry.asset === bundle), observed: assets };
    }, expected);
    if (!served) throw new Error(`Expected bundle was not served by the page: ${expected}. Observed: ${JSON.stringify(resources.observed)}`);
    console.log(JSON.stringify({ page: baseUrl.origin + baseUrl.pathname, expectedBundle: expected, served: true, resources: resources.expected }, null, 2));
    process.exitCode = 0;
  } else {

  const result = await page.evaluate(async ({ command, title, productId, apiScope, apiRoot, accept, packageName, packageBase64, expectedVersion, skipFeatureDeployment }) => {
    const check = async (response, label) => {
      if (!response.ok) {
        const body = await response.text().catch(() => "");
        throw new Error(`${label} failed (HTTP ${response.status})${body ? `: ${body.slice(0, 300)}` : ""}`);
      }
      return response;
    };
    const readApps = async () => {
      const response = await check(await fetch(`${apiRoot}/web/${apiScope}/AvailableApps`, { headers: { Accept: accept } }), "AvailableApps");
      return (await response.json()).value ?? [];
    };
    const findApp = async () => {
      const apps = await readApps();
      const app = apps.find((candidate) => candidate.Title === title && candidate.ProductId?.toLowerCase() === productId.toLowerCase());
      if (!app && apps.some((candidate) => candidate.Title === title)) throw new Error(`A same-name catalog app has a different ProductId; refusing to use it.`);
      if (!app) throw new Error(`No app named "${title}" with this package ProductId in the ${apiScope} catalog.`);
      return app;
    };
    const isInstalled = (app) => Boolean(app.InstalledVersion && app.InstalledVersion !== "0.0.0.0" && app.InstalledVersion === app.AppCatalogVersion);
    const appInstanceStatusNames = { 0: "InvalidStatus", 1: "Installing", 4: "Uninstalling", 5: "Installed", 7: "Canceling", 8: "Upgrading", 9: "Initialized", 10: "UpgradeCanceling", 11: "Disabling", 12: "Disabled", 13: "SecretRolling", 14: "Recycling", 15: "Recycled", 16: "Restoring", 17: "RestoreCanceling" };
    const describeInstance = ({ Id, ProductId, Status, Title }) => ({ Id, ProductId, Status, statusName: appInstanceStatusNames[Status] ?? "Unknown", Title });
    const safe = (app) => ({ title: app.Title, id: app.ID, productId: app.ProductId, catalogVersion: app.AppCatalogVersion, installedVersion: app.InstalledVersion, installed: isInstalled(app), deployed: app.CurrentVersionDeployed, enabled: app.IsEnabled, valid: app.IsValidAppPackage });
    const readInstances = async () => {
      const response = await check(await fetch(`${apiRoot}/web/GetAppInstancesByProductId('${productId}')`, { headers: { Accept: accept } }), "GetAppInstancesByProductId");
      return (await response.json()).value ?? [];
    };
    const waitForInstanceStatus = async (status, attempts = 48) => {
      let instances = [];
      for (let attempt = 0; attempt < attempts; attempt++) {
        await new Promise((resolve) => setTimeout(resolve, 2500));
        instances = await readInstances();
        if (status === "absent" ? instances.length === 0 : instances.some((instance) => instance.Status === status)) return instances;
      }
      throw new Error(`Timed out waiting for app instance status ${status}. ${JSON.stringify(instances.map(describeInstance))}`);
    };
    const digest = async () => {
      const response = await check(await fetch(`${apiRoot}/contextinfo`, { method: "POST", headers: { Accept: accept } }), "contextinfo");
      return (await response.json()).FormDigestValue;
    };
    const action = async (app, name, body) => {
      const headers = { Accept: accept, "X-RequestDigest": await digest() };
      if (body !== undefined) headers["Content-Type"] = "application/json;odata=nometadata;charset=utf-8";
      return fetch(`${apiRoot}/web/${apiScope}/AvailableApps/GetById('${app.ID}')/${name}`, { method: "POST", headers, body: body === undefined ? undefined : JSON.stringify(body) });
    };

    if (command === "reinstall") {
      let app = await findApp();
      if (!app.IsValidAppPackage || !app.CurrentVersionDeployed || !app.IsEnabled) throw new Error(`App must be valid, deployed, and enabled before reinstall. ${JSON.stringify(safe(app))}`);
      const actions = [];
      const existingInstances = await readInstances();
      if (existingInstances.length) {
        const response = await action(app, "Uninstall");
        const body = (await response.text().catch(() => "")).slice(0, 300);
        actions.push({ action: "uninstall", status: response.status, body });
        if (!response.ok) throw new Error(`Uninstall failed: ${JSON.stringify({ actions, app: safe(app) })}`);
        await waitForInstanceStatus("absent");
        app = await findApp();
      }
      const response = await action(app, "Install");
      const body = (await response.text().catch(() => "")).slice(0, 300);
      actions.push({ action: "install", status: response.status, body });
      if (!response.ok) throw new Error(`Install failed: ${JSON.stringify({ actions, app: safe(app) })}`);
      const instances = await waitForInstanceStatus(5);
      app = await findApp();
      return { ...safe(app), installedVersionMatchesCatalog: isInstalled(app), installed: true, appInstance: { Id: instances.find((instance) => instance.Status === 5)?.Id, Status: 5 }, actions, action: "reinstalled" };
    }

    if (command === "status") {
      const app = await findApp();
      const instances = await readInstances();
      const instance = instances.find((item) => item.Status === 5);
      return { ...safe(app), installedVersionMatchesCatalog: isInstalled(app), installed: Boolean(instance), appInstance: instance ? describeInstance(instance) : null, appInstances: instances.map(describeInstance) };
    }
    if (command === "deploy") {
      const catalogApps = await readApps();
      if (catalogApps.some((candidate) => candidate.Title === title && candidate.ProductId?.toLowerCase() !== productId.toLowerCase())) {
        throw new Error("A same-name catalog app has a different ProductId; refusing to overwrite it.");
      }
      let app = catalogApps.find((candidate) => candidate.Title === title && candidate.ProductId?.toLowerCase() === productId.toLowerCase());
      if (app?.CurrentVersionDeployed && app.AppCatalogVersion === expectedVersion) {
        if (!app.IsValidAppPackage) throw new Error("SharePoint reports the matching package as invalid.");
        return { ...safe(app), action: "already-deployed" };
      }
      if (!app || app.AppCatalogVersion !== expectedVersion) {
        const bytes = Uint8Array.from(atob(packageBase64), (char) => char.charCodeAt(0));
        const response = await check(await fetch(`${apiRoot}/web/${apiScope}/Add(overwrite=true,url='${encodeURIComponent(packageName)}')`, {
          method: "POST", headers: { Accept: accept, "X-RequestDigest": await digest() }, body: bytes,
        }), "app-catalog upload");
        await response.text().catch(() => "");
        app = await findApp();
      }
      if (!app.IsValidAppPackage) throw new Error("SharePoint marked the uploaded package invalid.");
      if (app.AppCatalogVersion !== expectedVersion) throw new Error(`Catalog version ${app.AppCatalogVersion} does not match package version ${expectedVersion}.`);
      const response = await action(app, "Deploy", { skipFeatureDeployment });
      if (!response.ok && response.status !== 403) await check(response, "Deploy");
      for (let attempt = 0; attempt < 12; attempt++) {
        await new Promise((resolve) => setTimeout(resolve, 2500));
        app = await findApp();
        if (app.CurrentVersionDeployed) return { ...safe(app), action: "deployed" };
      }
      throw new Error(`Deploy was requested, but SharePoint did not report CurrentVersionDeployed=true. ${JSON.stringify(safe(app))}`);
    }

    let app = await findApp();
    if (!app.IsValidAppPackage || !app.CurrentVersionDeployed || !app.IsEnabled) {
      throw new Error(`App must be valid, deployed, and enabled before site install. ${JSON.stringify(safe(app))}`);
    }
    let instances = await readInstances();
    if (instances.some((instance) => instance.Status === 5)) return { ...safe(app), installedVersionMatchesCatalog: isInstalled(app), installed: true, appInstance: { Id: instances.find((instance) => instance.Status === 5).Id, Status: 5 }, action: "already-installed" };
    // Initialized (9) means the instance exists but Install/Upgrade has not completed;
    // issue the lifecycle action below instead of waiting for a transition that may never start.
    if (instances.some((instance) => [1, 8].includes(instance.Status))) {
      instances = await waitForInstanceStatus(5);
      app = await findApp();
      return { ...safe(app), installedVersionMatchesCatalog: isInstalled(app), installed: true, appInstance: { Id: instances.find((instance) => instance.Status === 5)?.Id, Status: 5 }, action: "install-completed" };
    }
    // The site-scoped instance is authoritative for lifecycle choice; InstalledVersion can be blank/stale.
    const verb = instances.length ? "Upgrade" : "Install";
    let response;
    for (let attempt = 0; attempt < 12; attempt++) {
      response = await action(app, verb);
      if (response.status !== 404) break;
      await new Promise((resolve) => setTimeout(resolve, 2500));
      app = await findApp();
    }
    const actionError = response.ok ? "" : (await response.text().catch(() => "")).slice(0, 300);
    if (!response.ok) throw new Error(`${verb} failed (HTTP ${response.status})${actionError ? `: ${actionError}` : ""}`);
    // InstalledVersion can lag independently; AppInstance.Status is the site-scoped lifecycle state.
    try {
      instances = await waitForInstanceStatus(5);
    } catch (error) {
      throw new Error(`${verb} returned HTTP ${response.status}, but SharePoint did not install the app instance: ${error.message}`);
    }
    app = await findApp();
    return { ...safe(app), installedVersionMatchesCatalog: isInstalled(app), installed: true, appInstance: { Id: instances.find((instance) => instance.Status === 5)?.Id, Status: 5 }, action: verb.toLowerCase(), actionResponse: { status: response.status, body: actionError } };
  }, { command, title: solution.name, productId: solution.id, apiScope, apiRoot: `${webPath}/_api`, accept, packageName, packageBase64, expectedVersion, skipFeatureDeployment });
  console.log(JSON.stringify(result, null, 2));
  }
} finally {
  await context?.close();
  unlock();
}
