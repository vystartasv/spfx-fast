#!/usr/bin/env node
import { mkdirSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { chromium } from "playwright";
import { acquireProfileLock } from "../packages/playwright/src/profile-lock.js";

const usage = `Usage:
  node scripts/sharepoint-delta-workload-check.mjs --site <site-url> --page <validation-page-url> --profile-dir <signed-in-profile> [--artifacts <directory>]

Runs the 3,000-row IndexedDB list-delta and 3,000-folder drive-delta probes on a published dashboard page. Use --list-only or --drive-only to run one probe. It verifies the loaded bundle, confirms disposable writes, and saves full-page screenshots.`;
const args = process.argv.slice(2);
if (args.includes("--help") || args.includes("-h")) { console.log(usage); process.exit(0); }
const options = {};
for (let i = 0; i < args.length; i++) {
  if (args[i] === "--drive-only") { options["drive-only"] = true; continue; }
  if (args[i] === "--list-only") { options["list-only"] = true; continue; }
  if (!args[i].startsWith("--") || !args[i + 1] || args[i + 1].startsWith("--")) throw new Error(`Invalid option: ${args[i]}\n${usage}`);
  options[args[i].slice(2)] = args[++i];
}
for (const key of ["site", "page", "profile-dir"]) if (!options[key]) throw new Error(`Missing --${key}\n${usage}`);
if (options["drive-only"] && options["list-only"]) throw new Error("Choose only one of --drive-only or --list-only.");

const site = new URL(options.site);
const pageUrl = new URL(options.page);
if (site.protocol !== "https:" || pageUrl.protocol !== "https:" || site.origin !== pageUrl.origin || pageUrl.username || pageUrl.password || pageUrl.search || pageUrl.hash) {
  throw new Error("--site and --page must be HTTPS URLs on the same origin; --page cannot contain credentials, query, or hash.");
}
const profileDir = resolve(options["profile-dir"]);
const artifacts = resolve(options.artifacts ?? "test-results/sharepoint-delta-workloads");
const manifestPath = resolve("examples/fixtures/spfx-fast-dashboard/dist/7a1ec127-4cf7-4dfe-872a-0ab52f0cb56c.manifest.json");
const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
const expectedBundle = manifest.loaderConfig?.scriptResources?.["spfx-fast-dashboard-web-part"]?.path;
if (!expectedBundle) throw new Error("Build the dashboard fixture before running this check.");
mkdirSync(artifacts, { recursive: true });

const unlock = acquireProfileLock(profileDir);
let context;
try {
  context = await chromium.launchPersistentContext(profileDir, { headless: true, viewport: { width: 1440, height: 960 } });
  const page = context.pages()[0] ?? await context.newPage();
  const loadedBundles = new Set();
  const networkIssues = [];
  page.on("response", (response) => {
    const url = new URL(response.url());
    const match = url.pathname.match(/spfx-fast-dashboard-web-part_[a-f0-9]+\.js/i);
    if (match) loadedBundles.add(match[0]);
    if (response.status() >= 400 && (url.origin === site.origin || url.hostname === "graph.microsoft.com")) {
      networkIssues.push({ method: response.request().method(), path: url.pathname, status: response.status() });
    }
  });
  page.on("requestfailed", (request) => {
    const url = new URL(request.url());
    if (url.origin === site.origin || url.hostname === "graph.microsoft.com") {
      networkIssues.push({ method: request.method(), path: url.pathname, error: request.failure()?.errorText ?? "request failed" });
    }
  });
  await page.goto(pageUrl.href, { waitUntil: "domcontentloaded", timeout: 60_000 });

  const runCheck = async (buttonName, resultName, screenshotName) => {
    const button = page.getByRole("button", { name: buttonName });
    await button.waitFor({ state: "visible", timeout: 90_000 });
    if (!loadedBundles.has(expectedBundle)) throw new Error(`The page did not load the current bundle ${expectedBundle}.`);
    page.once("dialog", (dialog) => dialog.accept());
    await button.click();
    const result = page.locator("p[aria-live=polite]").filter({ hasText: resultName });
    await result.waitFor({ state: "visible", timeout: 600_000 });
    const resultText = (await result.innerText()).replace(/\s+/g, " ").trim();
    await page.screenshot({ path: `${artifacts}/${screenshotName}`, fullPage: true });
    if (!resultText.startsWith("PASS")) throw new Error(`${resultName} failed: ${resultText}`);
    return resultText;
  };

  try {
    const list = options["drive-only"] ? undefined : await runCheck("Test list delta paging", "3,000-item work-queue benchmark", "list-delta-3000.png");
    const drive = options["list-only"] ? undefined : await runCheck("Test drive delta paging", "Graph drive delta paging", "drive-delta-3000.png");
    console.log(JSON.stringify({ status: "passed", bundle: expectedBundle, ...(list ? { list } : {}), ...(drive ? { drive } : {}), screenshots: [
      ...(!options["drive-only"] ? [`${artifacts}/list-delta-3000.png`] : []), ...(!options["list-only"] ? [`${artifacts}/drive-delta-3000.png`] : []),
    ] }, null, 2));
  } catch (error) {
    console.log(JSON.stringify({ status: "failed", bundle: expectedBundle, error: error instanceof Error ? error.message : "validation failed", networkIssues: networkIssues.slice(-30) }, null, 2));
    process.exitCode = 1;
  }
} finally {
  await context?.close();
  unlock();
}
