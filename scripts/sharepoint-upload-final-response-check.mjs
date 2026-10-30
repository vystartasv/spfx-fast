#!/usr/bin/env node
import { mkdirSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { chromium } from "playwright";
import { acquireProfileLock } from "../packages/playwright/src/profile-lock.js";

const usage = `Usage:
  node scripts/sharepoint-upload-final-response-check.mjs --site <site-url> --page <validation-page-url> --profile-dir <signed-in-profile> [--replace | --rename] [--artifacts <directory>]

Runs a disposable upload check. Default drops the final successful response; --replace tests lost-response recovery for an existing file; --rename observes Graph's duplicate-name behavior.`;
const args = process.argv.slice(2);
if (args.includes("--help") || args.includes("-h")) { console.log(usage); process.exit(0); }
const options = {};
for (let i = 0; i < args.length; i++) {
  if (args[i] === "--replace") { options.replace = true; continue; }
  if (args[i] === "--rename") { options.rename = true; continue; }
  if (!args[i].startsWith("--") || !args[i + 1] || args[i + 1].startsWith("--")) throw new Error(`Invalid option: ${args[i]}\n${usage}`);
  options[args[i].slice(2)] = args[++i];
}
for (const key of ["site", "page", "profile-dir"]) if (!options[key]) throw new Error(`Missing --${key}\n${usage}`);
if (options.replace && options.rename) throw new Error("Choose only one of --replace or --rename.");

const site = new URL(options.site);
const pageUrl = new URL(options.page);
if (site.protocol !== "https:" || pageUrl.protocol !== "https:" || site.origin !== pageUrl.origin || pageUrl.username || pageUrl.password || pageUrl.search || pageUrl.hash) {
  throw new Error("--site and --page must be HTTPS URLs on the same origin; --page cannot contain credentials, query, or hash.");
}
const profileDir = resolve(options["profile-dir"]);
const artifacts = resolve(options.artifacts ?? "test-results/sharepoint-upload-final-response");
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
  const networkFailures = [];
  page.on("response", (response) => {
    const url = new URL(response.url());
    const match = url.pathname.match(/spfx-fast-dashboard-web-part_[a-f0-9]+\.js/i);
    if (match) loadedBundles.add(match[0]);
    if (response.status() >= 400 && (url.origin === site.origin || url.hostname === "graph.microsoft.com")) {
      const failure = { method: response.request().method(), host: url.hostname, path: url.pathname, status: response.status() };
      if (url.pathname.endsWith("/createUploadSession")) {
        try {
          const body = response.request().postDataJSON();
          failure.requestBodyKeys = Object.keys(body ?? {}).sort();
          failure.itemBodyKeys = Object.keys(body?.item ?? {}).sort();
          failure.conflictBehavior = body?.item?.["@microsoft.graph.conflictBehavior"];
        } catch { /* Request body may not be JSON. */ }
        void response.json().then((body) => {
          const error = body?.error;
          if (typeof error?.code === "string") failure.graphCode = error.code.slice(0, 100);
          if (typeof error?.message === "string") failure.graphMessage = error.message.slice(0, 240);
        }).catch(() => undefined);
      }
      networkFailures.push(failure);
    }
  });
  page.on("requestfailed", (request) => {
    const url = new URL(request.url());
    if (url.origin === site.origin || url.hostname === "graph.microsoft.com") {
      networkFailures.push({ method: request.method(), host: url.hostname, path: url.pathname, error: request.failure()?.errorText ?? "request failed" });
    }
  });
  if (!options.replace && !options.rename) await page.addInitScript(() => {
    const originalFetch = window.fetch.bind(window);
    const state = { firstChunkAccepted: false, finalResponseDropped: false };
    Object.defineProperty(window, "__spfxFastFinalResponseProbe", { value: state });
    window.fetch = async (input, init) => {
      const request = input instanceof Request ? input : undefined;
      const method = init?.method ?? request?.method ?? "GET";
      const headers = new Headers(init?.headers ?? request?.headers);
      const range = headers.get("Content-Range") ?? "";
      const response = await originalFetch(input, init);
      if (method === "PUT" && range.startsWith("bytes 0-") && response.status === 202) state.firstChunkAccepted = true;
      if (method === "PUT" && state.firstChunkAccepted && (response.status === 200 || response.status === 201)) {
        state.finalResponseDropped = true;
        await response.body?.cancel().catch(() => undefined);
        throw new TypeError("Simulated lost final upload response after SharePoint committed the file.");
      }
      return response;
    };
  });
  await page.goto(pageUrl.href, { waitUntil: "domcontentloaded", timeout: 60_000 });
  const resultName = options.rename ? "Graph rename conflict" : "File upload/download/cleanup";
  const button = page.getByRole("button", { name: options.replace ? "Test replace final response recovery" : options.rename ? "Test Graph rename conflict" : "Test upload recovery" });
  await button.waitFor({ state: "visible", timeout: 90_000 });
  if (!loadedBundles.has(expectedBundle)) throw new Error(`The page did not load the current bundle ${expectedBundle}.`);
  page.once("dialog", (dialog) => dialog.accept());
  await button.click();
  await page.waitForFunction((name) => {
    const result = Array.from(document.querySelectorAll("p[aria-live=polite]"))
      .find((element) => element.textContent?.includes(name));
    return /^(PASS|FAIL)\b/.test(result?.textContent?.trim() ?? "");
  }, resultName, { timeout: 300_000 });
  const result = page.locator("p[aria-live=polite]").filter({ hasText: resultName }).last();
  const resultText = (await result.innerText()).replace(/\s+/g, " ").trim();
  const probe = options.rename
    ? { conflictBehavior: "rename" }
    : options.replace
    ? { finalResponseDropped: resultText.includes("after its final response was lost") }
    : await page.evaluate(() => window.__spfxFastFinalResponseProbe);
  const screenshot = `${artifacts}/upload-${options.rename ? "rename-conflict" : "final-response-recovery"}.png`;
  await page.screenshot({ path: screenshot, fullPage: true });
  const passed = resultText.startsWith("PASS") && (options.rename || (probe?.finalResponseDropped && (options.replace || probe?.firstChunkAccepted)));
  console.log(JSON.stringify({ status: passed ? "passed" : "failed", mode: options.rename ? "rename" : options.replace ? "replace" : "unique-name", bundle: expectedBundle, ...probe, result: resultText, networkFailures: networkFailures.slice(-20), screenshot }, null, 2));
  if (!passed) process.exitCode = 1;
} finally {
  await context?.close();
  unlock();
}
