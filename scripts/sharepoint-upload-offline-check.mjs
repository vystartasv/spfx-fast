#!/usr/bin/env node
import { mkdirSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { chromium } from "playwright";
import { acquireProfileLock } from "../packages/playwright/src/profile-lock.js";

const usage = `Usage:
  node scripts/sharepoint-upload-offline-check.mjs --site <site-url> --page <validation-page-url> --profile-dir <signed-in-profile> [--artifacts <directory>]

Runs the dashboard's 128 MiB upload probe, takes the browser offline after the first accepted chunk, reconnects, then verifies server-range resume, file bytes, and cleanup. Add --post-window to stay offline through bounded retries, then resume from the page UI.`;
const args = process.argv.slice(2);
if (args.includes("--help") || args.includes("-h")) { console.log(usage); process.exit(0); }
const options = {};
for (let i = 0; i < args.length; i++) {
  if (args[i] === "--post-window") { options["post-window"] = true; continue; }
  if (!args[i].startsWith("--") || !args[i + 1] || args[i + 1].startsWith("--")) throw new Error(`Invalid option: ${args[i]}\n${usage}`);
  options[args[i].slice(2)] = args[++i];
}
for (const key of ["site", "page", "profile-dir"]) if (!options[key]) throw new Error(`Missing --${key}\n${usage}`);

const site = new URL(options.site);
const pageUrl = new URL(options.page);
if (site.protocol !== "https:" || pageUrl.protocol !== "https:" || site.origin !== pageUrl.origin || pageUrl.username || pageUrl.password || pageUrl.search || pageUrl.hash) {
  throw new Error("--site and --page must be HTTPS URLs on the same origin; --page cannot contain credentials, query, or hash.");
}
const profileDir = resolve(options["profile-dir"]);
const artifacts = resolve(options.artifacts ?? "test-results/sharepoint-upload-offline");
const postWindow = options["post-window"] === true;
const buttonName = "Test 128 MiB upload and download";
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
  const network = { firstAcceptedBytes: 0, offlineEntered: false, failedPut: false, failedStatusGet: false, statusGetFailures: 0, resumedRange: 0 };
  const loadedBundles = new Set();
  let uploadUrl;
  let outageOccurred = false;

  page.on("response", async (response) => {
    const request = response.request();
    if (request.method() === "PUT" && response.status() === 202 && !network.offlineEntered) {
      uploadUrl = response.url();
      const range = request.headers()["content-range"] ?? "";
      const match = range.match(/^bytes 0-(\d+)\//);
      network.firstAcceptedBytes = match ? Number(match[1]) + 1 : 0;
      await context.setOffline(true);
      network.offlineEntered = true;
    } else if (request.method() === "GET" && uploadUrl === response.url() && response.ok()) {
      try {
        const state = await response.json();
        const next = Number(state?.nextExpectedRanges?.[0]?.split("-")[0]);
        if (Number.isInteger(next)) network.resumedRange = next;
      } catch { /* The client still verifies upload completion and file bytes. */ }
    }
    const match = new URL(response.url()).pathname.match(/spfx-fast-dashboard-web-part_[a-f0-9]+\.js/i);
    if (match) loadedBundles.add(match[0]);
  });

  page.on("requestfailed", async (request) => {
    if (network.offlineEntered && request.method() === "PUT" && request.url() === uploadUrl && !network.failedPut) {
      network.failedPut = true;
      outageOccurred = true;
    } else if (outageOccurred && request.method() === "GET" && request.url() === uploadUrl) {
      network.failedStatusGet = true;
      network.statusGetFailures++;
      if (!postWindow && network.statusGetFailures === 1) await context.setOffline(false);
    }
  });

  await page.goto(pageUrl.href, { waitUntil: "domcontentloaded", timeout: 60_000 });
  const button = page.getByRole("button", { name: buttonName });
  await button.waitFor({ state: "visible", timeout: 90_000 });
  if (!loadedBundles.has(expectedBundle)) throw new Error(`The page did not load the current bundle ${expectedBundle}.`);
  page.once("dialog", (dialog) => dialog.accept());
  await button.click();
  const result = page.locator("p[aria-live=polite]").filter({ hasText: "File upload/download/cleanup" });
  if (postWindow) {
    const resume = page.getByRole("button", { name: "Resume upload and verify" });
    await resume.waitFor({ state: "visible", timeout: 180_000 });
    if (!network.failedPut || network.statusGetFailures < 6) throw new Error(`The upload did not exhaust its retry window (failed PUT: ${network.failedPut}; failed status reads: ${network.statusGetFailures}).`);
    await context.setOffline(false);
    await resume.click();
  }
  await page.waitForFunction(() => {
    const result = Array.from(document.querySelectorAll("p[aria-live=polite]")).find((element) => element.textContent?.includes("File upload/download/cleanup"));
    return /^(PASS|FAIL)\b/.test(result?.textContent?.trim() ?? "");
  }, { timeout: 600_000 });
  if (postWindow) {
    await page.waitForFunction(() => {
      const button = Array.from(document.querySelectorAll("button")).find((element) => element.textContent?.trim() === "Test 128 MiB upload and download");
      return Boolean(button && !button.disabled);
    }, { timeout: 10_000 });
  }
  await context.setOffline(false);
  const resultText = (await result.innerText()).replace(/\s+/g, " ").trim();
  const screenshot = `${artifacts}/upload-offline-${postWindow ? "post-window-" : ""}recovery.png`;
  await page.screenshot({ path: screenshot, fullPage: true });

  const valid = resultText.startsWith("PASS") && network.offlineEntered && network.firstAcceptedBytes > 0 && network.failedPut && network.failedStatusGet && network.resumedRange >= network.firstAcceptedBytes && (!postWindow || resultText.includes("resumed from its live session"));
  console.log(JSON.stringify({
    status: valid ? "passed" : "failed",
    bundle: expectedBundle,
    browserWentOffline: network.offlineEntered,
    firstAcceptedBytes: network.firstAcceptedBytes,
    browserOfflinePutFailed: network.failedPut,
    browserOfflineStatusCheckFailed: network.failedStatusGet,
    statusReadFailuresWhileOffline: network.statusGetFailures,
    serverReportedResumeRange: network.resumedRange,
    result: resultText,
    screenshot,
  }, null, 2));
  if (!valid) process.exitCode = 1;
} finally {
  await context?.setOffline(false).catch(() => undefined);
  await context?.close();
  unlock();
}
