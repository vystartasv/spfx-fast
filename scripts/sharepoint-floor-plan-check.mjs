#!/usr/bin/env node
import { mkdirSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { chromium } from "playwright";
import { acquireProfileLock } from "../packages/playwright/src/profile-lock.js";

const usage = `Usage:
  node scripts/sharepoint-floor-plan-check.mjs --page <published-page-url> --profile-dir <signed-in-profile> [--artifacts <directory>]

Checks the published SPFx dashboard bundle, uploads a disposable synthetic PNG, reads it through DriveFloorPlanMap, captures the component, and removes the temporary file.`;
const args = process.argv.slice(2);
if (args.includes("--help") || args.includes("-h")) { console.log(usage); process.exit(0); }
const options = {};
for (let index = 0; index < args.length; index++) {
  if (!args[index].startsWith("--") || !args[index + 1] || args[index + 1].startsWith("--")) throw new Error(`Invalid option: ${args[index]}\n${usage}`);
  options[args[index].slice(2)] = args[++index];
}
if (!options.page || !options["profile-dir"]) throw new Error(`Missing --page or --profile-dir\n${usage}`);
const pageUrl = new URL(options.page);
if (pageUrl.protocol !== "https:" || pageUrl.username || pageUrl.password || pageUrl.search || pageUrl.hash) {
  throw new Error("--page must be an HTTPS SharePoint page without credentials, query, or hash.");
}
const profileDir = resolve(options["profile-dir"]);
const artifacts = resolve(options.artifacts ?? "test-results/sharepoint-floor-plan");
const manifest = JSON.parse(readFileSync(resolve("examples/fixtures/spfx-fast-dashboard/dist/7a1ec127-4cf7-4dfe-872a-0ab52f0cb56c.manifest.json"), "utf8"));
const expectedBundle = manifest.loaderConfig?.scriptResources?.["spfx-fast-dashboard-web-part"]?.path;
if (!expectedBundle) throw new Error("Build the dashboard fixture before running this check.");
mkdirSync(artifacts, { recursive: true });

const unlock = acquireProfileLock(profileDir);
let context;
let page;
let cleanupPassed = false;
try {
  context = await chromium.launchPersistentContext(profileDir, { headless: true, viewport: { width: 1440, height: 960 } });
  page = context.pages()[0] ?? await context.newPage();
  const loadedBundles = new Set();
  page.on("response", (response) => {
    const match = new URL(response.url()).pathname.match(/spfx-fast-dashboard-web-part_[a-f0-9]+\.js/i);
    if (match) loadedBundles.add(match[0]);
  });
  await page.goto(pageUrl.href, { waitUntil: "domcontentloaded", timeout: 60_000 });
  const dashboard = page.getByRole("tablist", { name: "Dashboard views" });
  await dashboard.waitFor({ state: "visible", timeout: 90_000 });
  if (!loadedBundles.has(expectedBundle)) throw new Error(`The published page did not load the current bundle ${expectedBundle}.`);
  await dashboard.getByRole("tab", { name: "Floor plan" }).click();
  const probe = page.getByRole("region", { name: "Private floor-plan image check" });
  const start = probe.getByRole("button", { name: "Test private floor-plan image" });
  await start.waitFor({ state: "visible", timeout: 30_000 });
  page.once("dialog", (dialog) => dialog.accept());
  await start.click();
  const map = probe.getByRole("region", { name: "Synthetic facilities floor plan" });
  const image = map.locator('img[alt=""]');
  await image.waitFor({ state: "visible", timeout: 90_000 });
  await page.waitForFunction(() => document.querySelector('img[alt=""]')?.naturalWidth > 0, undefined, { timeout: 90_000 });
  const dimensions = await image.evaluate((element) => ({ width: element.naturalWidth, height: element.naturalHeight }));
  await map.screenshot({ path: `${artifacts}/private-floor-plan.png` });
  await probe.getByRole("button", { name: "Remove temporary floor-plan image" }).click();
  const result = probe.getByRole("status").filter({ hasText: /removed/i });
  await result.waitFor({ state: "visible", timeout: 60_000 });
  cleanupPassed = true;
  console.log(JSON.stringify({ status: "passed", bundle: expectedBundle, image: dimensions, cleanup: "temporary file removed", screenshot: `${artifacts}/private-floor-plan.png` }, null, 2));
} catch (error) {
  if (page && !cleanupPassed) {
    try {
      const remove = page.getByRole("button", { name: "Remove temporary floor-plan image" });
      if (await remove.isVisible({ timeout: 2000 })) {
        await remove.click();
        await page.getByRole("status").filter({ hasText: /removed/i }).waitFor({ timeout: 30_000 });
        cleanupPassed = true;
      }
    } catch { /* Report manual cleanup details below. */ }
  }
  console.log(JSON.stringify({ status: "failed", error: error instanceof Error ? error.message : "floor-plan validation failed", cleanup: cleanupPassed ? "temporary file removed" : "inspect the page and remove any SPFxFastFloorPlan_*.png file" }, null, 2));
  process.exitCode = 1;
} finally {
  await context?.close();
  unlock();
}
