import assert from "node:assert/strict";
import { mkdirSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";
import { chromium } from "playwright";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const scratch = mkdtempSync(join(tmpdir(), "spfx-fast-browser-a11y-"));
const bundle = join(scratch, "harness.js");
const browserErrors = [];
let axeScanCount = 0;
const screenshotDir = process.env.SPFX_FAST_BROWSER_SCREENSHOT_DIR;
if (screenshotDir) mkdirSync(screenshotDir, { recursive: true });

await build({
  entryPoints: [join(root, "tests/browser/harness.tsx")],
  outfile: bundle,
  bundle: true,
  format: "iife",
  globalName: "SPFxFastHarness",
  platform: "browser",
  target: "chrome120",
  alias: {
    react: join(root, "node_modules/react"),
    "@spfx-fast/core/markdown": join(root, "src/markdown.tsx"),
  },
  jsx: "automatic",
  logLevel: "silent",
});

const browser = await chromium.launch({ headless: true });

async function createPage() {
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  page.on("pageerror", (error) => browserErrors.push(error.message));
  await page.setContent(`<!doctype html>
    <html lang="en"><head><meta charset="utf-8"><title>SPFx Fast sample</title></head>
    <body><div id="app"></div></body></html>`);
  await page.evaluate(() => {
    if (typeof crypto.randomUUID !== "function") Object.defineProperty(crypto, "randomUUID", { value: () => "00000000-0000-4000-8000-000000000001" });
  });
  await page.addScriptTag({ path: bundle });
  await page.addScriptTag({ path: join(root, "node_modules/axe-core/axe.min.js") });
  return page;
}

async function assertNoAxeViolations(page, surface) {
  axeScanCount++;
  // Fluent UI's Tabster focus sentinels are intentionally hidden focus targets;
  // axe reports aria-hidden-focus on these implementation nodes, so exclude only them.
  const results = await page.evaluate(async () => window.axe.run({
    exclude: ["[data-tabster-dummy]"],
  }, {
    runOnly: { type: "tag", values: ["wcag2a", "wcag2aa", "wcag21a", "wcag21aa", "wcag22aa"] },
  }));
  assert.deepEqual(results.violations.map(({ id, impact, help, nodes }) => ({
    id, impact, help, elements: nodes.map(({ target, html, failureSummary }) => ({ target, html, failureSummary })),
  })), [], `${surface} must pass axe WCAG A/AA rules`);
}

try {
  const formPage = await createPage();
  await formPage.evaluate(() => window.SPFxFastHarness.mountForm());
  const title = formPage.getByRole("textbox", { name: "Title" });
  await title.fill("Quarterly review");
  await formPage.getByRole("button", { name: "Save" }).click();

  const summary = formPage.getByRole("status");
  await summary.waitFor({ state: "visible" });
  assert.match(await summary.innerText(), /Priority/);
  const priority = formPage.getByRole("combobox", { name: "Priority" });
  assert.equal(await priority.getAttribute("aria-invalid"), "true");
  assert.equal(await priority.evaluate((element) => element === document.activeElement), true, "invalid choice receives keyboard focus");
  await assertNoAxeViolations(formPage, "invalid form");

  await priority.press("Enter");
  await formPage.getByRole("option", { name: "High" }).waitFor({ state: "visible" });
  await priority.press("Enter");
  assert.match(await priority.innerText(), /High/, "keyboard selection updates the required choice");
  await formPage.getByRole("button", { name: "Save" }).click();
  await formPage.waitForFunction(() => window.__spfxFastSubmitCount === 1);
  await assertNoAxeViolations(formPage, "valid form");
  if (screenshotDir) await formPage.screenshot({ path: join(screenshotDir, "item-form.png"), fullPage: true });
  await formPage.close();

  const referenceFormPage = await createPage();
  await referenceFormPage.evaluate(() => window.SPFxFastHarness.mountPersonLookupForm());
  for (const [field, option] of [["Owner", "Synthetic owner"], ["Project", "Synthetic project"]]) {
    const control = referenceFormPage.getByRole("combobox", { name: field });
    await control.click();
    await referenceFormPage.getByRole("option", { name: option }).click();
  }
  await referenceFormPage.getByRole("button", { name: "Save" }).click();
  await referenceFormPage.waitForFunction(() => window.__spfxFastReferenceSubmit !== undefined);
  assert.deepEqual(await referenceFormPage.evaluate(() => window.__spfxFastReferenceSubmit), {
    OwnerLookupId: "13",
    ProjectLookupId: "42",
  }, "person and lookup selections serialize to Graph LookupId fields");
  await assertNoAxeViolations(referenceFormPage, "Person and Lookup form");
  await referenceFormPage.close();

  const tenantPage = await createPage();
  tenantPage.on("dialog", (dialog) => dialog.accept());
  await tenantPage.evaluate(() => window.SPFxFastHarness.mountTenantValidationPanel());
  await tenantPage.getByRole("button", { name: "Run signed-in checks" }).click();
  await tenantPage.getByText("Cache repeat", { exact: false }).waitFor();
  await tenantPage.getByText("Repeated list read was served from cache.").waitFor();
  assert.deepEqual(await tenantPage.getByRole("list", { name: "Validation results" }).locator("strong").allTextContents(), Array(12).fill("PASS"), "all twelve bounded signed-in checks report success");
  await tenantPage.getByRole("button", { name: "Test drive delta (one page)" }).click();
  await tenantPage.getByText("Read one page (2 item(s)); checkpoint saved. Run again to continue.").waitFor();
  assert.equal(await tenantPage.evaluate(() => window.__spfxFastDeltaMaxPages), 1, "tenant delta probe must stay bounded to one page");
  await tenantPage.getByTestId("tenant-validation-run-list-crud").click();
  const crudResult = tenantPage.getByTestId("tenant-validation-list-crud-result");
  await crudResult.waitFor({ state: "visible" });
  assert.match(await crudResult.innerText(), /Create, read, update, read, and delete succeeded for one temporary list item\./);
  await tenantPage.getByRole("button", { name: "Test 10 MiB upload and download" }).click();
  await tenantPage.getByText("Graph resumable upload and fresh binary download succeeded for 10 MiB + 1 byte; content matched. Temporary file deleted.").waitFor();
  assert.equal(await tenantPage.evaluate(() => window.__spfxFastTenantDeletes), 2, "CRUD and file checks clean up their temporary records");
  await assertNoAxeViolations(tenantPage, "tenant validation panel");
  await tenantPage.close();

  const samplesPage = await createPage();
  const sampleWriteProbe = await samplesPage.evaluate(() => window.SPFxFastHarness.probeSampleWrite());
  assert.ok(sampleWriteProbe.id, `sample Graph write transport failed: ${sampleWriteProbe.error ?? "no item id returned"}`);
  await samplesPage.evaluate(() => window.SPFxFastHarness.mountSampleWebparts());
  await samplesPage.getByRole("heading", { name: "Request dashboard" }).waitFor();
  try { await samplesPage.getByRole("region", { name: "Request dashboard" }).getByText("Laptop for new starter", { exact: true }).waitFor({ timeout: 5000 }); }
  catch (error) {
    console.error(`Sample gallery diagnostic: ${await samplesPage.locator("body").innerText()}\nPage errors: ${browserErrors.join(" | ") || "none"}`);
    if (screenshotDir) await samplesPage.screenshot({ path: join(screenshotDir, "sample-webparts-failure.png"), fullPage: true });
    throw error;
  }
  await samplesPage.getByRole("heading", { name: "Team task queue" }).waitFor();
  await samplesPage.getByText("Review laptop supplier quote", { exact: true }).waitFor();
  assert.equal(await samplesPage.getByText("Update West Wing asset register", { exact: true }).count(), 0, "server-filtered queue excludes completed tasks");
  await samplesPage.getByText(/North Campus operations/).waitFor();
  await samplesPage.getByRole("heading", { name: "Manager approvals" }).waitFor();
  assert.equal(await samplesPage.getByRole("region", { name: "Approval inbox" }).getByRole("button", { name: "Approve" }).count(), 2);
  if (screenshotDir) {
    await samplesPage.screenshot({ path: join(screenshotDir, "sample-webparts-gallery.png"), fullPage: true });
    await samplesPage.getByRole("region", { name: "Request dashboard" }).screenshot({ path: join(screenshotDir, "request-dashboard.png") });
    await samplesPage.getByRole("region", { name: "Team task queue" }).screenshot({ path: join(screenshotDir, "team-task-queue.png") });
    await samplesPage.getByRole("region", { name: "Team knowledge" }).screenshot({ path: join(screenshotDir, "knowledge-panel.png") });
    await samplesPage.getByRole("region", { name: "Approval inbox" }).screenshot({ path: join(screenshotDir, "approval-inbox.png") });
  }
  await samplesPage.getByRole("textbox", { name: "Request", exact: true }).fill("New ergonomic chair request");
  await samplesPage.getByRole("textbox", { name: "Requester", exact: true }).fill("Demo User");
  await samplesPage.getByRole("button", { name: "Submit request" }).click();
  try { await samplesPage.getByText("New ergonomic chair request", { exact: true }).waitFor({ timeout: 5000 }); }
  catch (error) {
    console.error(`Sample create diagnostic: ${await samplesPage.locator("body").innerText()}\nPage errors: ${browserErrors.join(" | ") || "none"}`);
    throw error;
  }
  await samplesPage.getByRole("region", { name: "Approval inbox" }).getByRole("button", { name: "Approve" }).first().click();
  const approvalsRegion = samplesPage.getByRole("region", { name: "Approval inbox" });
  await approvalsRegion.getByText("Laptop for new starter", { exact: true }).waitFor({ state: "hidden" });
  assert.equal(await approvalsRegion.getByRole("button", { name: "Approve" }).count(), 1, "approval action updates the filtered sample list");
  await assertNoAxeViolations(samplesPage, "sample webparts gallery");
  await samplesPage.close();

  const requestAttachmentsPage = await createPage();
  await requestAttachmentsPage.evaluate(() => window.SPFxFastHarness.mountRequestDashboardAttachments());
  await requestAttachmentsPage.getByRole("heading", { name: "Request dashboard" }).waitFor();
  await requestAttachmentsPage.getByRole("textbox", { name: "Request", exact: true }).fill("Request with a quote");
  await requestAttachmentsPage.getByRole("textbox", { name: "Requester", exact: true }).fill("Demo User");
  await requestAttachmentsPage.getByRole("button", { name: "Submit request" }).click();
  await requestAttachmentsPage.getByRole("heading", { name: "Item attachments" }).waitFor();
  await requestAttachmentsPage.getByText("List item 42 · 1 attachment").waitFor();
  await requestAttachmentsPage.getByRole("button", { name: "Download request-quote.txt" }).waitFor();
  await assertNoAxeViolations(requestAttachmentsPage, "request form with authenticated attachments");
  if (screenshotDir) await requestAttachmentsPage.screenshot({ path: join(screenshotDir, "request-attachments.png"), fullPage: true });
  await requestAttachmentsPage.close();

  const markdownPage = await createPage();
  await markdownPage.evaluate(() => window.SPFxFastHarness.mountMarkdownSecurity());
  const markdownRegion = markdownPage.getByRole("region", { name: "Markdown safety test" });
  await markdownRegion.getByRole("heading", { name: "Safe heading" }).waitFor();
  await markdownRegion.getByRole("table").waitFor();
  assert.equal(await markdownRegion.locator("script").count(), 0, "Markdown HTML is rendered as inert text");
  assert.equal(await markdownRegion.locator("img").count(), 0, "Markdown cannot create executable image handlers");
  assert.equal(await markdownRegion.locator('a[href^="javascript:"]').count(), 0, "unsafe Markdown URLs are removed");
  assert.equal(await markdownPage.evaluate(() => window.__spfxFastXss), undefined, "Markdown content cannot execute scripts");
  assert.equal(await markdownRegion.locator("[onerror]").count(), 0, "raw HTML event handlers are not rendered");
  await assertNoAxeViolations(markdownPage, "sanitized Markdown renderer");
  await markdownPage.close();

  const refreshPage = await createPage();
  await refreshPage.evaluate(() => window.SPFxFastHarness.mountCachedRefresh());
  const cachedRow = refreshPage.getByText("Cached record");
  await cachedRow.waitFor({ state: "visible", timeout: 2000 });
  assert.equal(await refreshPage.getByRole("status").innerText(), "cache", "cached rows are visible while the network refresh is pending");
  await refreshPage.getByText("Fresh record").waitFor({ state: "visible" });
  assert.equal(await refreshPage.getByRole("status").innerText(), "network", "background refresh replaces cached rows with current data");
  assert.equal(await refreshPage.evaluate(() => window.__spfxFastRefreshCalls), 1);
  await assertNoAxeViolations(refreshPage, "cached-first refresh");
  await refreshPage.close();

  const listPage = await createPage();
  const startedAt = await listPage.evaluate(() => performance.now());
  await listPage.evaluate(() => window.SPFxFastHarness.mountList());
  await listPage.getByRole("row").nth(1).waitFor();
  const initialRenderMs = await listPage.evaluate((start) => performance.now() - start, startedAt);
  assert.equal(await listPage.getByRole("row").count(), 11, "the list renders one ten-row page plus the header");
  assert.equal(await listPage.getByRole("columnheader").count(), 5, "the list sample includes five useful fields");
  assert.equal(await listPage.getByRole("navigation", { name: "List pagination" }).getByText("Page 1 of 100 · Showing 1–10 of 1000").count(), 1, "the list exposes current page and range");
  if (screenshotDir) await listPage.screenshot({ path: join(screenshotDir, "list-desktop.png") });
  await listPage.getByRole("button", { name: "Next page" }).click();
  assert.equal(await listPage.getByRole("row").count(), 11, "pagination renders one bounded page at a time");
  assert.equal(await listPage.getByRole("navigation", { name: "List pagination" }).getByText("Page 2 of 100 · Showing 11–20 of 1000").count(), 1, "next page updates the visible range");
  await listPage.getByRole("button", { name: "Previous page" }).click();
  await listPage.setViewportSize({ width: 375, height: 667 });
  const overflow = await listPage.evaluate(() => document.documentElement.scrollWidth > document.documentElement.clientWidth);
  assert.equal(overflow, false, "the list does not create page-level horizontal overflow at a narrow viewport");
  if (screenshotDir) await listPage.screenshot({ path: join(screenshotDir, "list-mobile.png") });
  await assertNoAxeViolations(listPage, "list at narrow viewport");
  await listPage.close();

  const startupPage = await createPage();
  await startupPage.evaluate(() => window.SPFxFastHarness.mountStartupProbe());
  await startupPage.waitForTimeout(180);
  const startupRows = await startupPage.locator("li").allInnerTexts();
  assert.match(startupRows[0] ?? "", /^Tasks first result: \d+(\.\d+)? ms \(cache\)$/, `startup output: ${await startupPage.locator("body").innerText()}; browser errors: ${browserErrors.join(" | ")}`);
  assert.match(startupRows[1] ?? "", /^Tasks network refresh: \d+(\.\d+)? ms$/);
  assert.match(startupRows[2] ?? "", /^Projects first result: \d+(\.\d+)? ms \(network\)$/);
  assert.match(startupRows[3] ?? "", /^Projects network refresh: \d+(\.\d+)? ms$/);
  await assertNoAxeViolations(startupPage, "dashboard startup timing probe");
  if (screenshotDir) await startupPage.screenshot({ path: join(screenshotDir, "dashboard-startup.png") });
  await startupPage.close();

  const drivePage = await createPage();
  await drivePage.evaluate(() => window.SPFxFastHarness.mountDriveDelta());
  await drivePage.getByRole("heading", { name: "Drive sync" }).waitFor({ state: "visible" });
  try { await drivePage.getByText("Showing cached inventory").waitFor({ state: "visible" }); }
  catch (error) { console.error(`Drive panel diagnostic: ${await drivePage.locator("body").innerText()}\nPage errors: ${browserErrors.join(" | ") || "none"}`); throw error; }
  await drivePage.getByText("Syncing changes").waitFor({ state: "visible" });
  await drivePage.getByText("Fresh file.md").waitFor({ state: "visible" });
  await drivePage.getByText("New file.txt").waitFor({ state: "visible" });
  const driveTypes = drivePage.getByRole("region", { name: "Drive file types" });
  await driveTypes.waitFor({ state: "visible" });
  assert.match(await driveTypes.innerText(), /Documents\s+2[\s\S]*Spreadsheets\s+1[\s\S]*Images\s+1/);
  await drivePage.getByRole("status").filter({ hasText: "Drive is up to date" }).waitFor({ state: "visible" });
  assert.equal(await drivePage.getByText(/6 items · 4 files · 2 folders/).count(), 1, "drive summary reports current inventory counts");
  assert.match(await drivePage.getByText(/Delta sync stores the current drive state/).innerText(), /not an edit history/);
  await assertNoAxeViolations(drivePage, "cache-first drive delta dashboard");
  if (screenshotDir) await drivePage.screenshot({ path: join(screenshotDir, "drive-delta.png"), fullPage: true });
  await drivePage.getByRole("button", { name: "Download", exact: true }).first().click();
  await drivePage.waitForFunction(() => window.__spfxDriveDownloads === 1);
  await drivePage.getByRole("button", { name: "Sync now" }).click();
  await drivePage.close();

  const libraryPage = await createPage();
  await libraryPage.evaluate(() => window.SPFxFastHarness.mountDocumentLibraryPanel());
  const libraryRegion = libraryPage.getByRole("region", { name: "Library file types" });
  try { await libraryPage.getByText("Document library", { exact: true }).waitFor(); }
  catch (error) {
    console.error(`Document library diagnostic: ${await libraryPage.locator("body").innerText()}\nPage errors: ${browserErrors.join(" | ") || "none"}`);
    throw error;
  }
  await libraryRegion.waitFor();
  assert.equal(await libraryRegion.locator("li").count(), 5, "file-type summary lists each category");
  assert.match(await libraryRegion.innerText(), /Documents\s+2[\s\S]*Spreadsheets\s+1[\s\S]*Images\s+1/);
  await libraryPage.getByRole("button", { name: "Operations-guide.md", exact: true }).click();
  const previewRegion = libraryPage.getByRole("region", { name: "Selected Markdown file" });
  await previewRegion.getByRole("heading", { name: "Operations-guide.md" }).waitFor();
  await previewRegion.getByRole("heading", { name: "Operations guide" }).waitFor();
  await previewRegion.getByRole("link", { name: "Open in SharePoint" }).waitFor();
  await libraryPage.waitForFunction(() => ![...document.querySelectorAll("button")].some((button) => button.textContent === "Refresh files" && button.disabled));
  await assertNoAxeViolations(libraryPage, "document library with Markdown preview");
  await libraryPage.getByRole("button", { name: "Download Quarterly-plan.docx" }).click();
  await libraryPage.waitForFunction(() => window.__spfxDriveDownloads === 1);
  await libraryPage.getByLabel("Upload a file").setInputFiles({ name: "incident-report.txt", mimeType: "text/plain", buffer: Buffer.from("resume test") });
  await libraryPage.getByRole("button", { name: "Resume upload" }).waitFor();
  await assertNoAxeViolations(libraryPage, "document library with resumable upload action");
  if (screenshotDir) await libraryPage.screenshot({ path: join(screenshotDir, "drive-upload-recovery.png"), fullPage: true });
  await libraryPage.getByRole("button", { name: "Resume upload" }).click();
  await libraryPage.getByRole("button", { name: "Resume upload" }).waitFor({ state: "detached" });
  if (screenshotDir) await libraryPage.screenshot({ path: join(screenshotDir, "drive-library.png"), fullPage: true });
  await libraryPage.setViewportSize({ width: 375, height: 667 });
  assert.equal(await libraryPage.evaluate(() => document.documentElement.scrollWidth > document.documentElement.clientWidth), false, "document library fits a narrow viewport");
  if (screenshotDir) await libraryPage.screenshot({ path: join(screenshotDir, "drive-library-mobile.png"), fullPage: true });
  await previewRegion.getByRole("button", { name: "Close preview" }).click();
  await previewRegion.waitFor({ state: "detached" });
  await libraryPage.close();

  const calendarPage = await createPage();
  await calendarPage.evaluate(() => window.SPFxFastHarness.mountCalendarAgenda());
  const calendarRegion = calendarPage.getByRole("region", { name: "Upcoming calendar events" });
  await calendarRegion.getByRole("heading", { name: "Upcoming events" }).waitFor();
  await calendarRegion.getByText("Planning meeting").waitFor();
  await calendarRegion.getByText("Room 4").waitFor();
  await calendarPage.waitForFunction(() => [...document.querySelectorAll("button")].some((button) => button.textContent === "Refresh" && !button.disabled));
  await calendarPage.waitForFunction(() => [...document.querySelectorAll("button")].some((button) => button.textContent === "Refresh" && !button.disabled && getComputedStyle(button).color === "rgb(36, 36, 36)"));
  await assertNoAxeViolations(calendarPage, "calendar agenda");
  await calendarRegion.getByRole("button", { name: "Load more events" }).click();
  await calendarRegion.getByText("Next review").waitFor();
  await calendarRegion.getByRole("button", { name: "Refresh" }).click();
  await calendarRegion.getByText("Planning meeting").waitFor();
  await calendarRegion.getByText("Next review").waitFor({ state: "hidden" });
  assert.equal(await calendarRegion.getByText("Next review").count(), 0, "refresh replaces events instead of duplicating the prior page");
  await calendarPage.close();

  const dashboardCalendarPage = await createPage();
  await dashboardCalendarPage.evaluate(() => window.SPFxFastHarness.mountDashboardCalendar());
  const dashboardViews = dashboardCalendarPage.getByRole("tablist", { name: "Dashboard views" });
  await dashboardCalendarPage.getByRole("region", { name: "Work dashboard" }).waitFor();
  assert.equal(await dashboardCalendarPage.evaluate(() => window.__spfxDashboardCalendarCalls), 0, "calendar does not load before the user selects that view");
  await dashboardViews.getByRole("tab", { name: "Calendar" }).click();
  await dashboardCalendarPage.getByText("Synthetic planning meeting").waitFor();
  await dashboardCalendarPage.waitForFunction(() => [...document.querySelectorAll("button")].some((button) => button.textContent === "Refresh" && !button.disabled && getComputedStyle(button).color === "rgb(36, 36, 36)"));
  assert.equal(await dashboardCalendarPage.evaluate(() => window.__spfxDashboardCalendarCalls), 1, "the SPFx dashboard calendar view loads through the provided client");
  await assertNoAxeViolations(dashboardCalendarPage, "dashboard calendar view");
  if (screenshotDir) await dashboardCalendarPage.screenshot({ path: join(screenshotDir, "dashboard-calendar.png"), fullPage: true });
  await dashboardViews.getByRole("tab", { name: "Work" }).click();
  await dashboardCalendarPage.getByRole("region", { name: "Work dashboard" }).waitFor();
  await dashboardCalendarPage.getByText("Synthetic planning meeting").waitFor({ state: "detached" });
  await assertNoAxeViolations(dashboardCalendarPage, "dashboard work view");
  await dashboardCalendarPage.close();

  const devicePage = await createPage();
  await devicePage.evaluate(() => window.SPFxFastHarness.mountDeviceCapture());
  const barcode = devicePage.getByRole("textbox", { name: "Barcode" });
  await barcode.fill("LAP-123");
  await barcode.press("Tab");
  await devicePage.waitForFunction(() => window.__spfxFastLastBarcode === "LAP-123");
  assert.equal(await devicePage.evaluate(() => window.__spfxFastGeolocationCalls), 0, "device access waits for user action");
  await assertNoAxeViolations(devicePage, "device capture form");
  await devicePage.getByLabel("Take or select a barcode photo").setInputFiles({ name: "asset.png", mimeType: "image/png", buffer: Buffer.from("synthetic image") });
  await devicePage.waitForFunction(() => window.__spfxFastLastBarcode === "ASSET-42");
  assert.equal(await devicePage.evaluate(() => window.__spfxFastBarcodeScans), 1, "photo scanning uses the browser BarcodeDetector after a user selects an image");
  assert.equal(await devicePage.evaluate(() => window.__spfxFastBitmapCloses), 1, "photo scanning closes the decoded image bitmap");
  await devicePage.getByRole("button", { name: "Capture current location" }).click();
  await devicePage.getByText(/Location captured/).waitFor();
  assert.deepEqual(await devicePage.evaluate(() => window.__spfxFastLocation), { latitude: 51.5, longitude: -0.12, accuracy: 20 });
  assert.equal(await devicePage.evaluate(() => window.__spfxFastGeolocationCalls), 1);
  await devicePage.close();

  const noScannerPage = await createPage();
  await noScannerPage.evaluate(() => window.SPFxFastHarness.mountDeviceCaptureWithoutScanner());
  await noScannerPage.getByLabel("Take or select a barcode photo").setInputFiles({ name: "asset.png", mimeType: "image/png", buffer: Buffer.from("synthetic image") });
  await noScannerPage.getByText("Barcode scanning is unavailable in this browser. Enter the code below.").waitFor();
  await noScannerPage.getByRole("textbox", { name: "Barcode" }).fill("MANUAL-7");
  await noScannerPage.getByRole("textbox", { name: "Barcode" }).press("Tab");
  assert.equal(await noScannerPage.evaluate(() => window.__spfxFastLastBarcode), "MANUAL-7", "manual entry remains available without BarcodeDetector");
  await assertNoAxeViolations(noScannerPage, "device capture without scanner support");
  await noScannerPage.close();

  const deniedLocationPage = await createPage();
  await deniedLocationPage.evaluate(() => window.SPFxFastHarness.mountDeviceCaptureLocationDenied());
  const locationButton = deniedLocationPage.getByRole("button", { name: "Capture current location" });
  await locationButton.click();
  await deniedLocationPage.getByText("Location permission denied").waitFor();
  await locationButton.waitFor({ state: "visible" });
  assert.equal(await locationButton.isEnabled(), true, "location denial clears the busy state so the user can retry");
  await assertNoAxeViolations(deniedLocationPage, "device location permission denial");
  await deniedLocationPage.close();

  const mapPage = await createPage();
  await mapPage.evaluate(() => window.SPFxFastHarness.mountFloorPlanMap());
  const mapRegion = mapPage.getByRole("region", { name: "Office locations" });
  const floorMap = mapRegion.getByRole("group", { name: "Office locations map" });
  await floorMap.waitFor();
  assert.equal(await floorMap.getByRole("button", { name: "Outside map" }).count(), 0, "locations outside normalized map coordinates are omitted");
  await mapRegion.getByRole("navigation", { name: "Office locations location list" }).getByRole("button", { name: "Room 4", exact: true }).click();
  assert.equal(await mapPage.evaluate(() => window.__spfxFastSelectedLocation), "room-4", "selecting a location reports its SharePoint item to the host");
  assert.match(await mapRegion.locator('[aria-live="polite"]').innerText(), /Room 4: Meeting room/);
  const deskMarker = floorMap.getByRole("button", { name: "Desk A" });
  await deskMarker.press("Enter");
  assert.equal(await mapPage.evaluate(() => window.__spfxFastSelectedLocation), "desk-a", "keyboard activation selects a map marker");
  assert.equal(await deskMarker.getAttribute("aria-pressed"), "true");
  await assertNoAxeViolations(mapPage, "floor-plan map");
  if (screenshotDir) await mapRegion.screenshot({ path: join(screenshotDir, "floor-plan-map.png") });
  await mapPage.setViewportSize({ width: 375, height: 667 });
  assert.equal(await mapPage.evaluate(() => document.documentElement.scrollWidth > document.documentElement.clientWidth), false, "floor-plan map fits a narrow viewport");
  await assertNoAxeViolations(mapPage, "floor-plan map at narrow viewport");
  if (screenshotDir) await mapRegion.screenshot({ path: join(screenshotDir, "floor-plan-map-mobile.png") });
  await mapPage.close();

  const driveMapPage = await createPage();
  await driveMapPage.evaluate(() => window.SPFxFastHarness.mountDriveFloorPlanMap());
  const driveMapRegion = driveMapPage.getByRole("region", { name: "Secure office plan floor plan" });
  const driveMapImage = driveMapRegion.locator('img[alt=""]');
  await driveMapImage.waitFor();
  assert.match(await driveMapImage.getAttribute("src"), /^blob:/, "private plan bytes render through a temporary local object URL");
  await driveMapPage.waitForFunction(() => document.querySelector('img[alt=""]')?.naturalWidth === 800);
  assert.equal(await driveMapPage.evaluate(() => window.__spfxFloorPlanReads), 1, "the cached authenticated file hook reads the plan once");
  const meeting = driveMapRegion.getByRole("navigation", { name: "Secure office plan location list" }).getByRole("button", { name: "Meeting 1" });
  await meeting.click();
  assert.equal(await meeting.getAttribute("aria-pressed"), "true", "the private-image map keeps its keyboard-operable location list");
  assert.match(await driveMapRegion.locator('[aria-live="polite"]').innerText(), /Meeting 1: 6 seats/);
  await assertNoAxeViolations(driveMapPage, "authenticated floor-plan image");
  if (screenshotDir) await driveMapRegion.screenshot({ path: join(screenshotDir, "drive-floor-plan-map.png") });
  await driveMapPage.evaluate(() => window.SPFxFastHarness.unmountDriveFloorPlanMap());
  assert.equal(await driveMapPage.evaluate(() => window.__spfxFloorPlanUrlsCreated), 1);
  assert.equal(await driveMapPage.evaluate(() => window.__spfxFloorPlanUrlsRevoked), 1, "unmount releases the private image object URL");
  await driveMapPage.close();

  const contentIdentityPage = await createPage();
  await contentIdentityPage.evaluate(() => window.SPFxFastHarness.mountDriveContentIdentityProbe());
  const loadedFile = contentIdentityPage.getByLabel("Loaded file");
  assert.equal(await loadedFile.innerText(), "First cached file");
  await contentIdentityPage.getByRole("button", { name: "Load another file" }).click();
  await contentIdentityPage.waitForFunction(() => window.__spfxSecondFileRequested);
  await contentIdentityPage.getByRole("status").filter({ hasText: "Loading file" }).waitFor();
  assert.equal(await loadedFile.innerText(), "", "file identity changes never show content from the previous item");
  await assertNoAxeViolations(contentIdentityPage, "file content identity change");
  await contentIdentityPage.close();

  const privatePlanPage = await createPage();
  await privatePlanPage.evaluate(() => window.SPFxFastHarness.mountPrivateFloorPlanProbe());
  privatePlanPage.once("dialog", (dialog) => dialog.accept());
  await privatePlanPage.getByRole("button", { name: "Test private floor-plan image" }).click();
  const privatePlanMap = privatePlanPage.getByRole("region", { name: "Synthetic facilities floor plan" });
  const privatePlanImage = privatePlanMap.locator('img[alt=""]');
  await privatePlanImage.waitFor();
  await privatePlanPage.waitForFunction(() => document.querySelector('img[alt=""]')?.naturalWidth === 800);
  assert.equal(await privatePlanPage.evaluate(() => window.__spfxPrivateFloorPlanUploads), 1);
  assert.equal(await privatePlanPage.evaluate(() => window.__spfxPrivateFloorPlanReads), 1);
  await assertNoAxeViolations(privatePlanPage, "private floor-plan tenant probe");
  await privatePlanPage.getByRole("button", { name: "Remove temporary floor-plan image" }).click();
  await privatePlanPage.getByText("Temporary image removed.").waitFor();
  assert.equal(await privatePlanPage.evaluate(() => window.__spfxPrivateFloorPlanDeletes), 1);
  await privatePlanPage.close();

  const attachmentPage = await createPage();
  await attachmentPage.evaluate(() => window.SPFxFastHarness.mountItemAttachmentsPanel());
  await attachmentPage.getByRole("heading", { name: "Item attachments" }).waitFor();
  const attachmentList = attachmentPage.getByRole("list", { name: "Item attachments" });
  await attachmentList.locator("li").filter({ hasText: "meeting-notes.txt" }).waitFor();
  await attachmentPage.getByRole("button", { name: "Download meeting-notes.txt" }).click();
  await attachmentPage.waitForFunction(() => window.__spfxAttachmentDownloads === 1);
  const attachmentInput = attachmentPage.getByLabel("Choose a file to attach");
  await attachmentInput.setInputFiles({ name: "response.txt", mimeType: "text/plain", buffer: Buffer.from("Synthetic attachment") });
  await attachmentPage.getByRole("button", { name: "Add attachment" }).click();
  const uploadedAttachment = attachmentList.locator("li").filter({ hasText: "response.txt" });
  await uploadedAttachment.waitFor();
  await attachmentPage.waitForFunction(() => ![...document.querySelectorAll("button")].some((button) => button.textContent === "Refresh attachments" && button.disabled));
  assert.equal(await attachmentPage.evaluate(() => window.__spfxAttachmentMutations), 1, "upload creates the selected attachment");
  await assertNoAxeViolations(attachmentPage, "list-item attachment workflow");
  if (screenshotDir) await attachmentPage.screenshot({ path: join(screenshotDir, "list-attachments.png"), fullPage: true });
  await attachmentInput.setInputFiles({ name: "meeting-notes.txt", mimeType: "text/plain", buffer: Buffer.from("Updated notes") });
  await attachmentPage.getByRole("button", { name: "Replace attachment" }).click();
  await attachmentPage.getByRole("status").filter({ hasText: "meeting-notes.txt replaced." }).waitFor();
  assert.equal(await attachmentPage.evaluate(() => window.__spfxAttachmentMutations), 2, "same-name upload replaces the existing attachment");
  await attachmentPage.getByRole("button", { name: "Delete response.txt" }).click();
  await attachmentPage.getByRole("button", { name: "Delete now" }).click();
  await uploadedAttachment.waitFor({ state: "detached" });
  assert.equal(await attachmentPage.evaluate(() => window.__spfxAttachmentMutations), 3, "delete requires a second explicit action");
  await attachmentPage.close();

  const elementPage = await createPage();
  await elementPage.evaluate(() => window.SPFxFastHarness.mountCustomElement());
  await elementPage.getByRole("status").filter({ hasText: "1 items loaded" }).waitFor({ state: "visible" });
  assert.equal(await elementPage.getByRole("row").count(), 2, "custom element respects the host's max-items limit");
  assert.equal(await elementPage.locator("spfx-fast-browser-list strong").innerText(), "Rich text title", "SharePoint rich text keeps safe formatting");
  assert.equal(await elementPage.locator("spfx-fast-browser-list img, spfx-fast-browser-list script, spfx-fast-browser-list [onerror]").count(), 0, "unsafe SharePoint markup is removed");
  assert.equal(await elementPage.locator('spfx-fast-browser-list a[href^="javascript:"]').count(), 0, "unsafe SharePoint links are not rendered");
  assert.equal(await elementPage.evaluate(() => window.__spfxFastXss), undefined, "rich-text values cannot execute scripts");
  assert.match(await elementPage.locator("spfx-fast-browser-list").evaluate((element) => getComputedStyle(element).fontFamily), /system-ui/, "custom element preserves its accessible host styling after refresh");
  const refresh = elementPage.getByRole("button", { name: "Refresh" });
  await refresh.focus();
  await elementPage.keyboard.press("Enter");
  await elementPage.waitForFunction(() => window.__spfxFastListCalls === 2);
  await assertNoAxeViolations(elementPage, "custom element");
  if (screenshotDir) await elementPage.screenshot({ path: join(screenshotDir, "custom-element.png") });
  await elementPage.close();

  const failoverPage = await createPage();
  await failoverPage.evaluate(() => window.SPFxFastHarness.mountCacheFailover());
  await failoverPage.getByRole("cell", { name: "Available from SharePoint" }).waitFor({ state: "visible" });
  assert.equal(await failoverPage.evaluate(() => window.__spfxFastStorageErrors), 1, "primary cache failures are surfaced while memory fallback keeps data available");
  await assertNoAxeViolations(failoverPage, "memory cache failover");
  await failoverPage.close();

  const quotaPage = await createPage();
  await quotaPage.evaluate(() => window.SPFxFastHarness.mountQuotaDenial());
  await quotaPage.getByRole("status").filter({ hasText: "Quota fallback source: cache" }).waitFor({ state: "visible" });
  assert.deepEqual(await quotaPage.evaluate(() => window.__spfxFastQuotaResult), {
    firstSource: "network",
    secondSource: "cache",
    storageErrorName: "QuotaExceededError",
  }, "quota-denied writes report the storage exception and subsequent reads use the memory cache");
  await assertNoAxeViolations(quotaPage, "quota-denial memory cache fallback");
  await quotaPage.close();

  assert.deepEqual(browserErrors, [], "browser should not emit uncaught errors");
  console.log(`Browser checks passed: React and custom-element keyboard flows, dashboard work/calendar view switching, simulated tenant-panel read/cache/delta/CRUD/file flows, four sample webparts, request form attachment integration, sanitized Markdown/GFM and unsafe-content checks, file-library and list-attachment actions, calendar paging, barcode photo/manual fallback and location grant/denial flows, keyboard floor-plan selection and coordinate bounds, file-type charts, approval updates, cached-first list and drive refresh, dashboard startup timings, generic and quota-denial memory cache failover, bounded rendering, ${axeScanCount} axe scans, and narrow layouts; initial render ${initialRenderMs.toFixed(1)} ms.`);
} finally {
  await browser.close();
  rmSync(scratch, { recursive: true, force: true });
}
