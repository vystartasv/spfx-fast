import { mkdir } from "node:fs/promises";
import { join } from "node:path";

const actions = new Set(["navigate", "assertPageReady", "assertAuthenticated", "click", "fill", "select", "setChecked", "assertVisible", "assertText", "assertValue", "screenshot"]);
const locatorKeys = new Set(["role", "text", "label", "testId", "placeholder"]);

export function validateRecipe(recipe) {
  if (!recipe || typeof recipe !== "object" || Array.isArray(recipe)) throw new TypeError("Recipe must be a JSON object.");
  if (typeof recipe.name !== "string" || !recipe.name.trim()) throw new TypeError("Recipe needs a non-empty name.");
  if (!Array.isArray(recipe.steps) || recipe.steps.length === 0) throw new TypeError("Recipe needs at least one step.");

  for (const [index, step] of recipe.steps.entries()) {
    if (!step || typeof step !== "object" || Array.isArray(step) || !actions.has(step.action)) {
      throw new TypeError(`Recipe step ${index + 1} has an unsupported action.`);
    }
    if (step.action === "navigate" && (typeof step.path !== "string" || !step.path.startsWith("/") || step.path.startsWith("//") || /[?#]/.test(step.path))) {
      throw new TypeError(`Recipe step ${index + 1} needs a same-site path beginning with / and no query or fragment.`);
    }
    if (step.action === "navigate" && step.query !== undefined) {
      if (!step.query || typeof step.query !== "object" || Array.isArray(step.query) || Object.keys(step.query).length > 16) {
        throw new TypeError(`Recipe step ${index + 1} query must be an object with at most 16 parameters.`);
      }
      for (const [key, value] of Object.entries(step.query)) {
        if (!/^[A-Za-z][A-Za-z0-9_.-]{0,63}$/.test(key) || typeof value !== "string" || value.length > 2048) {
          throw new TypeError(`Recipe step ${index + 1} has an invalid query parameter.`);
        }
        if (key.toLowerCase() === "debugmanifestsfile") {
          let manifest;
          try { manifest = new URL(value); } catch { throw new TypeError("debugManifestsFile must use a local HTTPS URL."); }
          if (manifest.protocol !== "https:" || !["localhost", "127.0.0.1"].includes(manifest.hostname) || manifest.username || manifest.password) {
            throw new TypeError("debugManifestsFile must use a local HTTPS URL.");
          }
        }
      }
    }
    if (["click", "fill", "select", "setChecked", "assertVisible", "assertText", "assertValue"].includes(step.action)) validateLocator(step.locator, index + 1);
    if (["fill", "select"].includes(step.action) && typeof step.value !== "string") throw new TypeError(`Recipe step ${index + 1} needs a string value.`);
    if (step.confirm !== undefined && (step.action !== "click" || step.confirm !== true)) {
      throw new TypeError(`Recipe step ${index + 1} can only set confirm: true on a click action.`);
    }
    if (step.action === "setChecked" && typeof step.checked !== "boolean") throw new TypeError(`Recipe step ${index + 1} needs a boolean checked value.`);
    if (["assertText", "assertValue"].includes(step.action) && typeof step.expected !== "string") throw new TypeError(`Recipe step ${index + 1} needs expected text.`);
    if (step.action === "screenshot" && (typeof step.name !== "string" || !/^[\w-]+$/.test(step.name))) {
      throw new TypeError(`Recipe step ${index + 1} needs a screenshot name using only letters, numbers, _ or -.`);
    }
  }
  return recipe;
}

function validateLocator(locator, stepNumber) {
  if (!locator || typeof locator !== "object" || Array.isArray(locator)) throw new TypeError(`Recipe step ${stepNumber} needs a locator.`);
  const keys = Object.keys(locator).filter((key) => locatorKeys.has(key));
  if (keys.length !== 1 || typeof locator[keys[0]] !== "string" || !locator[keys[0]].trim()) {
    throw new TypeError(`Recipe step ${stepNumber} needs exactly one non-empty role, text, label, testId or placeholder locator.`);
  }
  if (keys[0] === "role" && (typeof locator.name !== "string" || !locator.name.trim())) {
    throw new TypeError(`Recipe step ${stepNumber} role locator needs an accessible name.`);
  }
}

export function validateBaseUrl(baseUrl) {
  let url;
  try { url = new URL(baseUrl); } catch { throw new TypeError("Provide a valid SPFX_BASE_URL."); }
  const secure = url.protocol === "https:" || (url.protocol === "http:" && ["localhost", "127.0.0.1"].includes(url.hostname));
  if (!secure) throw new TypeError("SPFX_BASE_URL must use HTTPS (HTTP is allowed only for localhost).");
  if (url.username || url.password || url.search || url.hash) throw new TypeError("SPFX_BASE_URL must not contain credentials, query values or a fragment.");
  return url;
}

function locatorFor(page, locator) {
  if (locator.role) return page.getByRole(locator.role, { name: locator.name, exact: locator.exact ?? true });
  if (locator.text) return page.getByText(locator.text, { exact: locator.exact ?? true });
  if (locator.label) return page.getByLabel(locator.label, { exact: locator.exact ?? true });
  if (locator.testId) return page.getByTestId(locator.testId);
  return page.getByPlaceholder(locator.placeholder, { exact: locator.exact ?? true });
}

function assertSameOrigin(page, base) {
  const current = new URL(page.url());
  if (current.origin !== base.origin) throw new Error("Navigation left the configured SharePoint origin; check sign-in state and page links.");
}

export async function runRecipe(page, recipe, { baseUrl, screenshotDir, timeoutMs = 15000, fast = false } = {}) {
  validateRecipe(recipe);
  const base = validateBaseUrl(baseUrl);
  const screenshots = [];
  if (screenshotDir) await mkdir(screenshotDir, { recursive: true });

  for (const [index, step] of recipe.steps.entries()) {
    try {
      if (step.action === "navigate") {
        const destination = new URL(step.path, base);
        if (destination.origin !== base.origin) throw new Error("Navigation must stay on the configured SharePoint origin.");
        for (const [key, value] of Object.entries(step.query ?? {})) destination.searchParams.set(key, value);
        const response = await page.goto(destination.href, { waitUntil: "domcontentloaded", timeout: timeoutMs });
        if (response && response.status() >= 400) throw new Error(`Page returned HTTP ${response.status()}.`);
        assertSameOrigin(page, base);
      } else if (step.action === "assertPageReady") {
        await page.waitForLoadState("domcontentloaded", { timeout: timeoutMs });
        await page.locator("body").waitFor({ state: "visible", timeout: timeoutMs });
        assertSameOrigin(page, base);
      } else if (step.action === "assertAuthenticated") {
        const sitePath = base.pathname.replace(/\/$/, "");
        const apiPath = `${sitePath}/_api/web/currentuser?$select=Id`;
        const result = await page.evaluate(async (path) => {
          const response = await fetch(path, {
            credentials: "same-origin",
            headers: { Accept: "application/json;odata=nometadata" },
          });
          return { status: response.status, contentType: response.headers.get("content-type") ?? "" };
        }, apiPath);
        if (result.status < 200 || result.status >= 300 || !result.contentType.toLowerCase().includes("json")) {
          throw new Error(`Signed-in SharePoint REST check failed (HTTP ${result.status}).`);
        }
        assertSameOrigin(page, base);
      } else if (step.action === "click") {
        if (step.confirm) {
          let dialogSeen = false;
          let dialogError;
          let dialogHandling = Promise.resolve();
          const handleDialog = (dialog) => {
            dialogSeen = true;
            dialogHandling = dialog.type() === "confirm"
              ? dialog.accept().catch(() => { dialogError = "The browser confirmation could not be accepted."; })
              : dialog.dismiss().then(() => { dialogError = "Expected a browser confirmation dialog."; });
          };
          page.once("dialog", handleDialog);
          try {
            await locatorFor(page, step.locator).click({ timeout: timeoutMs });
            await dialogHandling;
            if (!dialogSeen) throw new Error("Expected a browser confirmation dialog, but none appeared.");
            if (dialogError) throw new Error(dialogError);
          } finally {
            page.off("dialog", handleDialog);
          }
        } else {
          await locatorFor(page, step.locator).click({ timeout: timeoutMs });
        }
        assertSameOrigin(page, base);
      } else if (step.action === "fill") {
        await locatorFor(page, step.locator).fill(step.value, { timeout: timeoutMs });
        assertSameOrigin(page, base);
      } else if (step.action === "select") {
        await locatorFor(page, step.locator).selectOption(step.value, { timeout: timeoutMs });
        assertSameOrigin(page, base);
      } else if (step.action === "setChecked") {
        await locatorFor(page, step.locator).setChecked(step.checked, { timeout: timeoutMs });
        assertSameOrigin(page, base);
      } else if (step.action === "assertVisible") {
        await locatorFor(page, step.locator).waitFor({ state: "visible", timeout: timeoutMs });
        assertSameOrigin(page, base);
      } else if (step.action === "assertText") {
        const target = locatorFor(page, step.locator);
        await target.waitFor({ state: "visible", timeout: timeoutMs });
        const actual = (await target.innerText()).replace(/\s+/g, " ").trim();
        const expected = step.expected.replace(/\s+/g, " ").trim();
        if (!(step.exact ? actual === expected : actual.includes(expected))) throw new Error(`Text assertion failed (${step.exact ? "exact" : "contains"}).`);
        assertSameOrigin(page, base);
      } else if (step.action === "assertValue") {
        const target = locatorFor(page, step.locator);
        await target.waitFor({ state: "visible", timeout: timeoutMs });
        const actual = await target.inputValue();
        if (!(step.exact === false ? actual.includes(step.expected) : actual === step.expected)) throw new Error(`Value assertion failed (${step.exact === false ? "contains" : "exact"}).`);
        assertSameOrigin(page, base);
      } else if (step.action === "screenshot") {
        if (fast) continue;
        if (!screenshotDir) throw new Error("Screenshots need an artifact directory.");
        assertSameOrigin(page, base);
        const file = join(screenshotDir, `${step.name}.png`);
        await page.screenshot({ path: file, fullPage: step.fullPage ?? false });
        screenshots.push(file);
      }
    } catch (cause) {
      const error = new Error(`Recipe step ${index + 1} (${step.action}) failed.`, { cause });
      error.step = index + 1;
      throw error;
    }
  }
  return { name: recipe.name, steps: recipe.steps.length, screenshots };
}

export function generateTestSource(recipe) {
  validateRecipe(recipe);
  return `import { test } from "@playwright/test";\nimport { runRecipe } from "@spfx-fast/playwright";\n\nconst recipe = ${JSON.stringify(recipe, null, 2)};\n\ntest(recipe.name, async ({ page }, testInfo) => {\n  const baseUrl = process.env.SPFX_BASE_URL;\n  if (!baseUrl) throw new Error("Set SPFX_BASE_URL to the SharePoint site under test.");\n  await runRecipe(page, recipe, {\n    baseUrl,\n    screenshotDir: testInfo.outputPath("spfx-fast-screenshots"),\n  });\n});\n`;
}
