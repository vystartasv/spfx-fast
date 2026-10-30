import assert from "node:assert/strict";
import { createServer } from "node:http";
import { spawn } from "node:child_process";
import { mkdir, mkdtemp, readFile, readdir, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";
import { generateTestSource, runRecipe, validateBaseUrl, validateRecipe } from "../packages/playwright/src/index.js";
import { acquireProfileLock } from "../packages/playwright/src/profile-lock.js";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));

test("recipe validation accepts semantic same-origin steps and generates native Playwright code", () => {
  const recipe = {
    name: "Dashboard smoke",
    steps: [
      { action: "navigate", path: "/sites/sample", query: { view: "dashboard" } },
      { action: "assertPageReady" },
      { action: "assertVisible", locator: { role: "heading", name: "Dashboard" } },
    ],
  };
  assert.equal(validateRecipe(recipe), recipe);
  assert.match(generateTestSource(recipe), /from "@playwright\/test"/);
  assert.match(generateTestSource(recipe), /from "@spfx-fast\/playwright"/);
  assert.match(generateTestSource(recipe), /SPFX_BASE_URL/);
  assert.doesNotMatch(generateTestSource(recipe), /@spfx-fast\/core/);
  assert.throws(() => validateRecipe({ name: "Unsafe manifest", steps: [{ action: "navigate", path: "/sites/sample", query: { debugManifestsFile: "https://example.com/manifest.js" } }] }), /debugManifestsFile must use a local HTTPS URL/);
});

test("tenant CRUD recipe is valid, asserts the full operation result, and accepts only its confirmation", async () => {
  const recipe = JSON.parse(await readFile(join(root, "packages/playwright/examples/tenant-crud-validation.recipe.json"), "utf8"));
  assert.equal(validateRecipe(recipe), recipe);
  assert.ok(recipe.steps.some((step) => step.action === "click" && step.confirm === true));
  assert.ok(recipe.steps.some((step) => step.action === "assertText" && step.expected.includes("Create, read, update, read, and delete")));
  assert.match(generateTestSource(recipe), /tenant-validation\.aspx/);
});

test("init scaffolds runnable tests and refuses to overwrite an existing project", async () => {
  const project = await mkdtemp(join(tmpdir(), "spfx-fast-playwright-init-"));
  const cli = join(root, "packages/playwright/bin/spfx-fast-playwright.mjs");
  try {
    await writeFile(join(project, "AGENTS.md"), "# Existing instructions\n\nKeep this content.\n");
    await runProcess(process.execPath, [cli, "init"], { cwd: project, env: process.env });
    const spec = await readFile(join(project, "tests/e2e/sharepoint-smoke.spec.mjs"), "utf8");
    const config = await readFile(join(project, "playwright.config.ts"), "utf8");
    const authIgnore = await readFile(join(project, "playwright/.auth/.gitignore"), "utf8");
    const agentGuide = await readFile(join(project, "SPFX-PLAYWRIGHT.md"), "utf8");
    const agents = await readFile(join(project, "AGENTS.md"), "utf8");
    assert.match(spec, /assertAuthenticated/);
    assert.match(spec, /@playwright\/test/);
    assert.match(config, /workers: 1/);
    assert.match(config, /retain-on-failure/);
    assert.match(authIgnore, /^\*/);
    assert.match(agentGuide, /never change production data/i);
    assert.match(agents, /spfx-fast-playwright:agent-instructions/);
    assert.match(agents, /SPFX-PLAYWRIGHT\.md/);
    assert.match(agents, /^# Existing instructions[\s\S]*Keep this content\./);
    await assert.rejects(runProcess(process.execPath, [cli, "init"], { cwd: project, env: process.env }), /Refusing to overwrite/);
    await runProcess(process.execPath, [cli, "init", "--force"], { cwd: project, env: process.env });
    const agentsAfterForce = await readFile(join(project, "AGENTS.md"), "utf8");
    assert.equal(agentsAfterForce.split("spfx-fast-playwright:agent-instructions").length - 1, 1, "re-running init must not duplicate the agent link");
    assert.match(agentsAfterForce, /Keep this content\./);
  } finally {
    await rm(project, { recursive: true, force: true });
  }
});

test("generated spec runs in native Playwright Test and saves its screenshot", async () => {
  const server = createServer((request, response) => {
    if (request.url === "/_api/web/currentuser?$select=Id") {
      response.writeHead(200, { "content-type": "application/json;odata=nometadata" });
      response.end('{"Id":1}');
      return;
    }
    response.writeHead(200, { "content-type": "text/html; charset=utf-8" });
    response.end('<!doctype html><html lang="en"><body><h1>SPFx fixture</h1><p>Ready</p></body></html>');
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const address = server.address();
  await mkdir(join(root, "test-results"), { recursive: true });
  const workspace = await mkdtemp(join(root, "test-results", "generated-playwright-"));
  const packageScope = join(workspace, "node_modules", "@spfx-fast");
  const specPath = join(workspace, "generated.spec.mjs");
  const configPath = join(workspace, "playwright.config.mjs");
  try {
    await mkdir(packageScope, { recursive: true });
    await symlink(resolve(root, "packages/playwright"), join(packageScope, "playwright"), process.platform === "win32" ? "junction" : "dir");
    const recipe = {
      name: "Generated SPFx fixture test",
      steps: [
        { action: "navigate", path: "/fixture" },
        { action: "assertPageReady" },
        { action: "assertAuthenticated" },
        { action: "assertVisible", locator: { role: "heading", name: "SPFx fixture" } },
        { action: "screenshot", name: "fixture" },
      ],
    };
    await writeFile(specPath, generateTestSource(recipe));
    await writeFile(configPath, `import { defineConfig } from "@playwright/test";\nexport default defineConfig({ testDir: ${JSON.stringify(workspace)}, testMatch: "**/*.spec.mjs", outputDir: ${JSON.stringify(join(workspace, "results"))}, workers: 1, reporter: "list" });\n`);
    await runProcess(process.execPath, [join(root, "node_modules/playwright/cli.js"), "test", "--config", configPath, specPath], {
      cwd: root,
      env: { ...process.env, SPFX_BASE_URL: `http://127.0.0.1:${address.port}` },
    });
    const resultFiles = await findFiles(join(workspace, "results"));
    assert.ok(resultFiles.some((file) => file.endsWith("fixture.png")), "generated test should save its screenshot");
  } finally {
    await new Promise((resolve) => server.close(resolve));
    await rm(workspace, { recursive: true, force: true });
  }
});

async function findFiles(directory) {
  let entries;
  try { entries = await readdir(directory, { withFileTypes: true }); } catch { return []; }
  const files = [];
  for (const entry of entries) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) files.push(...await findFiles(path));
    else files.push(path);
  }
  return files;
}

function runProcess(command, args, options) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { ...options, stdio: ["ignore", "pipe", "pipe"] });
    let stdout = "";
    let stderr = "";
    const timer = setTimeout(() => child.kill("SIGTERM"), 60000);
    child.stdout.setEncoding("utf8").on("data", (chunk) => { stdout += chunk; });
    child.stderr.setEncoding("utf8").on("data", (chunk) => { stderr += chunk; });
    child.once("error", (error) => { clearTimeout(timer); reject(error); });
    child.once("close", (code) => {
      clearTimeout(timer);
      if (code === 0) resolve();
      else reject(new Error(`Generated Playwright test failed (exit ${code}).\n${stdout}\n${stderr}`));
    });
  });
}

test("recipe runner drives a local SPFx-like page and captures a reviewable screenshot", async () => {
  let authMode = "valid";
  let receivedQuery;
  const server = createServer((request, response) => {
    if (request.url?.startsWith("/sites/sample?")) receivedQuery = new URL(request.url, "http://localhost").searchParams.get("view");
    if (request.url === "/_api/web/currentuser?$select=Id") {
      response.writeHead(200, { "content-type": authMode === "valid" ? "application/json;odata=nometadata" : "text/html" });
      response.end(authMode === "valid" ? '{"Id":1}' : "Sign in");
      return;
    }
    response.writeHead(200, { "content-type": "text/html; charset=utf-8" });
    response.end(`<!doctype html><html lang="en"><title>Test dashboard</title><body>
      <main><h1>Requests</h1><label>Request title<input aria-label="Request title"></label>
      <button>Save request</button><p role="status" hidden></p></main>
      <script>document.querySelector('button').addEventListener('click', () => {
        const status = document.querySelector('[role=status]'); status.hidden = false;
        status.textContent = 'Saved: ' + document.querySelector('input').value;
      });</script></body></html>`);
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const address = server.address();
  const baseUrl = `http://127.0.0.1:${address.port}`;
  const artifacts = await mkdtemp(join(tmpdir(), "spfx-fast-playwright-"));
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
    const result = await runRecipe(page, {
      name: "Save request smoke",
      steps: [
        { action: "navigate", path: "/sites/sample", query: { view: "dashboard" } },
        { action: "assertPageReady" },
        { action: "assertAuthenticated" },
        { action: "assertVisible", locator: { role: "heading", name: "Requests" } },
        { action: "fill", locator: { label: "Request title" }, value: "Laptop request" },
        { action: "click", locator: { role: "button", name: "Save request" } },
        { action: "assertText", locator: { text: "Saved: Laptop request", exact: true }, expected: "Saved: Laptop request", exact: true },
        { action: "screenshot", name: "request-saved" },
      ],
    }, { baseUrl, screenshotDir: artifacts });
    assert.equal(result.steps, 8);
    assert.equal(receivedQuery, "dashboard", "recipe query parameters are added without changing the site origin");
    assert.equal(result.screenshots.length, 1);
    const image = await readFile(result.screenshots[0]);
    assert.deepEqual(image.subarray(0, 8), Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));

    authMode = "login";
    await assert.rejects(runRecipe(page, {
      name: "Reject sign-in redirect",
      steps: [
        { action: "navigate", path: "/sites/sample" },
        { action: "assertPageReady" },
        { action: "assertAuthenticated" },
      ],
    }, { baseUrl, screenshotDir: artifacts }), /Recipe step 3 \(assertAuthenticated\) failed/);
  } finally {
    await browser.close();
    await new Promise((resolve) => server.close(resolve));
    await rm(artifacts, { recursive: true, force: true });
  }
});

test("recipe runner accepts only an explicitly requested browser confirmation", async () => {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  try {
    await page.route("http://127.0.0.1:3000/**", (route) => route.fulfill({
      contentType: "text/html",
      body: `<button onclick="if (confirm('Create test record?')) document.querySelector('[role=status]').textContent='Record created'">Create test record</button><button>No confirmation</button><p role="status"></p>`,
    }));
    const result = await runRecipe(page, {
      name: "Confirmed test write",
      steps: [
        { action: "navigate", path: "/fixture" },
        { action: "click", locator: { role: "button", name: "Create test record" }, confirm: true },
        { action: "assertText", locator: { text: "Record created" }, expected: "Record created", exact: true },
      ],
    }, { baseUrl: "http://127.0.0.1:3000" });
    assert.equal(result.steps, 3);
    await assert.rejects(runRecipe(page, {
      name: "Missing confirmation",
      steps: [
        { action: "navigate", path: "/fixture" },
        { action: "click", locator: { role: "button", name: "No confirmation" }, confirm: true },
      ],
    }, { baseUrl: "http://127.0.0.1:3000" }), (error) => error.cause?.message.includes("Expected a browser confirmation dialog"));
  } finally { await browser.close(); }
});

test("fast recipe runs skip success screenshots but preserve step checks", async () => {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  try {
    await page.route("http://127.0.0.1:3000/**", (route) => route.fulfill({ body: "<h1>Requests</h1>", contentType: "text/html" }));
    const result = await runRecipe(page, {
      name: "Fast check",
      steps: [
        { action: "navigate", path: "/fast" },
        { action: "assertVisible", locator: { role: "heading", name: "Requests" } },
        { action: "screenshot", name: "skip-on-success" },
      ],
    }, { baseUrl: "http://127.0.0.1:3000", fast: true });
    assert.equal(result.steps, 3);
    assert.deepEqual(result.screenshots, []);
  } finally { await browser.close(); }
});

test("persistent profile lock rejects concurrent package jobs and releases cleanly", async () => {
  const profile = await mkdtemp(join(tmpdir(), "spfx-fast-profile-lock-"));
  const release = acquireProfileLock(profile);
  try { assert.throws(() => acquireProfileLock(profile), /Another SPFx Fast Playwright job/); }
  finally { release(); }
  const releaseAgain = acquireProfileLock(profile);
  releaseAgain();
  await rm(profile, { recursive: true, force: true });
});

test("validation rejects unsafe origins, profile-like URLs, and ambiguous locators", () => {
  assert.throws(() => validateBaseUrl("http://example.com"), /HTTPS/);
  assert.throws(() => validateBaseUrl("https://tenant.sharepoint.com/?code=secret"), /query/);
  assert.throws(() => validateRecipe({ name: "unsafe", steps: [{ action: "navigate", path: "//outside.example" }] }), /same-site/);
  assert.throws(() => validateRecipe({ name: "bad locator", steps: [{ action: "click", locator: { role: "button", text: "Save" } }] }), /exactly one/);
  assert.throws(() => validateRecipe({ name: "unsafe confirmation", steps: [{ action: "click", locator: { text: "Save" }, confirm: false }] }), /confirm: true/);
  assert.throws(() => validateRecipe({ name: "misplaced confirmation", steps: [{ action: "fill", locator: { label: "Name" }, value: "test", confirm: true }] }), /confirm: true/);
  assert.throws(() => validateRecipe({ name: "bad screenshot", steps: [{ action: "screenshot", name: "../private" }] }), /screenshot name/);
});
