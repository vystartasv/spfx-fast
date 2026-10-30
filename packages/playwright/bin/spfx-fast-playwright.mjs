#!/usr/bin/env node
import { access, mkdir, readFile, readdir, stat, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { spawnSync } from "node:child_process";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";
import { parseArgs } from "node:util";
import { generateTestSource, runRecipe, validateBaseUrl, validateRecipe } from "../src/index.js";
import { acquireProfileLock } from "../src/profile-lock.js";

const help = `SPFx Fast Playwright\n\nCommands:\n  init [--force]                            Scaffold config and a smoke recipe\n  validate <recipe.json>                    Validate a recipe\n  generate <recipe.json> --out <test.spec.mjs>  Generate a native Playwright test\n  run <recipe.json> --base-url <url> --profile-dir <dir> [--artifacts <dir>] [--fast] [--headed]\n  record --url <url> --profile-dir <dir> --out <test.spec.ts>  Start Playwright Codegen\n\nRuns are headless by default. --fast skips success screenshots and traces; do one full run after fixes.\n`;

const parsed = parseArgs({
  args: process.argv.slice(2),
  allowPositionals: true,
  strict: true,
  options: {
    out: { type: "string" },
    "base-url": { type: "string" },
    "profile-dir": { type: "string" },
    artifacts: { type: "string" },
    "timeout-ms": { type: "string" },
    url: { type: "string" },
    headless: { type: "boolean", default: false },
    headed: { type: "boolean", default: false },
    fast: { type: "boolean", default: false },
    force: { type: "boolean", default: false },
    help: { type: "boolean", short: "h", default: false },
  },
});

const [command, recipeFile] = parsed.positionals;
const options = parsed.values;

async function readRecipe(file) {
  if (!file) throw new Error("Supply a recipe JSON file.");
  try { return validateRecipe(JSON.parse(await readFile(resolve(file), "utf8"))); }
  catch (error) {
    if (error instanceof TypeError) throw error;
    throw new Error("Could not read a valid recipe JSON file.");
  }
}

async function outputPath(file) {
  if (!file) throw new Error("Supply --out.");
  const target = resolve(file);
  if (!options.force) {
    try { await access(target); throw new Error("Output already exists; pass --force to replace it."); }
    catch (error) { if (error.message.includes("already exists")) throw error; }
  }
  await mkdir(dirname(target), { recursive: true });
  return target;
}

async function requireProfileDirectory(value) {
  if (!value) throw new Error("Supply an existing --profile-dir. The command will not create or copy a signed-in profile.");
  const path = resolve(value);
  let metadata;
  try { metadata = await stat(path); } catch { throw new Error("The supplied profile directory does not exist."); }
  if (!metadata.isDirectory()) throw new Error("The supplied profile path must be a directory.");
  return path;
}

async function main() {
  if (options.help || !command || command === "help") {
    console.log(help);
    return;
  }
  if (command === "init") {
    const agentGuide = await readFile(new URL("../agent/PLAYWRIGHT.md", import.meta.url), "utf8");
    const smokeRecipe = {
      name: "SharePoint page is available to the signed-in user",
      steps: [
        { action: "navigate", path: "/sites/your-site" },
        { action: "assertPageReady" },
        { action: "assertAuthenticated" },
        { action: "screenshot", name: "sharepoint-page", fullPage: true },
      ],
    };
    const files = [
      ["playwright.config.ts", `import { defineConfig } from "@playwright/test";\n\nexport default defineConfig({\n  testDir: "./tests/e2e",\n  fullyParallel: false,\n  workers: 1,\n  reporter: [["list"], ["html", { open: "never" }]],\n  use: {\n    baseURL: process.env.SPFX_BASE_URL,\n    storageState: process.env.SPFX_STORAGE_STATE || undefined,\n    actionTimeout: 15_000,\n    navigationTimeout: 45_000,\n    screenshot: "only-on-failure",\n    trace: "retain-on-failure",\n    viewport: { width: 1440, height: 900 },\n  },\n});\n`],
      ["tests/e2e/sharepoint-smoke.recipe.json", `${JSON.stringify(smokeRecipe, null, 2)}\n`],
      ["tests/e2e/sharepoint-smoke.spec.mjs", generateTestSource(smokeRecipe)],
      ["playwright/.auth/.gitignore", "*\n!.gitignore\n"],
      ["SPFX-PLAYWRIGHT.md", agentGuide],
    ];
    if (!options.force) {
      for (const [file] of files) {
        try { await access(resolve(file)); throw new Error(`Refusing to overwrite ${file}; pass --force to replace generated files.`); }
        catch (error) { if (error.message.startsWith("Refusing")) throw error; }
      }
    }
    for (const [file, contents] of files) {
      const target = resolve(file);
      await mkdir(dirname(target), { recursive: true });
      await writeFile(target, contents, "utf8");
    }
    const agentsFile = resolve("AGENTS.md");
    let existingAgents = "";
    try { existingAgents = await readFile(agentsFile, "utf8"); }
    catch (error) { if (error.code !== "ENOENT") throw error; }
    const marker = "<!-- spfx-fast-playwright:agent-instructions -->";
    if (!existingAgents.includes(marker)) {
      const link = `${marker}\n## SharePoint end-to-end tests\nRead [SPFX-PLAYWRIGHT.md](SPFX-PLAYWRIGHT.md) when creating or running authenticated SharePoint browser tests.\n`;
      const prefix = existingAgents ? `${existingAgents}${existingAgents.endsWith("\n") ? "" : "\n"}\n` : "# Agent instructions\n\n";
      await writeFile(agentsFile, `${prefix}${link}`, "utf8");
    }
    console.log("Created Playwright config, a generated smoke test, agent instructions, and an ignored auth-state folder.");
    console.log("Set SPFX_BASE_URL to the SharePoint web URL. Keep SPFX_STORAGE_STATE private.");
    return;
  }
  if (command === "validate") {
    const recipe = await readRecipe(recipeFile);
    console.log(`Recipe valid: ${recipe.name} (${recipe.steps.length} steps).`);
    return;
  }
  if (command === "generate") {
    const recipe = await readRecipe(recipeFile);
    const target = await outputPath(options.out);
    if (!/\.spec\.(?:mjs|ts)$/.test(target)) throw new Error("Generated tests must end in .spec.mjs or .spec.ts.");
    await writeFile(target, generateTestSource(recipe), "utf8");
    console.log(`Generated native Playwright test: ${target}`);
    return;
  }
  if (command === "record") {
    const profile = await requireProfileDirectory(options["profile-dir"]);
    const output = await outputPath(options.out);
    if (!/\.spec\.(?:mjs|ts)$/.test(output)) throw new Error("Recorded tests must end in .spec.mjs or .spec.ts.");
    const url = validateBaseUrl(options.url).href;
    const requireFromProject = createRequire(pathToFileURL(resolve("package.json")));
    const playwrightEntry = requireFromProject.resolve("playwright");
    const cli = resolve(dirname(playwrightEntry), "cli.js");
    const releaseProfile = acquireProfileLock(profile);
    try {
      const result = spawnSync(process.execPath, [cli, "codegen", `--user-data-dir=${profile}`, `--output=${output}`, "--viewport-size=1440,900", url], { stdio: "inherit" });
      if (result.error) throw new Error("Playwright Codegen could not start. Check that Playwright is installed in this project.");
      if (result.status !== 0) process.exitCode = result.status ?? 1;
    } finally { releaseProfile(); }
    return;
  }
  if (command === "run") {
    const recipe = await readRecipe(recipeFile);
    const base = validateBaseUrl(options["base-url"]);
    const profile = await requireProfileDirectory(options["profile-dir"]);
    const artifactRoot = resolve(options.artifacts ?? "test-results/spfx-fast-playwright");
    const timeoutMs = Number(options["timeout-ms"] ?? 15000);
    if (!Number.isInteger(timeoutMs) || timeoutMs < 1) throw new Error("--timeout-ms must be a positive integer.");
    await mkdir(artifactRoot, { recursive: true });
    let artifacts = artifactRoot;
    try {
      if ((await readdir(artifactRoot)).length > 0) {
        artifacts = resolve(artifactRoot, `run-${new Date().toISOString().replace(/[:.]/g, "-")}-${process.pid}`);
      }
    } catch {}
    await mkdir(artifacts, { recursive: true });
    const { chromium } = await import("playwright");
    const releaseProfile = acquireProfileLock(profile);
    let context;
    let page;
    let result;
    let error;
    try {
      context = await chromium.launchPersistentContext(profile, {
        headless: options.headed ? false : true,
        viewport: { width: 1440, height: 900 },
      });
      page = context.pages()[0] ?? await context.newPage();
      if (!options.fast) await context.tracing.start({ screenshots: true, snapshots: true, sources: false });
      result = await runRecipe(page, recipe, { baseUrl: base.href, screenshotDir: artifacts, timeoutMs, fast: options.fast });
    } catch (cause) {
      error = cause;
      if (page) {
        try { await page.screenshot({ path: resolve(artifacts, "failure.png"), fullPage: true }); } catch {}
      }
    } finally {
      if (context && !options.fast) try { await context.tracing.stop({ path: resolve(artifacts, "trace.zip") }); } catch {}
      await context?.close();
      releaseProfile();
    }
    const report = {
      name: recipe.name,
      status: error ? "failed" : "passed",
      stepCount: recipe.steps.length,
      screenshots: error ? ["failure.png"] : result.screenshots.map((file) => file.slice(artifacts.length + 1)),
      trace: options.fast ? null : "trace.zip",
      fastMode: options.fast,
      failedStep: error?.step ?? null,
      error: error ? error.message : null,
    };
    await writeFile(resolve(artifacts, "summary.json"), `${JSON.stringify(report, null, 2)}\n`);
    console.log(JSON.stringify({ ...report, artifactDirectory: artifacts }, null, 2));
    if (error) process.exitCode = 1;
    return;
  }
  throw new Error(`Unknown command: ${command}. Use --help.`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : "Command failed.");
  process.exitCode = 1;
});
