import assert from "node:assert/strict";
import { execFileSync, spawnSync } from "node:child_process";
import { copyFileSync, cpSync, existsSync, mkdtempSync, mkdirSync, readFileSync, readdirSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const scratch = mkdtempSync(join(tmpdir(), "spfx-fast-package-smoke-"));

function run(command, args, cwd) {
  const result = spawnSync(command, args, { cwd, encoding: "utf8" });
  if (result.status !== 0) throw new Error(`${command} ${args.join(" ")} failed:\n${result.stderr || result.stdout}`);
  return result.stdout;
}

try {
  const packed = JSON.parse(run("npm", ["pack", "--json", "--pack-destination", scratch], root))[0];
  const archive = join(scratch, packed.filename);
  if (process.env.SPFX_FAST_PACKAGE_ARCHIVE) {
    const destination = resolve(process.env.SPFX_FAST_PACKAGE_ARCHIVE);
    mkdirSync(dirname(destination), { recursive: true });
    copyFileSync(archive, destination);
    console.log(`Validated package archive saved to ${destination}`);
  }
  const extracted = join(scratch, "unpacked");
  mkdirSync(extracted);
  execFileSync("tar", ["-xzf", archive, "-C", extracted]);
  const packageRoot = join(extracted, "package");
  const manifest = JSON.parse(readFileSync(join(packageRoot, "package.json"), "utf8"));
  assert.match(readFileSync(join(packageRoot, "agent/SKILL.md"), "utf8"), /source checkout.*not included in the npm package/i);
  assert.match(readFileSync(join(packageRoot, "docs/AGENT-BROWSER-WORKFLOW.md"), "utf8"), /source checkout/i);
  const discoverableSkill = readFileSync(join(packageRoot, "skills/spfx-fast/SKILL.md"), "utf8");
  assert.match(discoverableSkill, /^---\nname: spfx-fast\ndescription: '.+'\n---/);
  assert.match(discoverableSkill, /createSPFxServices\(context, SPHttpClient\.configurations\.v1\)/);
  assert.match(discoverableSkill, /pnpm dlx --package @spfx-fast\/core spfx-fast dashboard/);
  assert.match(discoverableSkill, /npx --yes --package=@spfx-fast\/core -- spfx-fast/);
  assert.match(readFileSync(join(packageRoot, "README.md"), "utf8"), /npx skills add https:\/\/github\.com\/vystartasv\/spfx-fast\/tree\/main\/skills\/spfx-fast/);

  for (const relative of [
    "README.md",
    "agent/SKILL.md",
    "agent/capabilities.json",
    "skills/spfx-fast/SKILL.md",
    "docs/EXAMPLES.md",
    "docs/SPFX-EXAMPLES.md",
    "docs/API-COVERAGE.md",
    "docs/assets/screenshots/drive-delta-preview.png",
    "docs/assets/screenshots/request-attachments-preview.png",
    "docs/assets/screenshots/request-attachments.png",
    "docs/assets/screenshots/dashboard-calendar.png",
    "docs/assets/screenshots/floor-plan-map.png",
    "docs/assets/screenshots/tenant-upload-final-response-0.2.4.png",
    "docs/assets/screenshots/tenant-upload-replace-final-response-0.2.4.png",
    "docs/assets/screenshots/tenant-upload-rename-conflict.png",
    "docs/assets/screenshots/tenant-upload-rename-recovery-0.2.4.png",
    "docs/assets/screenshots/tenant-delta-3000/list-delta-3000.png",
    "docs/assets/screenshots/tenant-drive-delta-3000/drive-delta-3000.png",
    "examples/README.md",
    "examples/recipes.ts",
    "examples/FilteredTasks.tsx",
    "examples/DriveItemPermissions.ts",
    "examples/DriveMutations.ts",
    "examples/DriveCopy.ts",
    "examples/DeviceCapture.tsx",
    "examples/CalendarAgenda.tsx",
    "examples/FloorPlanMap.tsx",
    "examples/scenarios.json",
    "examples/use-cases.json",
    "docs/SCENARIOS.md",
    "docs/USE-CASES.md",
    "docs/ROADMAP.md",
    "docs/TENANT-VALIDATION.md",
    "docs/RELEASE.md",
    "examples/ListContentTypes.ts",
    "examples/UserProfiles.ts",
    "examples/Versions.ts",
    "examples/TaxonomyPickerData.ts",
    "examples/ListPermissions.ts",
    "examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/components/DriveDeltaPanel.tsx",
    "examples/SharePointRestPlayground.tsx",
    "examples/ItemAttachmentsPanel.tsx",
    "examples/fixtures/spfx-fast-dashboard/README.md",
    "examples/fixtures/spfx-fast-dashboard/package-lock.json",
    "examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/SpfxFastDashboardWebPart.tsx",
    "examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/components/DocumentLibraryPanel.tsx",
    "examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/components/CalendarAgenda.tsx",
    "examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/components/PerformanceProbe.tsx",
    "examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/components/WorkDashboard.tsx",
    "examples/fixtures/spfx-fast-dashboard/src/webparts/spfxFastDashboard/components/TenantValidationPanel.tsx",
    "examples/fixtures/spfx-fast-copilot/README.md",
    "examples/fixtures/spfx-fast-copilot/package-lock.json",
    "examples/fixtures/spfx-fast-copilot/src/copilotComponents/workItems/WorkItemsCopilotComponent.tsx",
    "examples/fixtures/spfx-fast-copilot/src/copilotComponents/workItems/components/WorkItems.tsx",
    "examples/fixtures/spfx-fast-copilot/copilot/instruction.txt",
    "examples/fixtures/spfx-fast-copilot/scripts/clean-generated.mjs",
    "templates/dashboard.tsx.template",
    "dist/index.d.ts",
    "dist/spfx.d.ts",
    "dist/react.d.ts",
    "dist/markdown.js",
    "dist/markdown.d.ts",
  ]) assert.doesNotThrow(() => readFileSync(join(packageRoot, relative)), `packed package missing ${relative}`);
  assert.equal(existsSync(join(packageRoot, "examples/fixtures/spfx-fast-dashboard/dist")), false, "generated fixture bundles are excluded from the package");
  assert.equal(existsSync(join(packageRoot, "examples/fixtures/spfx-fast-copilot/dist")), false, "generated Copilot fixture bundles are excluded from the package");
  assert.equal(existsSync(join(packageRoot, "examples/fixtures/spfx-fast-dashboard/node_modules")), false, "dashboard fixture dependencies are excluded from the package");
  assert.equal(existsSync(join(packageRoot, "examples/fixtures/spfx-fast-copilot/node_modules")), false, "Copilot fixture dependencies are excluded from the package");
  assert.equal(existsSync(join(packageRoot, "examples/fixtures/spfx-fast-copilot/teams")), false, "generated Copilot Teams packages are excluded from the package");

  const consumer = join(scratch, "consumer");
  mkdirSync(consumer);
  writeFileSync(join(consumer, "package.json"), JSON.stringify({ private: true, type: "module" }));
  run("npm", ["install", "--ignore-scripts", "--no-audit", "--no-fund", "--no-package-lock", "--no-save", archive], consumer);
  const nodeModules = join(consumer, "node_modules");
  for (const dependency of ["react", "@fluentui/react-components"]) {
    const target = join(root, "node_modules", dependency);
    const link = join(nodeModules, dependency);
    mkdirSync(join(link, ".."), { recursive: true });
    rmSync(link, { recursive: true, force: true });
    symlinkSync(target, link, "dir");
  }

  const entrypoints = ["@spfx-fast/core", ...Object.keys(manifest.exports).filter((name) => name.startsWith("./") && !name.endsWith(".json")).map((name) => `@spfx-fast/core/${name.slice(2)}`)];
  const importScript = entrypoints.map((name) => `await import(${JSON.stringify(name)});`).join("\n");
  run(process.execPath, ["--input-type=module", "-e", `${importScript}\nawait import("@spfx-fast/core/capabilities.json", { with: { type: "json" } });`], consumer);

  const capabilities = JSON.parse(run(process.execPath, [join(packageRoot, "bin/spfx-fast.mjs"), "catalog"], scratch));
  assert.equal(capabilities.name, "SPFx Fast");
  const packagedRecipes = readFileSync(join(packageRoot, "docs/EXAMPLES.md"), "utf8");
  assert.equal(capabilities.exampleCatalog.recipeCount, (packagedRecipes.match(/^\|\s*\d+\s*\|/gm) ?? []).length);
  assert.equal(capabilities.exampleCatalog.scenarioCount, 24);
  assert.equal(capabilities.exampleCatalog.useCaseCount, 100);
  const capabilityNames = capabilities.capabilities.map(({ name }) => name).sort();
  assert.deepEqual(Object.keys(capabilities.capabilityDocs ?? {}).sort(), capabilityNames, "packed capability catalog must link every capability");
  for (const [name, target] of Object.entries(capabilities.capabilityDocs)) {
    const [relative, anchor] = target.split("#");
    const page = readFileSync(join(packageRoot, relative), "utf8");
    const slugs = [...page.matchAll(/^#{1,6} (.+)$/gm)].map(([, heading]) => heading.toLowerCase().replace(/[^a-z0-9 -]/g, "").trim().replace(/\s+/g, "-"));
    assert.ok(slugs.includes(anchor), `packed ${name} documentation target is missing: ${target}`);
  }
  const installedRoot = join(nodeModules, "@spfx-fast/core");
  run(process.execPath, [join(installedRoot, "bin/spfx-fast.mjs"), "agent", "init"], consumer);
  const guide = readFileSync(join(consumer, "SPFX-FAST.md"), "utf8");
  const instructions = readFileSync(join(consumer, "AGENTS.md"), "utf8");
  for (const relative of ["agent/SKILL.md", "agent/capabilities.json", "docs/EXAMPLES.md", "docs/SPFX-EXAMPLES.md", "docs/API-COVERAGE.md"]) {
    assert.ok(readFileSync(join(packageRoot, relative)), `agent link target missing: ${relative}`);
  }
  assert.match(guide, /node_modules\/@spfx-fast\/core\/agent\/SKILL\.md/);
  assert.match(instructions, /read \[SPFX-FAST\.md\]\(SPFX-FAST\.md\) before making changes/);
  const generated = run(process.execPath, [join(installedRoot, "bin/spfx-fast.mjs"), "dashboard", "Smoke", "--site-id", "tenant,site,web", "--list", "Tasks:list-id:Title"], consumer);
  assert.match(generated, /Created/);
  assert.match(readFileSync(join(consumer, "src/spfxFast/Smoke.tsx"), "utf8"), /FluentProvider/);

  // Compile the shipped Copilot starter against this exact tarball, not the
  // repository's source directory or workspace-linked package.
  const fixture = join(scratch, "packed-copilot-fixture");
  const fixtureSource = join(packageRoot, "examples/fixtures/spfx-fast-copilot");
  cpSync(fixtureSource, fixture, {
    recursive: true,
    filter: (path) => !/(^|\/)(node_modules|dist|lib|lib-dts|sharepoint|solution|temp|teams|\.heft)(\/|$)/.test(path),
  });
  const fixtureManifestPath = join(fixture, "package.json");
  const fixtureManifest = JSON.parse(readFileSync(fixtureManifestPath, "utf8"));
  fixtureManifest.dependencies["@spfx-fast/core"] = `file:${archive}`;
  writeFileSync(fixtureManifestPath, `${JSON.stringify(fixtureManifest, null, 2)}\n`);
  rmSync(join(fixture, "package-lock.json"), { force: true });
  run("npm", ["install", "--no-audit", "--no-fund"], fixture);
  run("npm", ["run", "build"], fixture);
  assert.ok(existsSync(join(fixture, "sharepoint/solution/spfx-fast-copilot.sppkg")), "packed package builds into the Copilot SPFx solution");
  assert.deepEqual(
    readdirSync(join(fixture, "teams")).filter((name) => name.endsWith(".zip")),
    ["sharepoint-list-browser.zip"],
    "packed Copilot starter contains only its current agent package",
  );

  // Generate both one-list and multi-list UI from the installed CLI in the
  // published dashboard fixture, then compile those generated files in SPFx.
  const dashboardFixture = join(scratch, "packed-dashboard-fixture");
  const dashboardSource = join(packageRoot, "examples/fixtures/spfx-fast-dashboard");
  cpSync(dashboardSource, dashboardFixture, {
    recursive: true,
    filter: (path) => !/(^|\/)(node_modules|dist|lib|lib-dts|sharepoint|solution|temp|teams|\.heft)(\/|$)/.test(path),
  });
  const dashboardManifestPath = join(dashboardFixture, "package.json");
  const dashboardManifest = JSON.parse(readFileSync(dashboardManifestPath, "utf8"));
  dashboardManifest.dependencies["@spfx-fast/core"] = `file:${archive}`;
  writeFileSync(dashboardManifestPath, `${JSON.stringify(dashboardManifest, null, 2)}\n`);
  rmSync(join(dashboardFixture, "package-lock.json"), { force: true });
  run("npm", ["install", "--no-audit", "--no-fund"], dashboardFixture);
  const dashboardCli = join(dashboardFixture, "node_modules/@spfx-fast/core/bin/spfx-fast.mjs");
  run(process.execPath, [dashboardCli, "agent", "init"], dashboardFixture);
  const doctor = run(process.execPath, [dashboardCli, "agent", "doctor"], dashboardFixture);
  assert.match(doctor, /PASS SPFx 1\.24\+/);
  assert.match(doctor, /PASS Agent guide/);
  assert.match(doctor, /INFO Tenant consent cannot be verified/);
  assert.match(run(process.execPath, [dashboardCli, "scenarios"], dashboardFixture), /issue-reporting/);
  const useCases = run(process.execPath, [dashboardCli, "use-cases", "requests"], dashboardFixture);
  assert.match(useCases, /laptop-request/);
  assert.match(useCases, /pnpm dlx --package @spfx-fast\/core spfx-fast dashboard/);
  assert.match(useCases, /npx --yes --package=@spfx-fast\/core -- spfx-fast dashboard/);
  const createOutput = run(process.execPath, [dashboardCli, "create", "Tasks", "--site-id", "contoso.sharepoint.com,site,web", "--list-id", "tasks-id", "--title", "Tasks"], dashboardFixture);
  assert.match(createOutput, /TasksHost\.tsx/);
  const generatedServices = readFileSync(join(dashboardFixture, "src/spfxFast/createServices.ts"), "utf8");
  assert.match(generatedServices, /createSPFxServices\(context, SPHttpClient\.configurations\.v1\)/);
  assert.match(generatedServices, /WebPartContext/);
  const generatedHost = readFileSync(join(dashboardFixture, "src/spfxFast/TasksHost.tsx"), "utf8");
  assert.match(generatedHost, /createAppServices\(context\)/);
  assert.match(generatedHost, /role="status"/);
  run(process.execPath, [dashboardCli, "dashboard", "Overview", "--site-id", "contoso.sharepoint.com,site,web", "--list", "Tasks:tasks-id:Title,Status", "--list", "Projects:projects-id:Title"], dashboardFixture);
  writeFileSync(join(dashboardFixture, "inspected-schema.json"), JSON.stringify({
    siteId: "contoso.sharepoint.com,site,web",
    lists: [{
      title: "Requests",
      listId: "requests-id",
      filter: { field: "Status", operator: "eq", value: "Open" },
      pageSize: 25,
      maxItems: 250,
      fields: [
        { name: "Title", title: "Request", type: "text", required: true },
        { name: "Priority", title: "Priority", type: "choice", choices: ["High", "Normal"] },
      ],
    }],
  }));
  run(process.execPath, [dashboardCli, "dashboard", "RequestsDashboard", "--schema-file", "inspected-schema.json"], dashboardFixture);
  run(process.execPath, [dashboardCli, "dashboard", "IssueCenter", "--site-id", "contoso.sharepoint.com,site,web", "--scenario", "issue-reporting", "--list-id", "issues-id"], dashboardFixture);
  run(process.execPath, [dashboardCli, "dashboard", "LaptopRequests", "--site-id", "contoso.sharepoint.com,site,web", "--use-case", "laptop-request", "--list-id", "requests-id"], dashboardFixture);
  assert.match(readFileSync(join(dashboardFixture, "src/spfxFast/Tasks.tsx"), "utf8"), /useItems/);
  assert.match(readFileSync(join(dashboardFixture, "src/spfxFast/Overview.tsx"), "utf8"), /useListItems/);
  const requestsDashboard = readFileSync(join(dashboardFixture, "src/spfxFast/RequestsDashboard.tsx"), "utf8");
  assert.match(requestsDashboard, /"type": "choice"/);
  assert.match(requestsDashboard, /"required": true/);
  assert.match(requestsDashboard, /"operator": "eq"/);
  run("npm", ["run", "build"], dashboardFixture);
  assert.ok(existsSync(join(dashboardFixture, "sharepoint/solution/spfx-fast-dashboard.sppkg")), "one-list and multi-list generated apps compile into the SPFx solution");

  console.log(`Packed artifact smoke passed: ${entrypoints.length} entrypoints, agent docs, 24 form presets, 100 use cases, recipes, and compiled generated SPFx starters.`);
} finally {
  rmSync(scratch, { recursive: true, force: true });
}
