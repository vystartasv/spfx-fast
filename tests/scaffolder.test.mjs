import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

const cli = fileURLToPath(new URL("../bin/spfx-fast.mjs", import.meta.url));

test("CLI scaffolds a Fluent UI v9 list app and refuses overwrites", () => {
  const directory = mkdtempSync(join(tmpdir(), "spfx-fast-scaffold-"));
  try {
    const args = [
      cli,
      "create",
      "Tasks",
      "--site-id",
      "contoso.sharepoint.com,siteGuid,webGuid",
      "--list-id",
      "listGuid",
      "--title",
      "Team tasks",
      "--field",
      "Title:Title:text",
      "--field",
      "Status:Status:choice:Active,Done",
    ];
    const first = spawnSync(process.execPath, args, { cwd: directory, encoding: "utf8" });
    assert.equal(first.status, 0, first.stderr);
    const component = readFileSync(join(directory, "src/spfxFast/Tasks.tsx"), "utf8");
    assert.match(component, /@fluentui\/react-components/);
    assert.match(component, /FluentProvider/);
    assert.match(component, /"Active"/);
    const services = readFileSync(join(directory, "src/spfxFast/createServices.ts"), "utf8");
    assert.match(services, /createSPFxServices\(context, SPHttpClient\.configurations\.v1\)/);
    assert.match(services, /WebPartContext/);
    const host = readFileSync(join(directory, "src/spfxFast/TasksHost.tsx"), "utf8");
    assert.match(host, /createAppServices\(context\)/);
    assert.match(host, /role="status"/);
    assert.match(host, /role="alert"/);
    assert.match(host, /<Tasks client=\{services\.client\}/);
    assert.match(first.stdout, /<TasksHost context=\{this\.context\} \/>/);

    const second = spawnSync(process.execPath, args, { cwd: directory, encoding: "utf8" });
    assert.notEqual(second.status, 0);
    assert.match(second.stderr, /Refusing to overwrite/);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});

test("agent init writes project guidance without overwriting existing instruction files", () => {
  const directory = mkdtempSync(join(tmpdir(), "spfx-fast-agent-"));
  try {
    const first = spawnSync(process.execPath, [cli, "agent", "init"], { cwd: directory, encoding: "utf8" });
    assert.equal(first.status, 0, first.stderr);
    assert.match(readFileSync(join(directory, "SPFX-FAST.md"), "utf8"), /capabilities\.json/);
    assert.match(readFileSync(join(directory, "AGENTS.md"), "utf8"), /read \[SPFX-FAST\.md\]\(SPFX-FAST\.md\) before making changes/);
    assert.match(first.stdout, /added its discovery link to .*AGENTS\.md/);
    const second = spawnSync(process.execPath, [cli, "agent", "init"], { cwd: directory, encoding: "utf8" });
    assert.notEqual(second.status, 0);
    assert.match(second.stderr, /Refusing to overwrite/);
    assert.equal((readFileSync(join(directory, "AGENTS.md"), "utf8").match(/spfx-fast:begin/g) ?? []).length, 1);
  } finally { rmSync(directory, { recursive: true, force: true }); }
});

test("agent init preserves existing AGENTS.md content and adds one scoped guide link", () => {
  const directory = mkdtempSync(join(tmpdir(), "spfx-fast-agent-existing-"));
  const existing = "# Existing repository instructions\n\nKeep these rules intact.\n";
  try {
    writeFileSync(join(directory, "AGENTS.md"), existing);
    const result = spawnSync(process.execPath, [cli, "agent", "init"], { cwd: directory, encoding: "utf8" });
    assert.equal(result.status, 0, result.stderr);
    const instructions = readFileSync(join(directory, "AGENTS.md"), "utf8");
    assert.ok(instructions.startsWith(existing));
    assert.equal((instructions.match(/spfx-fast:begin/g) ?? []).length, 1);
    assert.match(instructions, /read \[SPFX-FAST\.md\]\(SPFX-FAST\.md\) before making changes/);
  } finally { rmSync(directory, { recursive: true, force: true }); }
});

test("agent doctor reports local dependencies and does not claim tenant consent", () => {
  const directory = mkdtempSync(join(tmpdir(), "spfx-fast-doctor-"));
  try {
    writeFileSync(join(directory, "package.json"), JSON.stringify({ dependencies: { "@microsoft/sp-core-library": "~1.24.0", "@spfx-fast/core": "^0.1.0" } }));
    mkdirSync(join(directory, "src"));
    writeFileSync(join(directory, "src", "webpart.ts"), 'import { createClient } from "@spfx-fast/core";');
    writeFileSync(join(directory, "SPFX-FAST.md"), "agent notes");
    const result = spawnSync(process.execPath, [cli, "agent", "doctor"], { cwd: directory, encoding: "utf8" });
    assert.equal(result.status, 0, result.stderr);
    assert.match(result.stdout, /PASS SPFx 1\.24\+/);
    assert.match(result.stdout, /INFO Tenant consent cannot be verified/);
  } finally { rmSync(directory, { recursive: true, force: true }); }
});

test("agent doctor locates a single SPFx project from a Rush workspace root", () => {
  const directory = mkdtempSync(join(tmpdir(), "spfx-fast-rush-doctor-"));
  const project = join(directory, "apps", "webpart");
  try {
    mkdirSync(project, { recursive: true });
    writeFileSync(join(directory, "rush.json"), JSON.stringify({ projects: [{ projectFolder: "apps/webpart" }] }));
    writeFileSync(join(project, "package.json"), JSON.stringify({ dependencies: { "@microsoft/sp-core-library": "1.24.0", "@spfx-fast/core": "^0.1.0" } }));
    writeFileSync(join(project, "SPFX-FAST.md"), "agent notes");
    mkdirSync(join(project, "src"));
    writeFileSync(join(project, "src", "webpart.ts"), 'import { createClient } from "@spfx-fast/core";');
    const result = spawnSync(process.execPath, [cli, "agent", "doctor"], { cwd: directory, encoding: "utf8" });
    assert.equal(result.status, 0, result.stderr);
    assert.match(result.stdout, /INFO Workspace layout: Rush/);
    assert.match(result.stdout, /INFO SPFx project: .*apps[\\/]webpart/);
    assert.match(result.stdout, /PASS SPFx 1\.24\+/);
  } finally { rmSync(directory, { recursive: true, force: true }); }
});

test("dashboard command creates a Fluent UI v9 multi-list component", () => {
  const directory = mkdtempSync(join(tmpdir(), "spfx-fast-dashboard-"));
  try {
    const result = spawnSync(process.execPath, [cli, "dashboard", "Overview", "--site-id", "tenant,site,web", "--list", "Tasks:list-a:Title,Status", "--list", "Projects:list-b:Title"], { cwd: directory, encoding: "utf8" });
    assert.equal(result.status, 0, result.stderr);
    const source = readFileSync(join(directory, "src/spfxFast/Overview.tsx"), "utf8");
    assert.match(source, /@fluentui\/react-components/);
    assert.match(source, /useItems\(client, siteId, source\.listId/);
    assert.match(source, /ListView/);
    assert.match(source, /Loading/);
    assert.match(source, /ItemForm/);
    assert.match(source, /setPendingDelete\(selected\)/);
    assert.match(source, /<Dialog/);
  } finally { rmSync(directory, { recursive: true, force: true }); }
});

test("dashboard command accepts inspected field schema and emits typed forms", () => {
  const directory = mkdtempSync(join(tmpdir(), "spfx-fast-dashboard-schema-"));
  try {
    writeFileSync(join(directory, "schema.json"), JSON.stringify({
      siteId: "tenant,site,web",
      lists: [{ title: "Requests", listId: "requests-id", filter: { field: "Status", operator: "eq", value: "Open" }, pageSize: 25, maxItems: 250, fields: [
        { name: "Title", title: "Request title", type: "text", required: true },
        { name: "Priority", title: "Priority", type: "choice", choices: ["High", "Normal"] },
        { name: "Owner", title: "Owner", type: "person", options: [{ value: 13, label: "Alex Example" }] },
        { name: "Project", title: "Project", type: "lookup", options: [{ value: "42", label: "Rollout" }] },
      ] }],
    }));
    const result = spawnSync(process.execPath, [cli, "dashboard", "RequestsDashboard", "--schema-file", "schema.json"], { cwd: directory, encoding: "utf8" });
    assert.equal(result.status, 0, result.stderr);
    const source = readFileSync(join(directory, "src/spfxFast/RequestsDashboard.tsx"), "utf8");
    assert.match(source, /"type": "choice"/);
    assert.match(source, /"High"/);
    assert.match(source, /"type": "person"/);
    assert.match(source, /"type": "lookup"/);
    assert.match(source, /"required": true/);
    assert.match(source, /updateItem/);
    assert.match(source, /useListItems/);
    assert.match(source, /"operator": "eq"/);
    assert.match(source, /Load more/);
  } finally { rmSync(directory, { recursive: true, force: true }); }
});

test("scenario catalog generates an inspected-list dashboard preset", () => {
  const directory = mkdtempSync(join(tmpdir(), "spfx-fast-dashboard-scenario-"));
  try {
    const list = spawnSync(process.execPath, [cli, "scenarios"], { cwd: directory, encoding: "utf8" });
    assert.equal(list.status, 0, list.stderr);
    assert.match(list.stdout, /issue-reporting\s+Issue reporting/);
    const result = spawnSync(process.execPath, [cli, "dashboard", "IssueCenter", "--site-id", "tenant,site,web", "--scenario", "issue-reporting", "--list-id", "issues-id"], { cwd: directory, encoding: "utf8" });
    assert.equal(result.status, 0, result.stderr);
    assert.match(result.stdout, /from 'scenario issue-reporting'/);
    const source = readFileSync(join(directory, "src/spfxFast/IssueCenter.tsx"), "utf8");
    assert.match(source, /"name": "Severity"/);
    assert.match(source, /"Critical"/);
    const unknown = spawnSync(process.execPath, [cli, "dashboard", "IssueCenter", "--site-id", "tenant,site,web", "--scenario", "not-a-scenario", "--list-id", "issues-id"], { cwd: directory, encoding: "utf8" });
    assert.notEqual(unknown.status, 0);
    assert.match(unknown.stderr, /Unknown scenario/);
  } finally { rmSync(directory, { recursive: true, force: true }); }
});

test("use-case catalog generates from a business-specific ID and supports category search", () => {
  const directory = mkdtempSync(join(tmpdir(), "spfx-fast-use-case-"));
  try {
    const list = spawnSync(process.execPath, [cli, "use-cases", "requests"], { cwd: directory, encoding: "utf8" });
    assert.equal(list.status, 0, list.stderr);
    assert.match(list.stdout, /laptop-request/);
    assert.doesNotMatch(list.stdout, /employee-onboarding/);
    const result = spawnSync(process.execPath, [cli, "dashboard", "LaptopRequests", "--site-id", "tenant,site,web", "--use-case", "laptop-request", "--list-id", "requests-id"], { cwd: directory, encoding: "utf8" });
    assert.equal(result.status, 0, result.stderr);
    assert.match(result.stdout, /from 'use case laptop-request'/);
    const source = readFileSync(join(directory, "src/spfxFast/LaptopRequests.tsx"), "utf8");
    assert.match(source, /"name": "DeviceType"/);
    const unknownCategory = spawnSync(process.execPath, [cli, "use-cases", "not-a-category"], { cwd: directory, encoding: "utf8" });
    assert.notEqual(unknownCategory.status, 0);
    assert.match(unknownCategory.stderr, /Unknown use-case category/);
  } finally { rmSync(directory, { recursive: true, force: true }); }
});
