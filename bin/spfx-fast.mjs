#!/usr/bin/env node
import { access, mkdir, readFile, readdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const args = process.argv.slice(2);
const command = args.shift();

function usage() {
  console.log(`SPFx Fast\n\nCommands:\n  spfx-fast create <Name> --site-id <GraphSiteId> --list-id <ListId> [--title <Title>] [--field <InternalName[:Title[:type[:choice1,choice2]]]>]...\n  spfx-fast dashboard <Name> --site-id <GraphSiteId> (--list <Title:ListId[:Field1,Field2]>... | --schema-file <inspected-schema.json> | --scenario <preset-id> --list-id <ListId> | --use-case <use-case-id> --list-id <ListId>) [--title <Title>]\n  spfx-fast scenarios\n  spfx-fast use-cases [category-id]\n  spfx-fast catalog`);
  console.log("  spfx-fast agent init\n  spfx-fast agent doctor");
}

function parseOptions(values) {
  const options = {};
  for (let i = 0; i < values.length; i++) {
    const key = values[i];
    if (!key.startsWith("--")) throw new Error(`Unexpected argument: ${key}`);
    if (key === "--help") { options.help = true; continue; }
    const value = values[++i];
    if (!value || value.startsWith("--")) throw new Error(`Expected a value after ${key}`);
    const name = key.slice(2);
    if (name === "field" || name === "list") (options[name] ??= []).push(value);
    else options[name] = value;
  }
  return options;
}

async function createDashboard(nameArg, rawOptions) {
  const options = parseOptions(rawOptions);
  if (options.help) return usage();
  const name = pascalCase(nameArg ?? "");
  let schema;
  let scenario;
  if (options["schema-file"]) {
    schema = JSON.parse(await readFile(resolve(options["schema-file"]), "utf8"));
    if (!Array.isArray(schema.lists)) throw new Error("Dashboard schema file must contain a lists array.");
  }
  if (options.scenario) {
    if (options["use-case"]) throw new Error("Choose either --scenario or --use-case.");
  }
  if (options.scenario || options["use-case"]) {
    if (schema || options.list?.length) throw new Error("Use --scenario or --use-case by itself with --list-id, not with --list or --schema-file.");
    if (!options["list-id"]) throw new Error("--list-id is required with --scenario or --use-case.");
    const catalog = JSON.parse(await readFile(resolve(root, "examples/scenarios.json"), "utf8"));
    let presetId = options.scenario;
    if (options["use-case"]) {
      const useCases = JSON.parse(await readFile(resolve(root, "examples/use-cases.json"), "utf8"));
      const useCase = useCases.categories.flatMap((category) => category.useCases).find((candidate) => candidate.id === options["use-case"]);
      if (!useCase) throw new Error(`Unknown use case '${options["use-case"]}'. Run 'pnpm dlx --package @spfx-fast/core spfx-fast use-cases' or 'npx --yes --package=@spfx-fast/core -- spfx-fast use-cases' to list available use cases.`);
      presetId = useCase.preset;
      scenario = { id: useCase.id, name: options["list-title"] ?? useCase.name, fields: catalog.scenarios.find((candidate) => candidate.id === presetId)?.fields };
      if (!scenario.fields) throw new Error(`Use case '${useCase.id}' refers to missing preset '${presetId}'.`);
    } else {
      scenario = catalog.scenarios.find((candidate) => candidate.id === presetId);
      if (!scenario) throw new Error(`Unknown scenario '${presetId}'. Run 'pnpm dlx --package @spfx-fast/core spfx-fast scenarios' or 'npx --yes --package=@spfx-fast/core -- spfx-fast scenarios' to list available scenarios.`);
    }
    schema = { lists: [{ title: options["list-title"] ?? scenario.name, listId: options["list-id"], fields: scenario.fields }] };
  }
  const siteId = options["site-id"] ?? schema?.siteId;
  if (!siteId) throw new Error("--site-id is required unless the schema file includes siteId.");
  if (options.list?.length && schema) throw new Error("Use either --list entries or --schema-file/--scenario, not both.");
  if (!options.list?.length && !schema) throw new Error("Provide --list Title:ListId[:Field1,Field2], --schema-file, or --scenario <id> --list-id <ListId>.");
  const sources = schema
    ? schema.lists.map((entry) => {
      if (!entry.title || !entry.listId || !Array.isArray(entry.fields) || !entry.fields.length) throw new Error("Every schema list needs title, listId, and at least one inspected field.");
      const fields = entry.fields.map((field) => {
        if (!field.name) throw new Error(`A field in '${entry.title}' is missing its internal name.`);
        const type = field.type ?? "text";
        if (!["text", "multiline", "number", "date", "boolean", "choice", "multiChoice", "person", "lookup"].includes(type)) throw new Error(`Unsupported field type '${type}' in '${entry.title}.${field.name}'.`);
        if (["choice", "multiChoice"].includes(type) && field.choices !== undefined && !Array.isArray(field.choices)) throw new Error(`Choice values for '${entry.title}.${field.name}' must be an array.`);
        if (["person", "lookup"].includes(type) && field.options !== undefined && (!Array.isArray(field.options) || field.options.some((option) => !["string", "number"].includes(typeof option.value) || typeof option.label !== "string"))) throw new Error(`Options for '${entry.title}.${field.name}' must have string or numeric values and string labels.`);
        return { name: field.name, title: field.title ?? field.name, type, required: Boolean(field.required), ...(["choice", "multiChoice"].includes(type) ? { choices: field.choices ?? [] } : {}), ...(["person", "lookup"].includes(type) && field.options ? { options: field.options } : {}) };
      });
      let filter;
      if (entry.filter !== undefined) {
        const candidate = entry.filter;
        const operators = ["eq", "ne", "lt", "le", "gt", "ge", "startswith"];
        if (!candidate || typeof candidate.field !== "string" || !/^[A-Za-z0-9_]{1,128}$/.test(candidate.field)
          || !operators.includes(candidate.operator)
          || !["string", "number", "boolean"].includes(typeof candidate.value)
          || (typeof candidate.value === "number" && !Number.isFinite(candidate.value))
          || (candidate.operator === "startswith" && typeof candidate.value !== "string")) {
          throw new Error(`Invalid Graph filter in '${entry.title}'. Use a safe field, supported operator, and string/number/boolean value.`);
        }
        filter = { field: candidate.field, operator: candidate.operator, value: candidate.value };
      }
      const pageSize = entry.pageSize === undefined ? 50 : Number(entry.pageSize);
      const maxItems = entry.maxItems === undefined ? 500 : Number(entry.maxItems);
      if (!Number.isInteger(pageSize) || pageSize < 1 || pageSize > 100) throw new Error(`pageSize in '${entry.title}' must be an integer from 1 to 100.`);
      if (!Number.isInteger(maxItems) || maxItems < 1 || maxItems > 10000) throw new Error(`maxItems in '${entry.title}' must be an integer from 1 to 10000.`);
      return { title: entry.title, listId: entry.listId, fields, ...(filter ? { filter, pageSize, maxItems } : {}) };
    })
    : options.list.map((value) => {
      const [title, listId, fieldNames = "Title"] = value.split(":");
      if (!title || !listId) throw new Error(`Invalid list '${value}'. Use Title:ListId[:Field1,Field2].`);
      const fields = fieldNames.split(",").filter(Boolean).map((field) => ({ name: field, title: field, type: "text" }));
      return { title, listId, fields: fields.length ? fields : [{ name: "Title", title: "Title", type: "text" }] };
    });
  if (!sources.length) throw new Error("Dashboard schema must contain at least one list.");
  if (new Set(sources.map(({ title }) => title.toLocaleLowerCase())).size !== sources.length) throw new Error("Dashboard list titles must be unique.");
  const template = await readFile(resolve(root, "templates/dashboard.tsx.template"), "utf8");
  const source = template
    .replaceAll("{{NAME}}", name)
    .replaceAll("{{TITLE}}", escapeJsxText(options.title ?? name))
    .replace("{{SITE_ID}}", JSON.stringify(siteId))
    .replace("{{SOURCES}}", JSON.stringify(sources, null, 2));
  const path = resolve("src/spfxFast", `${name}.tsx`);
  await writeNew(path, source);
  console.log(`Created ${path} with ${sources.length} authenticated list source(s)${scenario ? ` from '${options["use-case"] ? "use case" : "scenario"} ${scenario.id}'` : ""}, search, CRUD forms, delete confirmation, and per-list loading/error states.`);
  console.log('Pass the client returned by `createSPFxServices(context, SPHttpClient.configurations.v1)` as the `client` prop.');
}

async function listScenarios() {
  const catalog = JSON.parse(await readFile(resolve(root, "examples/scenarios.json"), "utf8"));
  for (const scenario of catalog.scenarios) console.log(`${scenario.id}\t${scenario.name}\t${scenario.description}`);
  console.log("Generate with pnpm: pnpm dlx --package @spfx-fast/core spfx-fast dashboard <Name> --site-id <GraphSiteId> --scenario <id> --list-id <inspected-ListId>");
  console.log("Generate with npm: npx --yes --package=@spfx-fast/core -- spfx-fast dashboard <Name> --site-id <GraphSiteId> --scenario <id> --list-id <inspected-ListId>");
}

async function listUseCases(categoryId) {
  const catalog = JSON.parse(await readFile(resolve(root, "examples/use-cases.json"), "utf8"));
  const categories = categoryId ? catalog.categories.filter((category) => category.id === categoryId) : catalog.categories;
  if (categoryId && categories.length === 0) throw new Error(`Unknown use-case category '${categoryId}'. Choose from: ${catalog.categories.map(({ id }) => id).join(", ")}.`);
  const flowNames = new Map(catalog.flowPatterns.map(({ id, name }) => [id, name]));
  for (const category of categories) {
    console.log(`\n${category.name} (${category.id})`);
    for (const item of category.useCases) console.log(`${item.id}\t${item.name}\t${item.preset}\t${flowNames.get(item.flow) ?? item.flow}${item.sensitive ? "\tSensitive data: check list permissions" : ""}`);
  }
  console.log("Generate with pnpm: pnpm dlx --package @spfx-fast/core spfx-fast dashboard <Name> --site-id <GraphSiteId> --use-case <id> --list-id <inspected-ListId>");
  console.log("Generate with npm: npx --yes --package=@spfx-fast/core -- spfx-fast dashboard <Name> --site-id <GraphSiteId> --use-case <id> --list-id <inspected-ListId>");
}

function pascalCase(value) {
  const result = value.replace(/[^a-zA-Z0-9]+(.)?/g, (_, next) => next ? next.toUpperCase() : "");
  if (!result || !/^[A-Za-z_$]/.test(result)) throw new Error("Name must contain a valid identifier.");
  return result[0].toUpperCase() + result.slice(1);
}

function escapeJsxText(value) {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll("{", "&#123;").replaceAll("}", "&#125;");
}

function parseFields(entries = []) {
  if (!entries.length) return [{ name: "Title", title: "Title", type: "text" }];
  return entries.map((entry) => {
    const [name, title = name, type = "text", choices = ""] = entry.split(":");
    if (!name || !title || !["text", "multiline", "number", "date", "boolean", "choice"].includes(type)) {
      throw new Error(`Invalid field '${entry}'. Supported types: text, multiline, number, date, boolean, choice.`);
    }
    return { name, title, type, ...(type === "choice" ? { choices: choices ? choices.split(",") : [] } : {}) };
  });
}

async function writeNew(path, content) {
  try { await access(path); throw new Error(`Refusing to overwrite ${path}`); }
  catch (error) { if (error.code !== "ENOENT") throw error; }
  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, content, "utf8");
}

async function exists(path) {
  try { await access(path); return true; } catch { return false; }
}

async function createApp(nameArg, rawOptions) {
  const options = parseOptions(rawOptions);
  if (options.help) return usage();
  if (!options["site-id"] || !options["list-id"]) throw new Error("--site-id and --list-id are required; inspect the target site first rather than guessing.");
  const name = pascalCase(nameArg ?? "");
  const title = options.title ?? name;
  const fields = parseFields(options.field);
  const template = await readFile(resolve(root, "templates/list-app.tsx.template"), "utf8");
  const clientTemplate = await readFile(resolve(root, "templates/create-client.ts.template"), "utf8");
  const source = template
    .replaceAll("{{NAME}}", name)
    .replaceAll("{{TITLE}}", escapeJsxText(title))
    .replaceAll("{{SITE_ID}}", JSON.stringify(options["site-id"]))
    .replaceAll("{{LIST_ID}}", JSON.stringify(options["list-id"]))
    .replaceAll("{{FIELDS}}", JSON.stringify(fields, null, 2));
  const config = JSON.stringify({ name, title, siteId: options["site-id"], listId: options["list-id"], fields }, null, 2) + "\n";
  const hostTemplate = await readFile(resolve(root, "templates/list-host.tsx.template"), "utf8");
  const hostSource = hostTemplate.replaceAll("{{NAME}}", name);
  const outputs = [
    [resolve("src/spfxFast", `${name}.tsx`), source],
    [resolve("src/spfxFast", `${name}Host.tsx`), hostSource],
    [resolve("src/spfxFast", "createServices.ts"), clientTemplate],
    [resolve("spfx-fast.config.json"), config],
  ];
  for (const [path] of outputs) {
    try { await access(path); throw new Error(`Refusing to overwrite ${path}`); }
    catch (error) { if (error.code !== "ENOENT") throw error; }
  }
  for (const [path, content] of outputs) await writeNew(path, content);
  console.log(`Created src/spfxFast/${name}.tsx, src/spfxFast/${name}Host.tsx, src/spfxFast/createServices.ts, and spfx-fast.config.json`);
  console.log("In your web part render, use `<" + name + "Host context={this.context} />` from `./components/" + name + "Host`; it initializes authenticated services and renders the app.");
}

async function initAgent() {
  const content = `# SPFx Fast agent guidance\n\nFor SharePoint Framework work, read the installed package instructions and capability catalog before choosing APIs:\n\n- Instructions: \`node_modules/@spfx-fast/core/agent/SKILL.md\`\n- Capabilities: \`node_modules/@spfx-fast/core/agent/capabilities.json\`\n- Recipes: \`node_modules/@spfx-fast/core/docs/SPFX-EXAMPLES.md\`\n- API coverage: \`node_modules/@spfx-fast/core/docs/API-COVERAGE.md\`\n\nUse the SPFx host's authenticated \`MSGraphClientV3\` / \`SPHttpClient\`. Do not acquire tokens independently.\n`;
  const guidePath = resolve("SPFX-FAST.md");
  try { await access(guidePath); throw new Error(`Refusing to overwrite ${guidePath}`); }
  catch (error) { if (error.code !== "ENOENT") throw error; }
  await writeNew(guidePath, content);

  const instructionsPath = resolve("AGENTS.md");
  const marker = "<!-- spfx-fast:begin -->";
  const instruction = `${marker}\nFor SharePoint Framework work, read [SPFX-FAST.md](SPFX-FAST.md) before making changes.\n<!-- spfx-fast:end -->\n`;
  let existing = "";
  try { existing = await readFile(instructionsPath, "utf8"); }
  catch (error) { if (error.code !== "ENOENT") throw error; }
  if (existing.includes(marker) || existing.includes("read [SPFX-FAST.md](SPFX-FAST.md) before making changes")) {
    console.log("Created SPFX-FAST.md; AGENTS.md already links the guide.");
  } else {
    const separator = existing && !existing.endsWith("\n") ? "\n\n" : existing ? "\n" : "";
    await writeFile(instructionsPath, `${existing}${separator}${instruction}`, "utf8");
    console.log(`Created SPFX-FAST.md and added its discovery link to ${instructionsPath}.`);
  }
}

async function doctor() {
  const startDirectory = resolve(".");
  let projectDirectory = startDirectory;
  let pkg = await readProjectManifest(projectDirectory);
  const rushRoot = await findRushRoot(startDirectory);
  if (!hasSpfxDependency(pkg) && rushRoot) {
    const rush = JSON.parse(await readFile(resolve(rushRoot, "rush.json"), "utf8"));
    const projects = Array.isArray(rush.projects) ? rush.projects : [];
    const candidates = [];
    for (const project of projects) {
      if (typeof project.projectFolder !== "string") continue;
      const directory = resolve(rushRoot, project.projectFolder);
      const manifest = await readProjectManifest(directory);
      if (hasSpfxDependency(manifest)) candidates.push({ directory, manifest });
    }
    if (candidates.length === 1) ({ directory: projectDirectory, manifest: pkg } = candidates[0]);
    else if (candidates.length > 1) {
      const paths = candidates.map(({ directory }) => directory.slice(rushRoot.length + 1)).join(", ");
      throw new Error("Found multiple SPFx projects in this Rush workspace (" + paths + "). Run agent doctor from the specific SPFx project folder.");
    }
  }
  if (!pkg) throw new Error("Run agent doctor from an SPFx project containing package.json, or from a Rush root with one SPFx project.");
  const workspaceLayout = rushRoot ? "Rush (" + rushRoot + ")" : "standalone";
  console.log("INFO Workspace layout: " + workspaceLayout);
  console.log("INFO SPFx project: " + projectDirectory);
  const dependencies = { ...pkg.dependencies, ...pkg.devDependencies };
  const spfxVersion = (dependencies["@microsoft/sp-core-library"] ?? "").replace(/^[~^v]/, "");
  const [major = 0, minor = 0] = spfxVersion.split(".").map(Number);
  const sourceFiles = [];
  async function collect(directory) {
    let entries;
    try { entries = await readdir(directory, { withFileTypes: true }); } catch { return; }
    for (const entry of entries) {
      if (entry.name === "node_modules" || entry.name === "dist" || entry.name === ".git") continue;
      const path = resolve(directory, entry.name);
      if (entry.isDirectory()) await collect(path);
      else if (/\.(tsx?|jsx?)$/.test(entry.name)) sourceFiles.push(path);
    }
  }
  await collect(resolve(projectDirectory, "src"));
  const source = (await Promise.all(sourceFiles.map((file) => readFile(file, "utf8")))).join("\n");
  const usesToolkit = /from\s+["']@spfx-fast\/core(?:\/[^"']*)?["']/.test(source);
  const usesFluentV8 = /from\s+["']@fluentui\/react(?:\/[^"']*)?["']/.test(source);
  const checks = [
    ["SPFx project", Boolean(spfxVersion), "@microsoft/sp-core-library dependency"],
    ["SPFx 1.24+", major > 1 || (major === 1 && minor >= 24), spfxVersion || "version not declared"],
    ["SPFx Fast installed", Boolean(dependencies["@spfx-fast/core"]), "@spfx-fast/core dependency"],
    ["Agent guide", await exists(resolve(projectDirectory, "SPFX-FAST.md")), "SPFX-FAST.md generated"],
    ["Toolkit imports", !usesToolkit || Boolean(dependencies["@spfx-fast/core"]), usesToolkit ? "source imports @spfx-fast/core" : "no toolkit imports found under src"],
    ["Fluent UI v8", !usesFluentV8, usesFluentV8 ? "@fluentui/react import found; toolkit templates require v9" : "no v8 package imports found under src"],
  ];
  for (const [label, ok, detail] of checks) console.log(`${ok ? "PASS" : "TODO"} ${label}: ${detail}`);
  let solution;
  try { solution = JSON.parse(await readFile(resolve(projectDirectory, "config/package-solution.json"), "utf8")); } catch { /* optional */ }
  const permissions = solution?.solution?.webApiPermissionRequests ?? [];
  console.log(`INFO Graph permissions: ${permissions.length ? permissions.map((p) => `${p.resource}/${p.scope}`).join(", ") : "none declared; verify required scopes and tenant admin approval"}`);
  console.log("INFO Tenant consent cannot be verified by this local static check.");
  if (checks.some(([, ok]) => !ok)) process.exitCode = 1;
}

async function readProjectManifest(directory) {
  try { return JSON.parse(await readFile(resolve(directory, "package.json"), "utf8")); }
  catch { return undefined; }
}

function hasSpfxDependency(pkg) {
  return Boolean(pkg && ({ ...pkg.dependencies, ...pkg.devDependencies })["@microsoft/sp-core-library"]);
}

async function findRushRoot(startDirectory) {
  let directory = startDirectory;
  while (true) {
    if (await exists(resolve(directory, "rush.json"))) return directory;
    const parent = dirname(directory);
    if (parent === directory) return undefined;
    directory = parent;
  }
}

try {
  if (command === "--help" || command === "help" || !command) usage();
  else if (command === "catalog") console.log(await readFile(resolve(root, "agent/capabilities.json"), "utf8"));
  else if (command === "scenarios") await listScenarios();
  else if (command === "use-cases") await listUseCases(args[0]);
  else if (command === "dashboard") await createDashboard(args.shift(), args);
  else if (command === "agent") {
    const action = args.shift();
    if (action === "init") await initAgent();
    else if (action === "doctor") await doctor();
    else throw new Error("Usage: spfx-fast agent <init|doctor>");
  }
  else if (command === "create") await createApp(args.shift(), args);
  else throw new Error(`Unknown command: ${command}`);
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
}
