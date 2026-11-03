#!/usr/bin/env node
import { access, mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";

const sensitivePatterns = [
  [/\b[\w.+-]+@[\w.-]+\.[A-Z]{2,}\b/i, "email address"],
  [/\b[a-f\d]{8}-(?:[a-f\d]{4}-){3}[a-f\d]{12}\b/i, "GUID"],
  [/\b[a-z\d-]+\.sharepoint(?:-df)?\.com\b|\b[a-z\d-]+\.onmicrosoft\.com\b/i, "tenant domain"],
  [/\/(?:sites|teams)\/(?!your-site(?:\/|$)|your-team(?:\/|$)|example(?:\/|$))[a-z\d._%-]+/i, "site-specific path"],
  [/\bBearer\s+\S+/i, "bearer token"],
  [/\beyJ[a-zA-Z\d_-]{10,}\.[a-zA-Z\d_-]{10,}\.[a-zA-Z\d_-]{10,}\b/, "JWT-like token"],
  [/[?&](?:access_token|refresh_token|client_secret|code)=[^&\s]+/i, "credential query parameter"],
];
const kinds = new Set(["debugging", "gotcha", "ui-pattern", "api", "performance", "tool-workflow"]);
const levels = new Set(["unverified", "reproduced", "tenant-validated", "documented"]);
const help = `SPFx Fast Knowledge\n\nCommands:\n  learn init [--out <file>]                     Create a local contribution draft\n  learn import-playwright <file> [--out <file>] Convert a Playwright learning draft\n  learn validate <file>                          Validate fields and screen common identifiers\n  learn patch <file> [--out <file>]             Create a reviewable Git patch\n\nNo command uploads data. Validation is a screening aid, not an anonymity guarantee.\n`;

function slug(value) {
  const result = value.normalize("NFKD").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 72);
  if (!result) throw new TypeError("The title must contain letters or numbers to create a patch filename.");
  return result;
}

function validate(entry) {
  if (!entry || Array.isArray(entry) || typeof entry !== "object") throw new TypeError("Finding must be a JSON object.");
  const allowed = ["format", "version", "title", "kind", "symptom", "diagnostics", "rootCause", "fix", "reproduction", "appliesTo", "limitations", "validation", "recipe"];
  for (const key of Object.keys(entry)) if (!allowed.includes(key)) throw new TypeError(`Unsupported field: ${key}`);
  if (entry.format !== "spfx-fast-learning" || entry.version !== 1) throw new TypeError("format/version must be spfx-fast-learning/1.");
  if (!kinds.has(entry.kind)) throw new TypeError(`kind must be one of ${[...kinds].join(", ")}.`);
  for (const key of ["title", "symptom", "diagnostics", "rootCause", "fix", "appliesTo", "limitations"]) {
    if (typeof entry[key] !== "string" || !entry[key].trim() || entry[key].length > 2000) throw new TypeError(`${key} must be a non-empty string of at most 2000 characters.`);
  }
  if (!Array.isArray(entry.reproduction) || entry.reproduction.length < 1 || entry.reproduction.length > 12 || entry.reproduction.some((step) => typeof step !== "string" || !step.trim() || step.length > 500)) throw new TypeError("reproduction must contain 1 to 12 non-empty steps of at most 500 characters.");
  if (!entry.validation || Array.isArray(entry.validation) || typeof entry.validation !== "object" || Object.keys(entry.validation).some((key) => !["level", "checkedOn", "sampleSize"].includes(key))) throw new TypeError("validation must contain level and optional checkedOn/sampleSize fields.");
  if (!levels.has(entry.validation.level)) throw new TypeError(`validation.level must be one of ${[...levels].join(", ")}.`);
  if (entry.validation.checkedOn !== undefined) {
    const date = new Date(`${entry.validation.checkedOn}T00:00:00Z`);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(entry.validation.checkedOn) || Number.isNaN(date.valueOf()) || date.toISOString().slice(0, 10) !== entry.validation.checkedOn) throw new TypeError("validation.checkedOn must be a valid YYYY-MM-DD date.");
  }
  if (entry.validation.level !== "unverified" && !entry.validation.checkedOn) throw new TypeError("Set validation.checkedOn for validated findings.");
  if (entry.validation.sampleSize !== undefined && (!Number.isInteger(entry.validation.sampleSize) || entry.validation.sampleSize < 1 || entry.validation.sampleSize > 10000)) throw new TypeError("validation.sampleSize must be an integer from 1 to 10000.");
  const text = JSON.stringify(entry);
  for (const [pattern, label] of sensitivePatterns) if (pattern.test(text)) throw new TypeError(`Possible ${label}; replace it with generic details before sharing.`);
  if (entry.recipe !== undefined && (typeof entry.recipe !== "string" || entry.recipe.length > 4000 || /https?:\/\/|[?&](?:access_token|code|client_secret)=/i.test(entry.recipe))) throw new TypeError("recipe must be short, generic text without URLs or credential parameters.");
  return entry;
}

function draft() {
  return {
    format: "spfx-fast-learning",
    version: 1,
    title: "Describe the reusable discovery",
    kind: "debugging",
    symptom: "What did the developer or agent observe?",
    diagnostics: "What evidence distinguished the cause from similar failures?",
    rootCause: "What caused the behavior?",
    fix: "What should another developer or agent do?",
    reproduction: ["Describe generic setup and steps. Do not include tenant-specific details."],
    appliesTo: "SPFx version, host, browser, and relevant service surface",
    limitations: "What remains unverified, and where may this not apply?",
    validation: { level: "unverified" },
    recipe: "Optional generic command or check; omit credentials and tenant URLs.",
  };
}

function patchFor(entry) {
  const path = `knowledge/findings/${slug(entry.title)}.json`;
  const content = `${JSON.stringify(entry, null, 2)}\n`;
  const lines = content.split("\n").filter((line, index, all) => index < all.length - 1).map((line) => `+${line}`).join("\n");
  return { path, content: `diff --git a/${path} b/${path}\nnew file mode 100644\nindex 0000000..0000000\n--- /dev/null\n+++ b/${path}\n@@ -0,0 +1,${content.split("\n").length - 1} @@\n${lines}\n` };
}

async function main() {
  const [command, action, file] = process.argv.slice(2);
  const outArg = process.argv.indexOf("--out");
  const out = outArg >= 0 ? process.argv[outArg + 1] : undefined;
  if (!command || command === "--help" || command === "help") return console.log(help);
  if (command !== "learn") throw new Error(help.trim());
  if (action === "init") {
    const target = resolve(out ?? file ?? "test-results/spfx-fast-knowledge/learning-draft.json");
    await mkdir(dirname(target), { recursive: true });
    try { await access(target); throw new Error(`Refusing to overwrite ${target}. Choose another --out path.`); }
    catch (error) { if (error.message.startsWith("Refusing")) throw error; }
    await writeFile(target, `${JSON.stringify(draft(), null, 2)}\n`, "utf8");
    console.log(`Created local draft: ${target}`);
    return;
  }
  if (action === "import-playwright") {
    if (!file) throw new Error("Usage: spfx-fast-knowledge learn import-playwright <file> [--out <file>]");
    let source;
    try { source = JSON.parse(await readFile(resolve(file), "utf8")); }
    catch { throw new Error("Could not read a Playwright learning JSON file."); }
    const topicKinds = { ui: "ui-pattern", accessibility: "ui-pattern", navigation: "ui-pattern", deployment: "debugging", permissions: "debugging", data: "api", performance: "performance", other: "gotcha" };
    const converted = validate({
      format: "spfx-fast-learning",
      version: 1,
      title: source.title,
      kind: topicKinds[source.topic] ?? "gotcha",
      symptom: source.problem,
      diagnostics: (source.reproduction ?? []).join(" "),
      rootCause: source.finding,
      fix: source.recommendation,
      reproduction: source.reproduction,
      appliesTo: JSON.stringify(source.appliesTo),
      limitations: source.limitations,
      validation: source.validation,
      ...(source.recipe ? { recipe: JSON.stringify(source.recipe) } : {}),
    });
    const target = resolve(out ?? "test-results/spfx-fast-knowledge/imported-learning.json");
    await mkdir(dirname(target), { recursive: true });
    try { await access(target); throw new Error(`Refusing to overwrite ${target}. Choose another --out path.`); }
    catch (error) { if (error.message.startsWith("Refusing")) throw error; }
    await writeFile(target, `${JSON.stringify(converted, null, 2)}\n`, "utf8");
    console.log(`Imported local draft: ${target}`);
    console.log("Review every mapped field, then validate and patch it. No upload occurred.");
    return;
  }
  if (action !== "validate" && action !== "patch") throw new Error(help.trim());
  if (!file) throw new Error(`Usage: spfx-fast-knowledge learn ${action} <finding.json>${action === "patch" ? " [--out <patch>]" : ""}`);
  let entry;
  try { entry = validate(JSON.parse(await readFile(resolve(file), "utf8"))); }
  catch (error) { if (error instanceof TypeError) throw error; throw new Error("Could not read a valid JSON finding."); }
  console.log(`Finding valid: ${entry.title} (${entry.kind}, ${entry.validation.level}).`);
  console.log("Review all text and evidence manually. This check cannot guarantee anonymity; no upload occurred.");
  if (action === "patch") {
    const patch = patchFor(entry);
    const target = resolve(out ?? `${slug(entry.title)}.patch`);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, patch.content, { encoding: "utf8", flag: "wx" });
    console.log(`Wrote contribution patch: ${target}`);
    console.log(`The patch adds only ${patch.path}; it contains no identity or attachments.`);
  }
}

main().catch((error) => { console.error(error.message); process.exitCode = 1; });
