import { validateRecipe } from "./index.js";

const fields = {
  root: ["format", "version", "title", "topic", "appliesTo", "problem", "finding", "reproduction", "recommendation", "limitations", "validation", "recipe"],
  appliesTo: ["service", "surface", "spfxVersion", "browser"],
  validation: ["level", "checkedOn", "sampleSize"],
};
const topics = new Set(["ui", "accessibility", "navigation", "deployment", "permissions", "data", "performance", "other"]);
const levels = new Set(["unverified", "reproduced", "tenant-validated", "documented"]);

function rejectUnknownKeys(value, allowed, path) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${path} must be an object.`);
  for (const key of Object.keys(value)) if (!allowed.includes(key)) throw new TypeError(`${path}.${key} is not a supported field.`);
}

export function createLearningDraft() {
  return {
    format: "spfx-fast-learning",
    version: 1,
    title: "Describe the reusable SharePoint discovery",
    topic: "ui",
    appliesTo: { service: "SharePoint Online", surface: "SPFx web part", spfxVersion: "1.24+", browser: "Chromium" },
    problem: "What was difficult to discover?",
    finding: "What behavior or reliable pattern did you observe?",
    reproduction: ["Describe the setup and steps without tenant-specific details."],
    recommendation: "What should another agent or developer do?",
    limitations: "State what has not been verified and where this may not apply.",
    validation: { level: "unverified" },
    recipe: {
      name: "Replace with a focused, reusable UI check",
      steps: [
        { action: "navigate", path: "/sites/your-site" },
        { action: "assertPageReady" },
        { action: "assertAuthenticated" },
      ],
    },
  };
}

function checkKeys(value, allowed, path) {
  rejectUnknownKeys(value, allowed, path);
  for (const key of allowed) {
    if (!Object.hasOwn(value, key) && !(path === "validation" && key !== "level")) {
      throw new TypeError(`${path}.${key} is required.`);
    }
  }
}

function strings(value, path, found = []) {
  if (typeof value === "string") {
    const patterns = [
      [/[\w.+-]+@[\w.-]+\.[A-Z]{2,}/i, "email address"],
      [/\b[a-f\d]{8}-(?:[a-f\d]{4}-){3}[a-f\d]{12}\b/i, "GUID"],
      [/\b[a-z\d-]+\.sharepoint(?:-df)?\.com\b|\b[a-z\d-]+\.onmicrosoft\.com\b/i, "tenant domain"],
      [/\/(?:sites|teams)\/(?!your-site(?:\/|$)|your-team(?:\/|$)|example(?:\/|$))[a-z\d._%-]+/i, "site-specific path"],
      [/\bBearer\s+\S+/i, "bearer token"],
      [/\beyJ[a-zA-Z\d_-]{10,}\.[a-zA-Z\d_-]{10,}\.[a-zA-Z\d_-]{10,}\b/, "JWT-like token"],
      [/[?&](?:access_token|refresh_token|client_secret|code)=[^&\s]+/i, "credential query parameter"],
    ];
    for (const [pattern, label] of patterns) if (pattern.test(value)) found.push({ path, label });
  } else if (Array.isArray(value)) {
    value.forEach((item, index) => strings(item, `${path}[${index}]`, found));
  } else if (value && typeof value === "object") {
    for (const [key, item] of Object.entries(value)) strings(item, `${path}.${key}`, found);
  }
  return found;
}

function requireText(value, path, max = 2000) {
  if (typeof value !== "string" || !value.trim() || value.length > max) throw new TypeError(`${path} must be a non-empty string of at most ${max} characters.`);
}

export function validateLearning(learning) {
  if (!learning || typeof learning !== "object" || Array.isArray(learning)) throw new TypeError("Learning must be a JSON object.");
  const sensitive = strings(learning, "learning");
  if (sensitive.length) throw new TypeError(`Possible ${sensitive[0].label} at ${sensitive[0].path}. Replace it with generic details before sharing.`);

  checkKeys(learning, fields.root, "learning");
  if (learning.format !== "spfx-fast-learning" || learning.version !== 1) throw new TypeError("Learning format/version must be spfx-fast-learning/1.");
  requireText(learning.title, "learning.title", 120);
  if (!topics.has(learning.topic)) throw new TypeError(`learning.topic must be one of: ${[...topics].join(", ")}.`);
  checkKeys(learning.appliesTo, fields.appliesTo, "learning.appliesTo");
  for (const key of fields.appliesTo) requireText(learning.appliesTo[key], `learning.appliesTo.${key}`, 120);
  requireText(learning.problem, "learning.problem");
  requireText(learning.finding, "learning.finding");
  if (!Array.isArray(learning.reproduction) || learning.reproduction.length < 1 || learning.reproduction.length > 12) throw new TypeError("learning.reproduction must contain 1 to 12 steps.");
  learning.reproduction.forEach((step, index) => requireText(step, `learning.reproduction[${index}]`, 500));
  requireText(learning.recommendation, "learning.recommendation");
  requireText(learning.limitations, "learning.limitations");
  checkKeys(learning.validation, fields.validation, "learning.validation");
  if (!levels.has(learning.validation.level)) throw new TypeError(`learning.validation.level must be one of: ${[...levels].join(", ")}.`);
  if (learning.validation.checkedOn !== undefined) {
    const checkedOn = new Date(`${learning.validation.checkedOn}T00:00:00Z`);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(learning.validation.checkedOn) || Number.isNaN(checkedOn.valueOf()) || checkedOn.toISOString().slice(0, 10) !== learning.validation.checkedOn) throw new TypeError("learning.validation.checkedOn must be a valid YYYY-MM-DD date.");
  }
  if (learning.validation.level !== "unverified" && !learning.validation.checkedOn) throw new TypeError("Set validation.checkedOn for reproduced, tenant-validated, and documented findings.");
  if (learning.validation.sampleSize !== undefined && (!Number.isInteger(learning.validation.sampleSize) || learning.validation.sampleSize < 1 || learning.validation.sampleSize > 10000)) throw new TypeError("learning.validation.sampleSize must be an integer from 1 to 10000.");
  try {
    validateRecipe(learning.recipe);
    rejectUnknownKeys(learning.recipe, ["name", "steps"], "learning.recipe");
    learning.recipe.steps.forEach((step, index) => {
      const stepPath = `learning.recipe.steps[${index}]`;
      rejectUnknownKeys(step, ["action", "path", "query", "locator", "value", "confirm", "checked", "expected", "exact", "name", "fullPage"], stepPath);
      if (step.locator) rejectUnknownKeys(step.locator, ["role", "text", "label", "testId", "placeholder", "name", "exact"], `${stepPath}.locator`);
      if (step.query && Object.keys(step.query).some((key) => key !== "debugManifestsFile")) throw new TypeError(`${stepPath}.query may only contain the local debugManifestsFile parameter.`);
    });
  }
  catch (error) { throw new TypeError(`learning.recipe is invalid: ${error.message}`); }
  return learning;
}
