import { createRequire } from "node:module";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";

const [siteArg, profileArg, productId = "5f68f9b1-18bb-47e8-a85f-9957dce645e4"] = process.argv.slice(2);
if (!siteArg || !profileArg) throw new Error("Usage: node scripts/retry-sharepoint-app-instance-install.mjs <site-url> <signed-in-profile-dir> [product-id]");
const site = new URL(siteArg);
if (site.protocol !== "https:") throw new TypeError("The SharePoint site URL must use HTTPS.");
const profileDir = resolve(profileArg);
const repo = "/Users/vilius/.t3/projects/spfx-research";
const require = createRequire(import.meta.url);
const { chromium } = require(`${repo}/node_modules/playwright`);
const { acquireProfileLock } = await import(pathToFileURL(`${repo}/packages/playwright/src/profile-lock.js`));
const unlock = acquireProfileLock(profileDir);
let context;
try {
  context = await chromium.launchPersistentContext(profileDir, { headless: true, serviceWorkers: "block" });
  const page = context.pages()[0] ?? await context.newPage();
  await page.goto(site.href, { waitUntil: "domcontentloaded", timeout: 60000 });
  if (page.url().includes("login.microsoftonline.com")) throw new Error("The signed-in SharePoint profile is not authenticated.");
  const result = await page.evaluate(async ({ productId }) => {
    const api = `${location.origin}${location.pathname.replace(/\/$/, "")}/_api`;
    const headers = { Accept: "application/json;odata=nometadata" };
    const readInstances = async () => {
      const response = await fetch(`${api}/web/GetAppInstancesByProductId('${productId}')`, { headers });
      if (!response.ok) throw new Error(`App instance read failed (${response.status}).`);
      return (await response.json()).value ?? [];
    };
    let instances = await readInstances();
    if (instances.length !== 1) throw new Error(`Expected one site instance; found ${instances.length}. Refusing an ambiguous retry.`);
    let instance = instances[0];
    if (instance.Status !== 9) throw new Error(`Retry is limited to Initialized (9); current status is ${instance.Status}.`);
    const digestResponse = await fetch(`${api}/contextinfo`, { method: "POST", headers });
    if (!digestResponse.ok) throw new Error(`Context digest failed (${digestResponse.status}).`);
    const digest = (await digestResponse.json()).FormDigestValue;
    const installResponse = await fetch(`${api}/web/GetAppInstanceById('${instance.Id}')/Install`, { method: "POST", headers: { ...headers, "X-RequestDigest": digest } });
    const installBody = (await installResponse.text().catch(() => "")).slice(0, 500);
    if (!installResponse.ok) throw new Error(`Instance Install failed (${installResponse.status}): ${installBody}`);
    for (let attempt = 0; attempt < 48; attempt++) {
      await new Promise(resolve => setTimeout(resolve, 2500));
      instances = await readInstances();
      instance = instances.find(item => item.Id === instance.Id) ?? instance;
      if (instance.Status === 5) return { actionStatus: installResponse.status, instanceId: instance.Id, status: instance.Status, statusName: "Installed", attempts: attempt + 1, installBody };
      if (instance.Status !== 9 && instance.Status !== 1) throw new Error(`Install moved to unexpected status ${instance.Status}.`);
    }
    throw new Error(`Instance Install returned ${installResponse.status} but status did not reach Installed: ${JSON.stringify({ instanceId: instance.Id, status: instance.Status, installBody })}`);
  }, { productId });
  console.log(JSON.stringify(result, null, 2));
} finally {
  await context?.close();
  unlock();
}
