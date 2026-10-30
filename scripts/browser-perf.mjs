import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { createServer } from "node:http";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";
import { chromium } from "playwright";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const scratch = mkdtempSync(join(tmpdir(), "spfx-fast-browser-perf-"));
const bundlePath = join(scratch, "harness.js");
const reportPath = process.env.SPFX_FAST_BROWSER_PERF_REPORT;
let browser;
let server;

try {
  await build({
    entryPoints: [join(root, "tests/browser/performance-harness.mjs")],
    outfile: bundlePath,
    bundle: true,
    format: "iife",
    platform: "browser",
    target: "chrome120",
    logLevel: "silent",
  });

  server = createServer((request, response) => {
    if (request.url === "/harness.js") {
      response.writeHead(200, { "content-type": "text/javascript; charset=utf-8", "cache-control": "no-store" });
      response.end(readFileSync(bundlePath));
      return;
    }
    response.writeHead(200, { "content-type": "text/html; charset=utf-8", "cache-control": "no-store" });
    response.end("<!doctype html><html lang=\"en\"><meta charset=\"utf-8\"><title>SPFx Fast browser benchmark</title><body><main id=\"app\"></main><script src=\"/harness.js\"></script></body></html>");
  });
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });

  browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  page.on("pageerror", (error) => console.error(`Browser error: ${error.message}`));
  const address = server.address();
  const origin = `http://127.0.0.1:${address.port}`;
  await page.goto(origin, { waitUntil: "load" });
  const result = await page.evaluate(() => globalThis.runSPFxFastBrowserBenchmark());
  if (new URL(result.origin).origin !== origin) throw new Error(`Benchmark ran on unexpected origin: ${result.origin}`);
  const output = JSON.stringify({ timestamp: new Date().toISOString(), ...result }, null, 2);
  console.log(output);
  if (reportPath) writeFileSync(resolve(reportPath), `${output}\n`, "utf8");
} finally {
  await browser?.close();
  if (server?.listening) await new Promise((resolve) => server.close(resolve));
  rmSync(scratch, { recursive: true, force: true });
}
