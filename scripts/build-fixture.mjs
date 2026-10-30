import { spawnSync } from "node:child_process";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

const fixtures = {
  dashboard: "examples/fixtures/spfx-fast-dashboard",
  copilot: "examples/fixtures/spfx-fast-copilot",
};
const fixture = fixtures[process.argv[2]];
if (!fixture) throw new Error(`Unknown fixture: ${process.argv[2] ?? ""}`);

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const directory = `${fixture}`;
const usePnpm = process.env.npm_config_user_agent?.startsWith("pnpm/") ?? false;
const manager = usePnpm ? "pnpm" : "npm";

function run(args) {
  const result = spawnSync(manager, args, { cwd: root, stdio: "inherit" });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}

if (usePnpm) {
  run(["--dir", directory, "install", "--frozen-lockfile"]);
  run(["--dir", directory, "run", "build"]);
} else {
  run(["ci", "--prefix", directory]);
  run(["run", "build", "--prefix", directory]);
}
