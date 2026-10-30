import { createHash } from "node:crypto";
import { hostname, tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { mkdirSync, readFileSync, readlinkSync, renameSync, rmSync, writeFileSync } from "node:fs";

function processExists(pid) {
  try { process.kill(pid, 0); return true; }
  catch (error) { return error?.code === "EPERM"; }
}

function assertProfileFree(profileDir) {
  if (process.platform === "win32") return;
  let value;
  try { value = readlinkSync(join(profileDir, "SingletonLock")); }
  catch { return; }
  const separator = value.lastIndexOf("-");
  const ownerHost = value.slice(0, separator);
  const ownerPid = Number(value.slice(separator + 1));
  if (ownerHost === hostname() && Number.isInteger(ownerPid) && processExists(ownerPid)) {
    throw new Error("The supplied browser profile is already open. Close that browser before running Playwright.");
  }
}

/** Serialize SPFx Fast jobs using one persistent browser profile. Never removes Chrome profile files. */
export function acquireProfileLock(profileDirectory) {
  const profileDir = resolve(profileDirectory);
  assertProfileFree(profileDir);
  const hash = createHash("sha256").update(profileDir).digest("hex").slice(0, 24);
  const lockDir = join(tmpdir(), `spfx-fast-playwright-${hash}.lock`);
  const owner = JSON.stringify({ host: hostname(), pid: process.pid });
  try {
    mkdirSync(lockDir);
    writeFileSync(join(lockDir, "owner.json"), owner, { flag: "wx", mode: 0o600 });
  } catch (error) {
    if (error?.code === "EEXIST") {
      let previous;
      try { previous = JSON.parse(readFileSync(join(lockDir, "owner.json"), "utf8")); }
      catch { throw new Error("A Playwright job may be starting with this browser profile; retry after it finishes."); }
      if (previous.host !== hostname() || processExists(Number(previous.pid))) {
        throw new Error("Another SPFx Fast Playwright job is using this browser profile.");
      }
      const staleDir = `${lockDir}.stale-${process.pid}-${Date.now()}`;
      try { renameSync(lockDir, staleDir); }
      catch (renameError) {
        if (renameError?.code === "ENOENT") return acquireProfileLock(profileDir);
        throw new Error("A Playwright job may be opening this browser profile; retry after it finishes.");
      }
      try {
        if (readFileSync(join(staleDir, "owner.json"), "utf8") !== JSON.stringify(previous)) {
          try { renameSync(staleDir, lockDir); } catch { /* Another run already restored or replaced the lock. */ }
          throw new Error("The browser-profile lock changed while checking it; retry after the active job finishes.");
        }
      } catch (verifyError) {
        if (verifyError instanceof Error && verifyError.message.startsWith("The browser-profile lock changed")) throw verifyError;
        throw new Error("Could not verify the browser-profile lock owner; retry after the active job finishes.");
      }
      rmSync(staleDir, { recursive: true, force: true });
      return acquireProfileLock(profileDir);
    }
    rmSync(lockDir, { recursive: true, force: true });
    throw error;
  }
  let released = false;
  return () => {
    if (released) return;
    released = true;
    try {
      const current = readFileSync(join(lockDir, "owner.json"), "utf8");
      if (current === owner) rmSync(lockDir, { recursive: true, force: true });
    } catch { /* The lock was already released or replaced. */ }
  };
}
