import { rmSync } from "node:fs";

// Heft does not remove generated Teams app ZIPs. A stale ZIP is copied into the
// .sppkg and can leave an obsolete agent alongside the current one.
rmSync(new URL("../teams", import.meta.url), { recursive: true, force: true });
