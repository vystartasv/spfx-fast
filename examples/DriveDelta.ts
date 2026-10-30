import type { createClient } from "@spfx-fast/core";

type Client = ReturnType<typeof createClient>;

/** Run one bounded delta segment; call again while complete is false to resume. */
export async function synchronizeDrive(client: Client, driveId: string, maxPages = 100) {
  const sync = await client.syncDrive(driveId, { maxPages });
  return { ...sync, items: await client.getDriveItems(driveId) };
}
