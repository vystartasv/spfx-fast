import type { createClient } from "@spfx-fast/core";

type Client = ReturnType<typeof createClient>;

/** Read bounded, uncached version metadata for a list item and a document-library file. */
export async function readVersionHistory(
  client: Client,
  siteId: string,
  listId: string,
  listItemId: string,
  driveId: string,
  driveItemId: string,
) {
  const [listItemVersions, fileVersions] = await Promise.all([
    client.getListItemVersions(siteId, listId, listItemId, { maxItems: 100 }),
    client.getDriveItemVersions(driveId, driveItemId, { maxItems: 100 }),
  ]);
  return { listItemVersions, fileVersions };
}

/** Download an uncached historical file version after the user selects it. */
export function downloadPreviousFileVersion(client: Client, driveId: string, driveItemId: string, versionId: string) {
  return client.downloadDriveItemVersion(driveId, driveItemId, versionId);
}

/** Call only after the user confirms that this file version should become current. */
export function restorePreviousFileVersion(client: Client, driveId: string, driveItemId: string, versionId: string) {
  return client.restoreDriveItemVersion(driveId, driveItemId, versionId);
}
