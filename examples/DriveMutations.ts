import type { createClient } from "../src/index.js";

type Client = ReturnType<typeof createClient>;

export function createFolder(client: Client, driveId: string, parentFolderId: string, name: string) {
  return client.createDriveFolder(driveId, parentFolderId, name, { conflictBehavior: "rename" });
}

export function moveFile(client: Client, driveId: string, fileId: string, destinationFolderId: string, name?: string) {
  return client.moveDriveItem(driveId, fileId, destinationFolderId, { name });
}

/** Call from a delete action only after the host UI or agent confirms the target. */
export function deleteToRecycleBin(client: Client, driveId: string, itemId: string, eTag?: string) {
  return client.deleteDriveItem(driveId, itemId, { eTag });
}
