import type { createSPFxServices } from "../src/spfx.js";

type Services = Awaited<ReturnType<typeof createSPFxServices>>;

/** Start a same- or cross-drive copy and wait for its SharePoint operation monitor. */
export async function copyDriveFile(
  services: Services,
  sourceDriveId: string,
  sourceItemId: string,
  targetDriveId: string,
  targetFolderId: string,
  onProgress?: (percent: number) => void,
) {
  const operation = await services.client.copyDriveItem(sourceDriveId, sourceItemId, {
    targetDriveId,
    targetFolderId,
    conflictBehavior: "rename",
  });
  return services.waitForDriveCopy(operation, {
    onProgress: (status) => onProgress?.(status.percentComplete ?? status.percentageComplete ?? 0),
  });
}
