import type { createClient } from "@spfx-fast/core";
import type { createSPFxServices } from "@spfx-fast/core/spfx";

export interface ISpfxFastDashboardProps {
  client: ReturnType<typeof createClient>;
  services: Awaited<ReturnType<typeof createSPFxServices>>;
  siteId: string;
  webTitle: string;
  taskListId: string;
  projectListId: string;
  driveId?: string;
  folderItemId?: string;
}
