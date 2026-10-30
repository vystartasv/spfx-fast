import { createLegacySharePointApi, type SharePointRegionalSettings } from "@spfx-fast/core/sharepoint";
import type { SharePointRestTransport } from "@spfx-fast/core/spfx";

/** Read locale, calendar, and work-week settings for the hosting SharePoint web. */
export function getWebRegionalSettings(rest: SharePointRestTransport): Promise<SharePointRegionalSettings> {
  return createLegacySharePointApi(rest).getRegionalSettings();
}
