import { createLegacySharePointApi, type SharePointRecycleBinItem } from "@spfx-fast/core/sharepoint";
import type { SharePointRestTransport } from "@spfx-fast/core/spfx";

/** Read one bounded page from the signed-in user's hosting-web recycle bin. */
export function getDeletedItems(rest: SharePointRestTransport, cursor?: string) {
  return createLegacySharePointApi(rest).getRecycleBinItemsPage({ cursor, pageSize: 50 });
}

/** Restore an inspected item to its original location. */
export function restoreDeletedItem(rest: SharePointRestTransport, item: SharePointRecycleBinItem) {
  return createLegacySharePointApi(rest).restoreRecycleBinItem(item.Id);
}

/** Permanently delete only after explicit user confirmation. */
export function permanentlyDeleteAfterConfirmation(rest: SharePointRestTransport, item: SharePointRecycleBinItem, confirmed: boolean) {
  if (!confirmed) throw new Error("Confirm permanent deletion before continuing.");
  return createLegacySharePointApi(rest).permanentlyDeleteRecycleBinItem(item.Id);
}
