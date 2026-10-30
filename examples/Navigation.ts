import { createLegacySharePointApi, type SharePointNavigationNode, type SharePointPage } from "@spfx-fast/core/sharepoint";
import type { SharePointRestTransport } from "@spfx-fast/core/spfx";

/** Read one bounded page of a web's Quick Launch, including one level of child links. */
export async function readQuickLaunch(rest: SharePointRestTransport): Promise<SharePointPage<SharePointNavigationNode>> {
  return createLegacySharePointApi(rest).getNavigationNodesPage({
    scope: "web",
    bar: "quickLaunch",
    includeChildren: true,
    pageSize: 50,
  });
}

/** Continue from the opaque cursor returned by the previous page. */
export async function readNextQuickLaunchPage(rest: SharePointRestTransport, cursor: string): Promise<SharePointPage<SharePointNavigationNode>> {
  return createLegacySharePointApi(rest).getNavigationNodesPage({ cursor, includeChildren: true, pageSize: 50 });
}

/** Read the site collection root web's top navigation instead of the current subweb's links. */
export async function readSiteTopNavigation(rest: SharePointRestTransport): Promise<SharePointPage<SharePointNavigationNode>> {
  return createLegacySharePointApi(rest).getNavigationNodesPage({ scope: "site", bar: "topNavigationBar", pageSize: 50 });
}
