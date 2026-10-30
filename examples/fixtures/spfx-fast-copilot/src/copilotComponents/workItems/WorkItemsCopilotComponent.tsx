import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';

import { BaseCopilotComponent } from '@microsoft/sp-copilot-component';
import { SPHttpClient } from '@microsoft/sp-http';
import { createSPFxServices } from '@spfx-fast/core/spfx';
import type { createClient } from '@spfx-fast/core';

import WorkItems from './components/WorkItems';
import type { IWorkItemsProps } from './components/IWorkItemsProps';
import type { IWorkItemsCopilotComponentProperties } from './WorkItemsCopilotComponentProperties';

import * as strings from 'WorkItemsCopilotComponentStrings';

/**
 * SPFx Copilot Component that renders a React-based UI demonstrating the
 * platform's headline capabilities:
 *
 * - **Brokered SSO data calls** — Microsoft Graph (`/me`) and SharePoint REST
 *   (`/_api/web`) with zero token code. The SPFx runtime's Pairwise Broker
 *   automatically provisions tokens for `SPHttpClient` and `MSGraphClientV3`.
 *
 * - **Host context & theming** — reads `hostContext.theme` and
 *   `hostContext.displayMode` to adapt to the Copilot host environment.
 *
 * - **Bridge actions** — demonstrates `requestDisplayModeAsync`,
 *   `openLinkAsync`, `sendFollowUpMessageAsync`, and `requestSizeChangeAsync`.
 *
 * Lifecycle:
 *  1. `onInit()` — fetches user and site data (runs once before first render).
 *  2. `render()` — mounts the React tree into `this.context.domElement`.
 *     Re-invoked by the framework on host-context changes.
 *  3. `onTeardown()` — unmounts React before the host tears down the iframe.
 */
export default class WorkItemsCopilotComponent extends BaseCopilotComponent<IWorkItemsCopilotComponentProperties> {
  private _client: ReturnType<typeof createClient> | undefined;
  private _siteId: string | undefined;
  private _listId: string | undefined;
  private _siteUrl: string = '';
  private _listTitle: string = '';
  private _loading: boolean = false;
  private _error: string | undefined;
  private _requestKey: string = '';
  private _requestId: number = 0;
  private _root: Root | undefined;

  protected async onInit(): Promise<void> {
    const services = await createSPFxServices(this.context, SPHttpClient.configurations.v1);
    this._client = services.client;
  }

  protected render(): void {
    const requestKey = JSON.stringify([this.properties.siteUrl ?? '', this.properties.listTitle ?? '']);
    if (requestKey !== this._requestKey) {
      this._requestKey = requestKey;
      this._siteId = undefined;
      this._listId = undefined;
      this._error = undefined;
      this._loadTarget(requestKey, this.properties.siteUrl, this.properties.listTitle);
    }

    const props: IWorkItemsProps = {
      client: this._client,
      siteId: this._siteId,
      listId: this._listId,
      listTitle: this._listTitle,
      siteUrl: this._siteUrl,
      hostContext: this.hostContext,
      bridge: this.context.copilotBridge,
      loading: this._loading,
      error: this._error,
      onRequestDisplayMode: async (mode) => {
        await this.requestDisplayModeAsync(mode);
      },
      targetDocument: this.context.domElement.ownerDocument,
      strings
    };

    if (!this._root) {
      this._root = createRoot(this.context.domElement);
    }

    this._root.render(React.createElement(WorkItems, props));
  }

  private _loadTarget(requestKey: string, requestedSiteUrl: string | undefined, requestedListTitle: string | undefined): void {
    const requestId = ++this._requestId;
    this._loading = true;
    this._error = undefined;
    this._siteUrl = '';
    this._listTitle = (requestedListTitle ?? '').trim();
    this.renderView();

    this._resolveTarget(requestedSiteUrl, requestedListTitle).then((target) => {
      if (requestId !== this._requestId || requestKey !== this._requestKey) return;
      this._siteId = target.siteId;
      this._listId = target.listId;
      this._siteUrl = target.siteUrl;
      this._listTitle = target.listTitle;
      this._loading = false;
      this.renderView();
    }, (reason: unknown) => {
      if (requestId !== this._requestId || requestKey !== this._requestKey) return;
      this._error = reason instanceof Error ? reason.message : String(reason);
      this._loading = false;
      this.renderView();
    });
  }

  private async _resolveTarget(requestedSiteUrl: string | undefined, requestedListTitle: string | undefined): Promise<{ siteId: string; listId: string; siteUrl: string; listTitle: string }> {
    if (!this._client) throw new Error('The authenticated SPFx services are not ready yet.');
    const listTitle = (requestedListTitle ?? '').trim();
    if (!listTitle) throw new Error('Provide the exact SharePoint list title.');
    const siteUrl = resolveSiteUrl(requestedSiteUrl, this.context.pageContext.web.absoluteUrl);
    const parsed = new URL(siteUrl);
    const site = await this._client.getSite(parsed.hostname, parsed.pathname) as { id: string };
    const list = await this._client.findList(site.id, listTitle) as { id: string; displayName?: string } | undefined;
    if (!list?.id) throw new Error(`No accessible list titled “${listTitle}” was found on this site.`);
    return { siteId: site.id, listId: list.id, siteUrl, listTitle: list.displayName ?? listTitle };
  }

  private renderView(): void {
    if (this._root) this._root.render(React.createElement(WorkItems, {
      client: this._client,
      siteId: this._siteId,
      listId: this._listId,
      listTitle: this._listTitle,
      siteUrl: this._siteUrl,
      hostContext: this.hostContext,
      bridge: this.context.copilotBridge,
      loading: this._loading,
      error: this._error,
      onRequestDisplayMode: async (mode) => {
        await this.requestDisplayModeAsync(mode);
      },
      targetDocument: this.context.domElement.ownerDocument,
      strings
    }));
  }

  protected async onTeardown(): Promise<void> {
    this._requestId++;
    this._root?.unmount();
    this._root = undefined;
  }
}

function resolveSiteUrl(value: string | undefined, fallback: string): string {
  const candidate = (value ?? '').trim() || fallback;
  let parsed: URL;
  try { parsed = new URL(candidate); }
  catch { throw new Error('Use a valid absolute SharePoint site URL.'); }
  const host = parsed.hostname.toLocaleLowerCase();
  const isSharePointHost = host.endsWith('.sharepoint.com') || host.endsWith('.sharepoint.us') || host.endsWith('.sharepoint.cn') || host.endsWith('.sharepoint.de') || host.endsWith('.sharepoint-mil.us');
  if (parsed.protocol !== 'https:' || !isSharePointHost || parsed.search || parsed.hash) {
    throw new Error('The site URL must be an HTTPS SharePoint site URL without a query or fragment.');
  }
  if (/\.aspx$/i.test(parsed.pathname) || /\/(Lists|SitePages)\//i.test(parsed.pathname)) {
    throw new Error('Pass the site URL only; remove the page or list path.');
  }
  return `${parsed.origin}${parsed.pathname}`.replace(/\/+$/, '');
}
