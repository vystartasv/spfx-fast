import * as React from "react";
import { createRoot, type Root } from "react-dom/client";
import { Version } from "@microsoft/sp-core-library";
import { SPHttpClient } from "@microsoft/sp-http";
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import type { createClient } from "@spfx-fast/core";
import { createSPFxServices } from "@spfx-fast/core/spfx";
import SpfxFastDashboard from "./components/SpfxFastDashboard";

export interface ISpfxFastDashboardWebPartProps {
  taskListId: string;
  projectListId: string;
  folderItemId?: string;
}

export default class SpfxFastDashboardWebPart extends BaseClientSideWebPart<ISpfxFastDashboardWebPartProps> {
  private root?: Root;
  private client?: ReturnType<typeof createClient>;
  private services?: Awaited<ReturnType<typeof createSPFxServices>>;
  private siteId = "";
  private driveId = "";
  private webTitle = "SharePoint";

  protected async onInit(): Promise<void> {
    await super.onInit();
    const services = await createSPFxServices(this.context, SPHttpClient.configurations.v1);
    this.services = services;
    this.client = services.client;
    const rest = services.rest;
    const web = await rest.get("/_api/web?$select=Title") as { Title?: string };
    this.webTitle = web.Title ?? this.webTitle;
    const webUrl = new URL(this.context.pageContext.web.absoluteUrl);
    const site = await this.client.getSite(webUrl.hostname, webUrl.pathname) as { id: string };
    this.siteId = site.id;
    try {
      const drive = await this.client.requestGraph(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(site.id)}/drive?$select=id,name`) as { id?: string };
      this.driveId = drive.id ?? "";
    } catch {
      // Keep list dashboards available when the signed-in user cannot read the default document library.
      this.driveId = "";
    }
  }

  public render(): void {
    this.root ??= createRoot(this.domElement);
    this.root.render(this.client && this.services ? <SpfxFastDashboard
      client={this.client}
      services={this.services}
      siteId={this.siteId}
      webTitle={this.webTitle}
      taskListId={this.properties.taskListId ?? ""}
      projectListId={this.properties.projectListId ?? ""}
      driveId={this.driveId}
      folderItemId={this.properties.folderItemId || "root"}
    /> : <div role="status">Connecting to SharePoint…</div>);
  }

  protected onDispose(): void {
    this.root?.unmount();
    this.root = undefined;
  }

  protected get dataVersion(): BaseClientSideWebPart<ISpfxFastDashboardWebPartProps>["dataVersion"] {
    return Version.parse("1.0") as unknown as BaseClientSideWebPart<ISpfxFastDashboardWebPartProps>["dataVersion"];
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: "Configure the Graph list IDs and optional document-library folder." },
        groups: [{
          groupName: "SharePoint lists",
          groupFields: [
            PropertyPaneTextField("taskListId", { label: "Tasks list ID" }),
            PropertyPaneTextField("projectListId", { label: "Projects list ID" }),
            PropertyPaneTextField("folderItemId", { label: "Document library folder item ID (root by default)" }),
          ],
        }],
      }],
    };
  }
}
