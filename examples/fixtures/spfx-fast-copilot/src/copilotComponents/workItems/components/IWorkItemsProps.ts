import type {
  ICopilotComponentHostContext,
  ISPCopilotBridge,
  SPCopilotDisplayMode
} from '@microsoft/sp-copilot-component';
import type { createClient } from '@spfx-fast/core';

export interface IWorkItemsStrings {
  ExpandButtonLabel: string;
  OpenSiteButtonLabel: string;
  FollowUpButtonLabel: string;
  ResizeButtonLabel: string;
  CompactButtonLabel: string;
  SiteBadgePrefix: string;
  ThemeBadgePrefix: string;
  ModeBadgePrefix: string;
  GreetingPrefix: string;
  UnknownTheme: string;
  DefaultDisplayMode: string;
  FollowUpMessage: string;
}

export interface IWorkItemsProps {
  client?: ReturnType<typeof createClient>;
  siteId?: string;
  listId?: string;
  listTitle: string;
  /** Absolute URL of the validated SharePoint site. */
  siteUrl: string;
  /** Host context (theme, display mode) from the Copilot host. */
  hostContext: ICopilotComponentHostContext;
  /** Bridge to communicate with the Copilot host (public API surface). */
  bridge: ISPCopilotBridge;
  loading: boolean;
  error?: string;
  /** Request the host to change display mode (e.g. 'fullscreen'). */
  onRequestDisplayMode: (mode: SPCopilotDisplayMode) => Promise<void>;
  /**
   * Document the FluentProvider should inject its theme styles into. Pass
   * `domElement.ownerDocument` so Griffel writes CSS into the correct iframe
   * document rather than the top-level page.
   */
  targetDocument: Document | undefined;
  /** Localized strings for UI labels. */
  strings: IWorkItemsStrings;
}
