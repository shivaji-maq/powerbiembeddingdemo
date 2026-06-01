import type { models } from "powerbi-client";
import type { EventHandler } from "powerbi-client-react";

export interface EmbedReportProps {
  reportRef: any;
  reportId: string;
  embedUrl: string;
  pageId?: string;
  embedReportEventHandlers: Map<string, EventHandler>;
  reportSettings?: models.ISettings;
  themeJson?: Record<string, unknown>;
  reportCssClassName?: string;
  accessToken: string;
  tokenType?: string;
  onReportLoadReportAttachmentFunction: any;
}
