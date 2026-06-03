import type { models } from "powerbi-client";
import type { EventHandler } from "powerbi-client-react";

export interface EmbedReportProps {
  reportRef: any;
  reportId: string;
  embedUrl: string;
  pageId?: string;
  embedReportEventHandlers: Map<string, EventHandler>;
  reportSettings?: models.ISettings;
  reportFilters?: models.ReportLevelFilters[];
  themeJson?: Record<string, unknown>;
  bookmarkState?: string;
  reportCssClassName?: string;
  accessToken: string;
  tokenType?: string;
  onReportLoadReportAttachmentFunction: any;
}
