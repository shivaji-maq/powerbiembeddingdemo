import { models } from "powerbi-client";
import { PowerBIEmbed } from "powerbi-client-react";
import type { EmbedReportProps } from "./@type";

const EmbedReport: React.FC<EmbedReportProps> = ({
  reportRef,
  reportId,
  embedUrl,
  pageId,
  embedReportEventHandlers,
  reportSettings,
  reportFilters,
  themeJson,
  reportCssClassName,
  accessToken,
  tokenType,
  onReportLoadReportAttachmentFunction,
}) => {
  const embedConfig: models.IReportEmbedConfiguration = {
    type: "report",
    id: reportId,
    embedUrl: embedUrl,
    accessToken: accessToken,
    tokenType: tokenType ? models.TokenType.Embed : models.TokenType.Aad,
    permissions: models.Permissions.All,
    settings: { panes: { filters: { expanded: false, visible: true } }, background: models.BackgroundType.Transparent, visualRenderedEvents: true },
  };

  if (pageId) embedConfig.pageName = pageId;
  if (reportFilters && reportFilters.length > 0) embedConfig.filters = reportFilters;
  if (reportSettings) embedConfig.settings = reportSettings;
  if (themeJson) embedConfig.theme = { themeJson };
  return (
    <PowerBIEmbed
      embedConfig={embedConfig}
      cssClassName={`${reportCssClassName} report-embed`}
      getEmbeddedComponent={(embeddedComponent) => {
        if (reportRef) reportRef.current = embeddedComponent;
        onReportLoadReportAttachmentFunction(embeddedComponent);
      }}
      eventHandlers={new Map(embedReportEventHandlers)}
    />
  );
};

export default EmbedReport;
