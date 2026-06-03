import { useRef, useMemo, useCallback } from "react";
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
  bookmarkState,
  reportCssClassName,
  accessToken,
  tokenType,
  onReportLoadReportAttachmentFunction,
}) => {
  // Lock theme, filters, settings, and bookmark to their initial values so that
  // subsequent programmatic changes (via report.applyTheme, report.setFilters,
  // bookmarksManager.applyState, etc.) don't cause PowerBIEmbed to detect a
  // config change and re-embed the report.
  const initialThemeRef = useRef(themeJson);
  const initialFiltersRef = useRef(reportFilters);
  const initialSettingsRef = useRef(reportSettings);
  const initialBookmarkStateRef = useRef(bookmarkState);

  const embedConfig: models.IReportEmbedConfiguration = useMemo(() => {
    const config: models.IReportEmbedConfiguration = {
      type: "report",
      id: reportId,
      embedUrl: embedUrl,
      accessToken: accessToken,
      tokenType: tokenType ? models.TokenType.Embed : models.TokenType.Aad,
      permissions: models.Permissions.All,
      settings: initialSettingsRef.current || {
        panes: { filters: { expanded: false, visible: true } },
        background: models.BackgroundType.Transparent,
        visualRenderedEvents: true,
      },
    };

    if (pageId) config.pageName = pageId;
    if (initialFiltersRef.current && initialFiltersRef.current.length > 0) {
      config.filters = initialFiltersRef.current;
    }
    if (initialThemeRef.current) {
      config.theme = { themeJson: initialThemeRef.current };
    }
    // Apply bookmark state at embed time per Microsoft showcase pattern.
    // This ensures the report loads directly into the bookmarked view
    // without flashing the original/default state first.
    if (initialBookmarkStateRef.current) {
      (config as any).bookmark = { state: initialBookmarkStateRef.current };
    }
    return config;
  }, [reportId, embedUrl, accessToken, tokenType, pageId]);

  const eventHandlers = useMemo(
    () => new Map(embedReportEventHandlers),
    [embedReportEventHandlers]
  );

  const handleGetEmbeddedComponent = useCallback(
    (embeddedComponent: any) => {
      if (reportRef) reportRef.current = embeddedComponent;
      onReportLoadReportAttachmentFunction(embeddedComponent);
    },
    [reportRef, onReportLoadReportAttachmentFunction]
  );

  return (
    <PowerBIEmbed
      embedConfig={embedConfig}
      cssClassName={`${reportCssClassName ?? ""} report-embed`}
      getEmbeddedComponent={handleGetEmbeddedComponent}
      eventHandlers={eventHandlers}
    />
  );
};

export default EmbedReport;
