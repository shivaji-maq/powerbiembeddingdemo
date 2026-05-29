import { useEffect, useRef, useState, useCallback, useMemo, forwardRef, useImperativeHandle } from "react";
import { useDispatch } from "react-redux";
import { models } from "powerbi-client";
import "powerbi-report-authoring";
import { PowerBIEmbed } from "powerbi-client-react";
import VisualCreatorModal from "./VisualCreatorModal";
import type { VisualCreatorState } from "./VisualCreatorModal";
import { addBookmark } from "../../redux/slices/bookmarkSlice/bookmarkSlice";
import type { BookmarkedVisual, VisualSnapshot } from "../../redux/slices/bookmarkSlice/@types";
import {
  schemas,
  propertyToSelector,
  visualTypeProperties,
  VISUAL_CREATOR_SHOWCASE,
} from "./visualConfig";
import type { DatasetField } from "../../lib/powerbiLib";
import "./QuickVisualCreator.css";

interface QuickVisualCreatorProps {
  accessToken: string;
  embedUrl: string;
  reportId: string;
  datasetId: string;
  workspaceId: string;
  tokenType?: "Aad" | "Embed";
}

interface BaseReportState {
  report: any;
  page: any;
  visuals: any[];
}

export default forwardRef(function QuickVisualCreator({ accessToken, embedUrl, reportId, tokenType = "Aad" }: QuickVisualCreatorProps, ref) {
  const dispatch = useDispatch();
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [datasetFields, setDatasetFields] = useState<DatasetField[]>([]);
  const [baseReportState, setBaseReportState] = useState<BaseReportState>({
    report: null,
    page: null,
    visuals: [],
  });
  const [, setAuthoringReport] = useState<any>(null);
  const [authoringPage, setAuthoringPage] = useState<any>(null);
  const [editingVisual, setEditingVisual] = useState<any>(null);
  const hasInitializedRef = useRef(false);

  const baseReportRef = useRef<any>(null);
  const authoringReportRef = useRef<any>(null);
  const baseStateRef = useRef(baseReportState);
  const handleLoadBookmarkRef = useRef<((bookmark: BookmarkedVisual) => void) | null>(null);
  const isLoadingBookmarkRef = useRef(false);
  const lastVisualStateRef = useRef<VisualCreatorState | null>(null);
  const allVisualSnapshotsRef = useRef<VisualSnapshot[]>([]);

  useEffect(() => {
    baseStateRef.current = baseReportState;
  }, [baseReportState]);

  // Immediately sync ref on state changes to avoid stale reads between renders
  const updateBaseState = useCallback((updater: (prev: BaseReportState) => BaseReportState) => {
    setBaseReportState((prev) => {
      const next = updater(prev);
      baseStateRef.current = next;
      return next;
    });
  }, []);

  // Expose openModal, loadBookmark, bookmarkLastVisual methods to parent via ref
  useImperativeHandle(ref, () => ({
    openModal: () => setIsModalOpen(true),
    loadBookmark: (bookmark: BookmarkedVisual) => handleLoadBookmarkRef.current?.(bookmark),
    bookmarkLastVisual: (name: string) => {
      const snapshots = allVisualSnapshotsRef.current;
      if (!snapshots.length) return false;
      const label = name || `View – ${new Date().toLocaleTimeString()}`;
      dispatch(addBookmark({
        id: crypto.randomUUID(),
        name: label,
        visuals: [...snapshots],
        createdAt: Date.now(),
      }));
      return true;
    },
    hasVisual: () => allVisualSnapshotsRef.current.length > 0,
  }), [dispatch]);

  // Base report configuration - memoized to prevent re-embed on every render
  const baseReportConfig = useMemo<models.IReportEmbedConfiguration>(() => ({
    type: "report",
    id: reportId,
    embedUrl: embedUrl,
    accessToken: accessToken,
    tokenType: tokenType === "Embed" ? models.TokenType.Embed : models.TokenType.Aad,
    permissions: models.Permissions.All,
    settings: {
      panes: {
        filters: { visible: false },
        pageNavigation: { visible: false },
      },
      extensions: [
        {
          command: {
            name: "changeVisual",
            title: "Change visual",
            extend: {
              visualOptionsMenu: {
                title: "Change visual",
                menuLocation: models.MenuLocation.Top,
              },
            },
          },
        },
      ],
      background: models.BackgroundType.Transparent,
    },
  }), [reportId, embedUrl, accessToken, tokenType]);

  // Authoring report configuration - memoized to prevent re-embed on every render
  const authoringReportConfig = useMemo<models.IReportEmbedConfiguration>(() => ({
    type: "report",
    id: reportId,
    embedUrl: embedUrl,
    accessToken: accessToken,
    tokenType: tokenType === "Embed" ? models.TokenType.Embed : models.TokenType.Aad,
    permissions: models.Permissions.All,
    settings: {
      panes: {
        filters: { visible: false },
        pageNavigation: { visible: false },
      },
      background: models.BackgroundType.Transparent,
    },
  }), [reportId, embedUrl, accessToken, tokenType]);

  // Rearrange visuals in a custom grid layout
  const rearrangeInCustomLayout = useCallback(
    async (report: any, page: any, visuals: any[]) => {
      if (!report || !visuals || visuals.length === 0) return;

      const containerWidth = 1200;
      const margin = VISUAL_CREATOR_SHOWCASE.MARGIN;
      const columns = VISUAL_CREATOR_SHOWCASE.COLUMNS;
      const aspectRatio = VISUAL_CREATOR_SHOWCASE.VISUAL_ASPECT_RATIO;

      const visualsTotalWidth = containerWidth - margin * (columns + 1);
      const visualWidth = visualsTotalWidth / columns;
      const visualHeight = visualWidth * aspectRatio;

      let x = margin;
      let y = margin;

      const visualsLayout: Record<string, any> = {};

      visuals.forEach((visual) => {
        visualsLayout[visual.name] = {
          x,
          y,
          width: visualWidth,
          height: visualHeight,
          displayState: {
            mode: models.VisualContainerDisplayMode.Visible,
          },
        };

        x += visualWidth + margin;
        if (x + visualWidth > containerWidth) {
          x = margin;
          y += visualHeight + margin;
        }
      });

      const rows = Math.ceil(visuals.length / columns);
      const reportHeight = rows * visualHeight + (rows + 1) * margin;

      const pagesLayout: Record<string, any> = {};
      pagesLayout[page.name] = {
        defaultLayout: {
          displayState: {
            mode: models.VisualContainerDisplayMode.Hidden,
          },
        },
        visualsLayout,
      };

      const settings = {
        background: models.BackgroundType.Transparent,
        layoutType: models.LayoutType.Custom,
        customLayout: {
          pageSize: {
            type: models.PageSizeType.Custom,
            width: containerWidth,
            height: reportHeight,
          },
          displayOption: models.DisplayOption.FitToWidth,
          pagesLayout,
        },
        commands: [
          {
            exportData: { displayOption: models.CommandDisplayOption.Hidden },
            drill: { displayOption: models.CommandDisplayOption.Hidden },
            spotlight: { displayOption: models.CommandDisplayOption.Hidden },
            sort: { displayOption: models.CommandDisplayOption.Hidden },
            seeData: { displayOption: models.CommandDisplayOption.Hidden },
          },
        ],
      };

      try {
        await report.updateSettings(settings);
      } catch (e) {
        console.error("Error updating layout:", e);
      }
    },
    []
  );

  // Handle base report rendered (report is fully ready after embed)
  const handleBaseReportRendered = useCallback(
    async (report: any) => {
      if (hasInitializedRef.current) return;
      hasInitializedRef.current = true;

      try {
        const pages = await report.getPages();
        const activePage = pages[0];
        await activePage.setActive();

        const visuals = await activePage.getVisuals();

        updateBaseState(() => ({ report, page: activePage, visuals }));
        baseReportRef.current = report;

        await rearrangeInCustomLayout(report, activePage, visuals);
        setIsLoading(false);

        extractFieldsFromVisuals(visuals);
      } catch (e) {
        console.error("Error on base report rendered:", e);
        setIsLoading(false);
      }
    },
    [rearrangeInCustomLayout, updateBaseState]
  );

  // Fallback: extract fields from existing visuals in the report
  const extractFieldsFromVisuals = useCallback(async (visuals: any[]) => {
    const discoveredFields: DatasetField[] = [];
    const seen = new Set<string>();

    for (const visual of visuals) {
      try {
        const capabilities = await visual.getCapabilities();
        if (!capabilities?.dataRoles) continue;

        for (const dr of capabilities.dataRoles) {
          try {
            const fields = await visual.getDataFields(dr.name);
            if (!fields) continue;
            for (const field of fields) {
              const table = field.table || "";
              const name = field.column || field.measure || "";
              const type = field.measure ? "measure" : "column";
              const key = `${table}||${name}||${type}`;
              if (name && !seen.has(key)) {
                seen.add(key);
                discoveredFields.push({ table, name, type });
              }
            }
          } catch { /* skip */ }
        }
      } catch { /* skip visual */ }
    }

    if (discoveredFields.length > 0) {
      setDatasetFields(discoveredFields);
    }
  }, []);

  // Handle authoring report loaded
  const handleAuthoringReportLoaded = useCallback(async (report: any) => {
    try {
      const pages = await report.getPages();
      const authoringPg = pages.length > 1 ? pages[1] : pages[0];
      await authoringPg.setActive();
      setAuthoringReport(report);
      setAuthoringPage(authoringPg);
      authoringReportRef.current = report;
    } catch (e) {
      console.error("Error loading authoring report:", e);
    }
  }, []);

  // Append created visual to the base report
  const handleCreateVisual = useCallback(
    async (visualState: VisualCreatorState) => {
      const currentState = baseStateRef.current;
      if (!currentState.report || !currentState.page || !visualState.visualType) return;
      lastVisualStateRef.current = visualState;

      try {
        const margin = VISUAL_CREATOR_SHOWCASE.MARGIN;
        const columns = VISUAL_CREATOR_SHOWCASE.COLUMNS;
        const containerWidth = 1200;
        const visualsTotalWidth = containerWidth - margin * (columns + 1);
        const visualWidth = visualsTotalWidth / columns;
        const visualHeight = visualWidth * VISUAL_CREATOR_SHOWCASE.VISUAL_ASPECT_RATIO;

        const existingCount = currentState.visuals.length;
        const row = Math.floor(existingCount / columns);
        const col = existingCount % columns;
        const x = margin + col * (visualWidth + margin);
        const y = margin + row * (visualHeight + margin);

        const layout = {
          x,
          y,
          width: visualWidth,
          height: visualHeight,
          displayState: { mode: models.VisualContainerDisplayMode.Visible },
        };

        const visualResponse = await currentState.page.createVisual(visualState.visualType, layout);
        const visual = visualResponse.visual || visualResponse;

        visual.setProperty(propertyToSelector("titleSize"), { schema: schemas.property, value: 13 });
        visual.setProperty(propertyToSelector("titleColor"), { schema: schemas.property, value: "#000" });

        if (visualState.visualType === "pieChart") {
          visual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: true });
        }

        const availableProps = visualTypeProperties[visualState.visualType] || [];
        Object.entries(visualState.properties).forEach(([propName, propValue]) => {
          if (propName === "titleText") {
            if (propValue && typeof propValue === "string" && propValue !== "") {
              visual.setProperty(propertyToSelector("titleText"), { schema: schemas.property, value: propValue });
            }
            return;
          }
          if (propName === "titleAlign") {
            if (propValue && typeof propValue === "string") {
              visual.setProperty(propertyToSelector("titleAlign"), { schema: schemas.property, value: propValue });
            }
            return;
          }
          if (availableProps.includes(propName) || propName === "title") {
            visual.setProperty(propertyToSelector(propName), { schema: schemas.property, value: propValue });
          }
        });

        if (visualState.visualType === "columnChart" || visualState.visualType === "barChart") {
          visual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: false });
        }

        const validDataRoles = Object.entries(visualState.dataRoles).filter(([, field]) => field !== null);
        for (const [dataRole, field] of validDataRoles) {
          if (field) {
            try {
              await visual.addDataField(dataRole, field);
            } catch (e) {
              console.error(`Error adding data field to ${dataRole}:`, e);
            }
          }
        }

        const updatedVisuals = [...currentState.visuals, visual];
        updateBaseState((prev) => ({ ...prev, visuals: updatedVisuals }));

        await rearrangeInCustomLayout(currentState.report, currentState.page, updatedVisuals);

        // Track snapshot for this visual
        const snapshot: VisualSnapshot = {
          visualType: visualState.visualType,
          dataRoles: visualState.dataRoles,
          properties: visualState.properties,
        };
        allVisualSnapshotsRef.current = [...allVisualSnapshotsRef.current, snapshot];

        // Auto-save all current visuals as a bookmark
        const autoName = [
          visualState.properties.titleText || null,
          visualState.visualType,
        ].filter(Boolean).join(" – ");
        dispatch(addBookmark({
          id: crypto.randomUUID(),
          name: autoName,
          visuals: [...allVisualSnapshotsRef.current],
          createdAt: Date.now(),
        }));
      } catch (e) {
        console.error("Error creating visual on base report:", e);
      }
    },
    [rearrangeInCustomLayout, updateBaseState, dispatch]
  );


  // Load all visuals from a bookmarked view onto the report
  const handleLoadBookmark = useCallback(
    async (bookmark: BookmarkedVisual) => {
      if (isLoadingBookmarkRef.current) return;

      const currentState = baseStateRef.current;
      if (!currentState.report || !currentState.page) return;

      isLoadingBookmarkRef.current = true;
      try {
        const margin = VISUAL_CREATOR_SHOWCASE.MARGIN;
        const columns = VISUAL_CREATOR_SHOWCASE.COLUMNS;
        const containerWidth = 1200;
        const visualsTotalWidth = containerWidth - margin * (columns + 1);
        const visualWidth = visualsTotalWidth / columns;
        const visualHeight = visualWidth * VISUAL_CREATOR_SHOWCASE.VISUAL_ASPECT_RATIO;

        // Append to whatever is already on the report
        let runningVisuals: any[] = [...currentState.visuals];

        for (const snap of bookmark.visuals) {
          const existingCount = runningVisuals.length;
          const row = Math.floor(existingCount / columns);
          const col = existingCount % columns;
          const layout = {
            x: margin + col * (visualWidth + margin),
            y: margin + row * (visualHeight + margin),
            width: visualWidth,
            height: visualHeight,
            displayState: { mode: models.VisualContainerDisplayMode.Visible },
          };

          const visualResponse = await currentState.page.createVisual(snap.visualType, layout);
          const visual = visualResponse.visual || visualResponse;

          visual.setProperty(propertyToSelector("titleSize"), { schema: schemas.property, value: 13 });
          visual.setProperty(propertyToSelector("titleColor"), { schema: schemas.property, value: "#000" });

          if (snap.visualType === "pieChart") {
            visual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: true });
          }

          const availableProps = visualTypeProperties[snap.visualType] || [];
          Object.entries(snap.properties).forEach(([propName, propValue]) => {
            if (propName === "titleText") {
              if (propValue && typeof propValue === "string" && propValue !== "") {
                visual.setProperty(propertyToSelector("titleText"), { schema: schemas.property, value: propValue });
              }
              return;
            }
            if (propName === "titleAlign") {
              if (propValue && typeof propValue === "string") {
                visual.setProperty(propertyToSelector("titleAlign"), { schema: schemas.property, value: propValue });
              }
              return;
            }
            if (availableProps.includes(propName) || propName === "title") {
              visual.setProperty(propertyToSelector(propName), { schema: schemas.property, value: propValue });
            }
          });

          if (snap.visualType === "columnChart" || snap.visualType === "barChart") {
            visual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: false });
          }

          const validDataRoles = Object.entries(snap.dataRoles).filter(([, field]) => field !== null);
          for (const [dataRole, field] of validDataRoles) {
            if (field) {
              try {
                await visual.addDataField(dataRole, field);
              } catch (e) {
                console.error(`Error adding data field to ${dataRole}:`, e);
              }
            }
          }

          runningVisuals = [...runningVisuals, visual];
        }

        // Sync tracking refs — append new snapshots to existing ones
        allVisualSnapshotsRef.current = [...allVisualSnapshotsRef.current, ...bookmark.visuals];
        updateBaseState((prev) => ({ ...prev, visuals: runningVisuals }));
        await rearrangeInCustomLayout(currentState.report, currentState.page, runningVisuals);
      } catch (e) {
        console.error("Error loading bookmarked view:", e);
      } finally {
        isLoadingBookmarkRef.current = false;
      }
    },
    [rearrangeInCustomLayout, updateBaseState]
  );

  handleLoadBookmarkRef.current = handleLoadBookmark;

  // Handle context menu command to edit visual
  const handleCommandTriggered = useCallback((event: any) => {
    const visualData = event.detail;
    setEditingVisual(visualData?.visual || null);
    setIsModalOpen(true);
  }, []);

  return (
    <div className="qvc-container">
      {isLoading && (
        <div className="qvc-loading-overlay">
          <div className="qvc-spinner"></div>
          <p>Loading report...</p>
        </div>
      )}

      {/* Header - removed button since it's now in toggle row */}
      <div className="qvc-header"></div>

      {/* Base Report Display */}
      <div className="qvc-report-container">
        <PowerBIEmbed
          embedConfig={baseReportConfig}
          cssClassName="qvc-report-embed"
          getEmbeddedComponent={(component) => {
            baseReportRef.current = component;
          }}
          eventHandlers={
            new Map([
              [
                "rendered",
                () => {
                  if (baseReportRef.current) {
                    handleBaseReportRendered(baseReportRef.current);
                  }
                },
              ],
              ["commandTriggered", handleCommandTriggered],
              [
                "error",
                (event: any) => {
                  console.error("Base report error:", event?.detail);
                },
              ],
            ])
          }
        />
      </div>

      {/* Hidden Authoring Report (for visual preview in modal) */}
      <div className="qvc-authoring-container">
        <PowerBIEmbed
          embedConfig={authoringReportConfig}
          cssClassName="qvc-authoring-embed"
          getEmbeddedComponent={(component) => {
            authoringReportRef.current = component;
          }}
          eventHandlers={
            new Map([
              [
                "loaded",
                () => {
                  if (authoringReportRef.current) {
                    handleAuthoringReportLoaded(authoringReportRef.current);
                  }
                },
              ],
              [
                "error",
                (event: any) => {
                  console.error("Authoring report error:", event?.detail);
                },
              ],
            ])
          }
        />
      </div>

      {/* Visual Creator Modal */}
      <VisualCreatorModal
        isOpen={isModalOpen}
        onClose={() => {
          setIsModalOpen(false);
          setEditingVisual(null);
        }}
        onCreateVisual={handleCreateVisual}
        report={baseReportState.report}
        authoringPage={authoringPage}
        editingVisual={editingVisual}
        datasetFields={datasetFields}
      />
    </div>
  );
});
