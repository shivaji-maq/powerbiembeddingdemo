/* eslint-disable @typescript-eslint/no-explicit-any */
import { useState, useCallback, useEffect, forwardRef, useImperativeHandle } from "react";
import { models } from "powerbi-client";
import "./LayoutCustomizer.css";

// Layout types matching the showcase
const SPAN_TYPE = { NONE: 0, ROWSPAN: 1, COLSPAN: 2 } as const;

const LAYOUT_SHOWCASE = {
  MARGIN: 16,
  VISUAL_ASPECT_RATIO: 9 / 16,
};

interface VisualItem {
  name: string;
  title: string;
  checked: boolean;
}

export interface LayoutState {
  selectedVisuals: string[];
  columns: number;
  spanType: number;
  isCustomLayoutActive: boolean;
  pageName?: string;
}

interface LayoutCustomizerProps {
  report: any;
  page: any;
}

export interface LayoutCustomizerHandle {
  getLayoutState: () => LayoutState;
  setLayoutState: (state: LayoutState) => void;
  resetToDefault: () => void;
}

const LayoutCustomizer = forwardRef<LayoutCustomizerHandle, LayoutCustomizerProps>(function LayoutCustomizer({ report, page }, ref) {
  const [visuals, setVisuals] = useState<VisualItem[]>([]);
  const [columns, setColumns] = useState(2);
  const [spanType, setSpanType] = useState<number>(SPAN_TYPE.NONE);
  const [isCustomLayoutActive, setIsCustomLayoutActive] = useState(false);
  const [showVisuals, setShowVisuals] = useState(false);
  const [showLayouts, setShowLayouts] = useState(false);
  const [pendingLayoutState, setPendingLayoutState] = useState<LayoutState | null>(null);
  const [layoutRenderTrigger, setLayoutRenderTrigger] = useState(0); // Trigger for forced layout render
  const [currentPageName, setCurrentPageName] = useState<string | undefined>(undefined); // Track current page for comparison

  const restoreDefaultLayout = useCallback(async () => {
    if (!report) return;

    try {
      // Reload the report to fully restore its original layout, page size, and all visuals
      await report.reload();
    } catch (e) {
      console.error("Error restoring default layout:", e);
    }
  }, [report]);

  // Expose layout state via ref for bookmark save/load
  useImperativeHandle(
    ref,
    () => ({
      getLayoutState: (): LayoutState => ({
        selectedVisuals: visuals.filter((v) => v.checked).map((v) => v.name),
        columns,
        spanType,
        isCustomLayoutActive,
        pageName: page?.name,
      }),
      setLayoutState: (state: LayoutState) => {
        if (!state) return;
        // Set pending state which will trigger effect to apply layout
        setPendingLayoutState(state);
      },
      resetToDefault: () => {
        setColumns(2);
        setSpanType(SPAN_TYPE.NONE);
        setIsCustomLayoutActive(false);
        setVisuals((prev) => prev.map((v) => ({ ...v, checked: true })));
        void restoreDefaultLayout();
      },
    }),
    [visuals, columns, spanType, isCustomLayoutActive, page?.name, restoreDefaultLayout],
  );

  const fetchPageVisuals = useCallback(async () => {
    if (!page || typeof page.getVisuals !== "function") {
      return [] as VisualItem[];
    }

    try {
      const pageVisuals = await page.getVisuals();
      const items: VisualItem[] = pageVisuals
        .filter((v: any) => v.name && v.title !== undefined && v.title !== "")
        .map((v: any) => ({ name: v.name, title: v.title, checked: true }));
      setVisuals(items);
      return items;
    } catch (e) {
      console.error("Error fetching visuals:", e);
      setVisuals([]);
      return [] as VisualItem[];
    }
  }, [page]);

  // Fetch visuals from the active page
  useEffect(() => {
    if (!page) return;
    setShowVisuals(false);
    setShowLayouts(false);
    
    // Only reset layout if it's a DIFFERENT page (by page name)
    // If the page name is the same but the reference changed (e.g., after theme application),
    // preserve the current layout state
    if (page.name !== currentPageName) {
      setIsCustomLayoutActive(false);
    }
    
    // Update current page name
    setCurrentPageName(page.name);
    
    void fetchPageVisuals();
  }, [fetchPageVisuals, page]);

  useEffect(() => {
    if (!pendingLayoutState) return;

    const savedPageName = pendingLayoutState.pageName;
    if (savedPageName && page?.name && savedPageName !== page.name) {
      setPendingLayoutState(null);
      return;
    }

    // Apply layout state from bookmark
    const nextColumns = pendingLayoutState.columns ?? 2;
    const nextSpanType = pendingLayoutState.spanType ?? SPAN_TYPE.NONE;
    const nextIsCustomLayoutActive = !!pendingLayoutState.isCustomLayoutActive;

    setColumns(nextColumns);
    setSpanType(nextSpanType);
    setIsCustomLayoutActive(nextIsCustomLayoutActive);

    // Update visual selections if available
    let visualsApplied = false;
    if (Array.isArray(pendingLayoutState.selectedVisuals) && visuals.length > 0) {
      const selectedNames = new Set(pendingLayoutState.selectedVisuals);
      setVisuals((prev) =>
        prev.map((v) => ({
          ...v,
          checked: selectedNames.has(v.name),
        })),
      );
      visualsApplied = true;
    }

    // If not custom layout, restore default
    if (!nextIsCustomLayoutActive) {
      void restoreDefaultLayout();
      setPendingLayoutState(null); // Clear after restoring
      setLayoutRenderTrigger((prev) => prev + 1); // Trigger render
    } else if (visualsApplied || !Array.isArray(pendingLayoutState.selectedVisuals)) {
      // Only clear pending state if we applied visuals or there are no specific visuals to select
      setPendingLayoutState(null);
      setLayoutRenderTrigger((prev) => prev + 1); // Trigger render
    }
    // If visuals weren't applied yet, keep pending state and wait for visuals to load
  }, [page?.name, pendingLayoutState, restoreDefaultLayout, visuals.length]);

  // Render visuals with current layout
  const renderVisuals = useCallback(async () => {
    if (!report || !visuals.length || !page) return;
    if (!isCustomLayoutActive) return;

    const containerWidth = 1200;
    let reportHeight = 0;

    const checkedVisuals = visuals.filter((v) => v.checked);
    if (checkedVisuals.length === 0) {
      const pagesLayout: Record<string, any> = {};
      pagesLayout[page.name] = {
        defaultLayout: {
          displayState: { mode: models.VisualContainerDisplayMode.Hidden },
        },
        visualsLayout: {},
      };

      try {
        await report.updateSettings({
          background: models.BackgroundType.Transparent,
          layoutType: models.LayoutType.Custom,
          customLayout: {
            pageSize: {
              type: models.PageSizeType.Custom,
              width: containerWidth,
              height: 600,
            },
            displayOption: models.DisplayOption.ActualSize,
            pagesLayout,
          },
        });
      } catch (e) {
        console.error("Error updating layout with no selected visuals:", e);
      }
      return;
    }

    const visualsTotalWidth = containerWidth - LAYOUT_SHOWCASE.MARGIN * (columns + 1) + LAYOUT_SHOWCASE.MARGIN / 2;
    const visualWidth = visualsTotalWidth / columns;
    const visualHeight = visualWidth * LAYOUT_SHOWCASE.VISUAL_ASPECT_RATIO;

    const visualsLayout: Record<string, any> = {};
    let x = LAYOUT_SHOWCASE.MARGIN;
    let y = LAYOUT_SHOWCASE.MARGIN;

    if (spanType === SPAN_TYPE.COLSPAN) {
      const rowsPerSection = 2;
      const visualsPerSection = 3;
      let rows = rowsPerSection * Math.floor(checkedVisuals.length / visualsPerSection);
      if (checkedVisuals.length % visualsPerSection) rows += 1;
      reportHeight = Math.max(reportHeight, rows * visualHeight + (rows + 1) * LAYOUT_SHOWCASE.MARGIN);

      checkedVisuals.forEach((element, idx) => {
        visualsLayout[element.name] = {
          x,
          y,
          width: idx % visualsPerSection === visualsPerSection - 1 ? visualWidth * 2 + LAYOUT_SHOWCASE.MARGIN : visualWidth,
          height: visualHeight,
          displayState: { mode: models.VisualContainerDisplayMode.Visible },
        };
        x += LAYOUT_SHOWCASE.MARGIN + (idx % visualsPerSection === visualsPerSection - 1 ? visualWidth * 2 : visualWidth);
        if (x + visualWidth > containerWidth) {
          x = LAYOUT_SHOWCASE.MARGIN;
          y += visualHeight + LAYOUT_SHOWCASE.MARGIN;
        }
      });
    } else if (spanType === SPAN_TYPE.ROWSPAN) {
      const rowsPerSection = 2;
      const visualsPerSection = 3;
      let rows = rowsPerSection * Math.floor(checkedVisuals.length / visualsPerSection);
      if (checkedVisuals.length % visualsPerSection) rows += 2;
      reportHeight = Math.max(reportHeight, rows * visualHeight + (rows + 1) * LAYOUT_SHOWCASE.MARGIN);

      checkedVisuals.forEach((element, idx) => {
        visualsLayout[element.name] = {
          x,
          y,
          width: visualWidth,
          height: !(idx % visualsPerSection) ? visualHeight * 2 + LAYOUT_SHOWCASE.MARGIN : visualHeight,
          displayState: { mode: models.VisualContainerDisplayMode.Visible },
        };
        x += visualWidth + LAYOUT_SHOWCASE.MARGIN;
        if (x + visualWidth > containerWidth) {
          x = (idx + 1) % visualsPerSection === 0 ? LAYOUT_SHOWCASE.MARGIN : 2 * LAYOUT_SHOWCASE.MARGIN + visualWidth;
          y += idx % visualsPerSection === 0 ? visualHeight * 2 + LAYOUT_SHOWCASE.MARGIN : visualHeight + LAYOUT_SHOWCASE.MARGIN;
        }
      });
    } else {
      // SPAN_TYPE.NONE
      let adjustedHeight = visualHeight;
      if (columns === 1) adjustedHeight = visualHeight / 2;

      const rows = Math.ceil(checkedVisuals.length / columns);
      reportHeight = Math.max(reportHeight, rows * adjustedHeight + (rows + 1) * LAYOUT_SHOWCASE.MARGIN);

      checkedVisuals.forEach((element) => {
        visualsLayout[element.name] = {
          x,
          y,
          width: visualWidth,
          height: adjustedHeight,
          displayState: { mode: models.VisualContainerDisplayMode.Visible },
        };
        x += visualWidth + LAYOUT_SHOWCASE.MARGIN;
        if (x + visualWidth > containerWidth) {
          x = LAYOUT_SHOWCASE.MARGIN;
          y += adjustedHeight + LAYOUT_SHOWCASE.MARGIN;
        }
      });
    }

    const pagesLayout: Record<string, any> = {};
    pagesLayout[page.name] = {
      defaultLayout: {
        displayState: { mode: models.VisualContainerDisplayMode.Hidden },
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
        displayOption: models.DisplayOption.ActualSize,
        pagesLayout,
      },
    };

    try {
      await report.updateSettings(settings);
    } catch (e) {
      console.error("Error updating layout:", e);
    }
  }, [report, page, visuals, columns, spanType, isCustomLayoutActive]);

  // Handle case where visuals load after pending layout state is set
  useEffect(() => {
    if (pendingLayoutState && visuals.length > 0 && pendingLayoutState.isCustomLayoutActive) {
      // Visuals have loaded while we have pending layout, trigger re-application
      setLayoutRenderTrigger((prev) => prev + 1);
    }
  }, [visuals.length, pendingLayoutState?.isCustomLayoutActive]);

  // Re-render when layout changes (directly watch layout state + explicit trigger)
  useEffect(() => {
    renderVisuals();
  }, [columns, spanType, isCustomLayoutActive, visuals, page, report, layoutRenderTrigger, renderVisuals]);

  const toggleVisual = (name: string) => {
    setVisuals((prev) => prev.map((v) => (v.name === name ? { ...v, checked: !v.checked } : v)));
  };

  const selectLayout = (cols: number, span: number) => {
    setColumns(cols);
    setSpanType(span);
    setIsCustomLayoutActive(true);
    setShowLayouts(false);
  };

  const chooseDefaultLayout = async () => {
    setIsCustomLayoutActive(false);
    setShowLayouts(false);
    setVisuals((prev) => prev.map((v) => ({ ...v, checked: true })));
    await restoreDefaultLayout();
    // Re-fetch visuals after reload to ensure state is fresh
    await fetchPageVisuals();
  };

  const layoutLabel = () => {
    if (!isCustomLayoutActive) return "Power BI default";
    if (spanType === SPAN_TYPE.COLSPAN) return "2 Col (Colspan)";
    if (spanType === SPAN_TYPE.ROWSPAN) return "2 Col (Rowspan)";
    if (columns === 1) return "1 Column";
    if (columns === 3) return "3 Columns";
    return "2 Columns";
  };

  return (
    <div className="layout-customizer">
      {/* Choose Visuals dropdown */}
      <div className="lc-dropdown-wrapper">
        <button
          className="lc-btn"
          onClick={() => {
            setShowVisuals(!showVisuals);
            setShowLayouts(false);
          }}
        >
          Choose Visuals ▾
        </button>
        {showVisuals && (
          <div className="lc-dropdown">
            {visuals.length === 0 ? (
              <div className="lc-dropdown-empty">No visuals found</div>
            ) : (
              visuals.map((v) => (
                <label key={v.name} className="lc-checkbox-item">
                  <input type="checkbox" checked={v.checked} onChange={() => toggleVisual(v.name)} />
                  <span className="lc-visual-title">{v.title}</span>
                </label>
              ))
            )}
          </div>
        )}
      </div>

      {/* Choose Layout dropdown */}
      <div className="lc-dropdown-wrapper">
        <button
          className="lc-btn"
          onClick={() => {
            setShowLayouts(!showLayouts);
            setShowVisuals(false);
          }}
        >
          Layout: {layoutLabel()} ▾
        </button>
        {showLayouts && (
          <div className="lc-dropdown lc-layout-dropdown">
            <button className="lc-layout-option" onClick={chooseDefaultLayout}>
              <span className="lc-layout-icon">▣</span> Power BI default
            </button>
            <button className="lc-layout-option" onClick={() => selectLayout(1, SPAN_TYPE.NONE)}>
              <span className="lc-layout-icon">▮</span> 1 Column
            </button>
            <button className="lc-layout-option" onClick={() => selectLayout(2, SPAN_TYPE.NONE)}>
              <span className="lc-layout-icon">▮▮</span> 2 Columns
            </button>
            <button className="lc-layout-option" onClick={() => selectLayout(2, SPAN_TYPE.COLSPAN)}>
              <span className="lc-layout-icon">▮▯</span> 2 Col (Colspan)
            </button>
            <button className="lc-layout-option" onClick={() => selectLayout(2, SPAN_TYPE.ROWSPAN)}>
              <span className="lc-layout-icon">▯▮</span> 2 Col (Rowspan)
            </button>
            <button className="lc-layout-option" onClick={() => selectLayout(3, SPAN_TYPE.NONE)}>
              <span className="lc-layout-icon">▮▮▮</span> 3 Columns
            </button>
          </div>
        )}
      </div>
    </div>
  );
});

export default LayoutCustomizer;
