// lib/powerbiLib/personalization.ts
import { models } from "powerbi-client";

/**
 * Apply personalized filters to a report
 */
export const applyPersonalizedFilters = async (
  embeddedReport: any,
  filters: models.IFilter[]
) => {
  try {
    await embeddedReport.setFilters(filters);
    console.log("Personalized filters applied successfully");
  } catch (error) {
    console.error("Error applying personalized filters:", error);
    throw error;
  }
};

/**
 * Get current report state for personalization
 */
export const getReportPersonalizationState = async (
  embeddedReport: any
): Promise<{
  filters: models.IFilter[];
  bookmarks: any[];
  pages: any[];
  currentPage: string;
}> => {
  try {
    const filters = await embeddedReport.getFilters();
    const bookmarks = await embeddedReport.bookmarksManager.getBookmarks();
    const pages = await embeddedReport.getPages();
    const currentPage = (await embeddedReport.getActivePage()).name;

    return {
      filters,
      bookmarks,
      pages,
      currentPage,
    };
  } catch (error) {
    console.error("Error getting personalization state:", error);
    throw error;
  }
};

/**
 * Clear all filters from report
 */
export const clearAllFilters = async (embeddedReport: any) => {
  try {
    await embeddedReport.setFilters([]);
    console.log("All filters cleared");
  } catch (error) {
    console.error("Error clearing filters:", error);
    throw error;
  }
};

/**
 * Apply specific filter to a table column
 */
export const applyColumnFilter = async (
  embeddedReport: any,
  tableName: string,
  columnName: string,
  values: any[]
) => {
  const filter: models.IFilter = {
    $schema: "http://powerbi.com/product/schema#basic",
    target: {
      table: tableName,
      column: columnName,
    } as any,
    operator: "In",
    values,
  };

  await applyPersonalizedFilters(embeddedReport, [filter]);
};

/**
 * Remove specific filter
 */
export const removeFilter = async (
  embeddedReport: any,
  tableName: string,
  columnName: string
) => {
  try {
    const currentFilters = await embeddedReport.getFilters();
    const updatedFilters = currentFilters.filter(
      (f: any) =>
        !(f.target.table === tableName && f.target.column === columnName)
    );
    await embeddedReport.setFilters(updatedFilters);
    console.log(`Filter removed: ${tableName}.${columnName}`);
  } catch (error) {
    console.error("Error removing filter:", error);
    throw error;
  }
};

/**
 * Get visual details and calculations
 */
export const getVisualCalculations = async (
  embeddedReport: any,
  pageName: string
): Promise<
  Array<{
    id: string;
    name: string;
    type: string;
    dataPoints?: number;
  }>
> => {
  try {
    if (!embeddedReport || typeof embeddedReport.getPages !== "function") {
      console.warn("Embedded report instance does not support page enumeration");
      return [];
    }

    const pages = await embeddedReport.getPages();
    const page = pages.find((p: any) => p?.name === pageName);
    if (!page) {
      console.error(`Page ${pageName} not found`);
      return [];
    }

    if (typeof page.getVisuals !== "function") {
      console.warn(
        `Visual metadata is unavailable for page ${pageName} in this embed context`
      );
      return [];
    }

    const visuals = await page.getVisuals();

    const visualsWithCalculations = visuals.map((visual: any) => ({
      id: visual.name,
      name: visual.title || visual.name,
      type: visual.type,
      dataPoints: visual.dataPoints?.length || 0,
    }));

    return visualsWithCalculations;
  } catch (error) {
    console.error("Error getting visual calculations:", error);
    return [];
  }
};

/**
 * Switch report between view and edit modes
 */
export const setReportMode = async (
  embeddedReport: any,
  mode: "view" | "edit"
) => {
  try {
    if (!embeddedReport || typeof embeddedReport.switchMode !== "function") {
      throw new Error("Embedded report is not ready to switch modes");
    }

    if (mode === "edit") {
      await embeddedReport.switchMode(models.ViewMode.Edit);
      console.log("Report switched to edit mode");
    } else {
      await embeddedReport.switchMode(models.ViewMode.View);
      console.log("Report switched to view mode");
    }
  } catch (error) {
    console.error(`Error switching to ${mode} mode:`, error);
    const rawMessage =
      (error as any)?.detailedMessage ||
      (error as any)?.message ||
      "Failed to switch report mode";

    if (
      mode === "edit" &&
      String(rawMessage).toLowerCase().includes("insufficientpermissions")
    ) {
      throw new Error(
        "Edit mode is not allowed for this report. Ensure your account has Member/Contributor access to the workspace and report edit permissions."
      );
    }

    throw new Error(String(rawMessage));
  }
};

/**
 * Save report changes
 */
export const saveReportChanges = async (embeddedReport: any) => {
  try {
    const reportState = {
      filters: await embeddedReport.getFilters(),
      bookmarks: await embeddedReport.bookmarksManager.getBookmarks(),
    };
    return reportState;
  } catch (error) {
    console.error("Error saving report changes:", error);
    throw error;
  }
};

/**
 * Export report state as JSON
 */
export const exportReportState = async (
  embeddedReport: any
): Promise<string> => {
  try {
    const state = await getReportPersonalizationState(embeddedReport);
    return JSON.stringify(state, null, 2);
  } catch (error) {
    console.error("Error exporting report state:", error);
    throw error;
  }
};

/**
 * Import and apply report state from JSON
 */
export const importReportState = async (
  embeddedReport: any,
  stateJson: string
) => {
  try {
    const state = JSON.parse(stateJson);
    if (state.filters) {
      await applyPersonalizedFilters(embeddedReport, state.filters);
    }
    console.log("Report state imported successfully");
  } catch (error) {
    console.error("Error importing report state:", error);
    throw error;
  }
};
