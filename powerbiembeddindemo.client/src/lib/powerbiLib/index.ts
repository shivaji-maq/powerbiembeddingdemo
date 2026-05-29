/* eslint-disable @typescript-eslint/no-explicit-any */
import axios from "axios";
import type { WorkSpace, SemanticModel, Report } from "./@types";

export interface DatasetField {
  table: string;
  name: string;
  type: "column" | "measure";
}

export const POWER_BI_API_CONST = {
  GROUP_BASE_URL: "https://api.powerbi.com/v1.0/myorg/groups",
  FABRIC_BASE_URL: "https://api.fabric.microsoft.com/v1",
};

const powerbiGroupBaseUrl = POWER_BI_API_CONST.GROUP_BASE_URL;

/**
 * Fetches the workspaces accessible to the user.
 *
 * @param {string} token - The access token for Power BI API.
 * @returns {Promise<WorkSpace[]>} A promise that resolves to an array of workspaces.
 */
export const fetchUserWorkspaces = async (token: string): Promise<WorkSpace[]> => {
  try {
    const response = await axios.get(powerbiGroupBaseUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
        "content-type": "application/json",
      },
    });

    const data = response.data?.value || [];
    console.log("[fetchUserWorkspaces] Raw API response:", data);

    // Map the API response and ensure isReadOnly is set correctly
    const workspaces: WorkSpace[] = data.map((ws: any) => {
      // Check if workspace has edit permissions based on type and other properties
      // Default isReadOnly to false unless specifically marked as read-only
      const isReadOnly = ws.type === "Personal" || ws.isReadOnly === true || false;

      return {
        id: ws.id,
        name: ws.name,
        isOnDedicatedCapacity: ws.isOnDedicatedCapacity || false,
        isReadOnly: isReadOnly,
      } as WorkSpace;
    });

    console.log("[fetchUserWorkspaces] Mapped workspaces:", workspaces);
    return workspaces;
  } catch (error) {
    console.error("[fetchUserWorkspaces] Error fetching workspaces:", error);
    throw error;
  }
};
/**
 * Fetches the semanticModels in a Power BI workspace.
 *
 * @param {string} token - The access token for Power BI API.
 * @param {string} groupId - The workspace ID.
 * @returns {Promise<SemanticModel[]>} A promise that resolves to an array of semanticModels.
 */
export const fetchSemanticModelsOfWorkspace = async (token: string, groupId: string): Promise<SemanticModel[]> => {
  const endPoint = `${powerbiGroupBaseUrl}/${groupId}/datasets`;

  try {
    const response = await axios.get(endPoint, {
      // method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json; charset=utf-8",
      },
    });

    const data: SemanticModel[] = response.data?.value;
    const sortedData = data.sort((a: SemanticModel, b: SemanticModel) => a?.name.localeCompare(b?.name));
    return sortedData;
  } catch (error: unknown) {
    if (axios.isAxiosError(error)) {
      throw new Error(`Failed to fetch report data. Status: ${error.response?.status} , Message: ${error.message}`, { cause: error });
    } else {
      throw new Error("Failed to fetch report data." + (error as Error).message, { cause: error });
    }
  }
};
/**
 * Fetches the reports in a Power BI workspace.
 *
 * @param {string} token - The access token for Power BI API.
 * @param {string} workspaceId - The workspace ID.
 * @returns {Promise<Report[]>} A promise that resolves to an array of reports.
 */
export const fetchReportsForWorkspace = async (token: string, workspaceId: string): Promise<Report[]> => {
  const reportUrl = `${powerbiGroupBaseUrl}/${workspaceId}/reports`;

  try {
    const response = await axios.get(reportUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
    });

    const reportData: Report[] = response.data?.value;

    return reportData.sort((a: Report, b: Report) => a?.name.localeCompare(b?.name));
  } catch (error: unknown) {
    if (axios.isAxiosError(error)) {
      throw new Error(`Failed to fetch report data. Status: ${error.response?.status} , Message: ${error.message}`, { cause: error });
    } else {
      throw new Error("Failed to fetch report data." + (error as Error).message, { cause: error });
    }
  }
};
/**
 * Fetches the pages of a report in a Power BI workspace.
 *
 * @param {string} token - The access token for Power BI API.
 * @param {string} workspaceId - The workspace ID.
 * @param {string} reportId - The report ID.
 * @returns {Promise<any[]>} A promise that resolves to an array of report pages.
 */
export async function fetchReportPages(token: string, workspaceId: string, reportId: string) {
  const pageUrl = `${powerbiGroupBaseUrl}/${workspaceId}/reports/${reportId}/pages`;

  try {
    const response = await axios.get(pageUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
    });

    return response.data?.value || [];
  } catch (error) {
    console.error("Failed to fetch report pages:", error);
    return [];
  }
}
