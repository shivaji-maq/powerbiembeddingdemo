/**
 * Redux slice for managing Power BI workspaces, semanticModels, and reports.
 *
 * @module powerBISlice
 *
 * @remarks
 * - Handles fetching, selection, and error states for workspaces, semanticModels, and reports.
 * - Provides async thunks for fetching workspaces, semanticModels, and reports from Power BI.
 * - Includes actions to set selected workspace, semanticModels, and reports.
 */
import { createAsyncThunk, createSlice } from "@reduxjs/toolkit";
import { getAccessToken, getAppModeAccessKey } from "../../../configs/msalInstance";
import { type SemanticModel, type Report, type WorkSpace } from "../../../lib/powerbiLib/@types";
import { fetchUserWorkspaces, fetchSemanticModelsOfWorkspace, fetchReportsForWorkspace } from "../../../lib/powerbiLib";
import type { PowerBISliceState } from "./@types";
/***
 * @file This file contains the Redux slice for managing Power BI workspaces, semanticModels, and reports.
 */
/**
 * Async thunk to fetch all Power BI workspaces accessible to the user.
 *
 * @function
 * @async
 * @returns {Promise<{ workspaces: WorkSpace[] }>} The fetched workspaces.
 * @throws {string} If fetching fails.
 */
export const fetchWorkspaces = createAsyncThunk<{ workspaces: WorkSpace[] }, string, { rejectValue: string }>(
  "powerbi/workspaces",
  async (accessToken: string, { rejectWithValue }) => {
    try {
      if (!accessToken) return rejectWithValue("Unable to get access token. Please login again.");

      const data = await fetchUserWorkspaces(accessToken);

      return { workspaces: data };
    } catch (error) {
      if (error instanceof Error) {
        return rejectWithValue(error.message);
      }
      return rejectWithValue("An unknown error occurred during login");
    }
  }
);
/**
 * Async thunk to fetch semanticModels for a given workspace in Power BI.
 *
 * @function
 * @async
 * @param {string} workspaceId - The workspace ID.
 * @returns {Promise<{ semanticModels: SemanticModel[] }>} The fetched semanticModels.
 * @throws {string} If fetching fails.
 */
export const fetchSemanticModels = createAsyncThunk<{ semanticModels: SemanticModel[] }, string, { rejectValue: string }>(
  "powerbi/semanticModels",
  async (workspaceId: string, { rejectWithValue }) => {
    try {
      const accessToken = await getAccessToken();
      if (!accessToken) return rejectWithValue("Unable to get access token. Please login again.");

      const semanticModels = await fetchSemanticModelsOfWorkspace(accessToken, workspaceId);

      return {
        semanticModels: semanticModels.map((d) => {
          return { ...d, workspaceId };
        }),
      };
    } catch (error) {
      if (error instanceof Error) {
        return rejectWithValue(error.message);
      }
      return rejectWithValue("An unknown error occurred during login");
    }
  }
);
/**
 * Async thunk to fetch reports for a given workspace in Power BI.
 *
 * @function
 * @async
 * @param {string} workspaceId - The workspace ID.
 * @returns {Promise<{ report: Report[] }>} The fetched reports.
 * @throws {string} If fetching fails.
 */
export const fetchReports = createAsyncThunk<{ report: Report[] }, { workspaceId: string; accessToken: string }, { rejectValue: string }>(
  "powerbi/report",
  async ({ workspaceId, accessToken }, { rejectWithValue }) => {
    try {
      if (!accessToken) return rejectWithValue("Unable to get access token. Please login again.");

      const report = await fetchReportsForWorkspace(accessToken, workspaceId);

      return {
        report: report.map((r: Report) => {
          return { ...r, workspaceId: workspaceId };
        }),
      };
    } catch (error) {
      if (error instanceof Error) {
        return rejectWithValue(error.message);
      }
      return rejectWithValue("An unknown error occurred during login");
    }
  }
);

/**
 * The initial state for the Power BI slice.
 *
 * @type {PowerBISliceState}
 */
const initialState: PowerBISliceState = {
  fetchingWorkspaces: false,
  workspaces: [],
  selectedWorkSpace: null,
  errorInWorkspace: null,
  fetchingSemanticModels: false,
  semanticModels: [],
  selectedSemanticModel: [],
  errorInSemanticModel: null,
  fetchingReports: false,
  reports: [],
  selectedReports: [],
  errorInReport: null,
};
/**
 * The Power BI slice, including reducers for selection and extraReducers for async thunks.
 *
 * @type {import('@reduxjs/toolkit').Slice}
 */
const powerBISlice = createSlice({
  name: "powerBI",
  initialState,
  reducers: {
    setSelectedWorkspace: (state, action: { payload: { workspace: WorkSpace } }) => {
      state.selectedWorkSpace = action.payload.workspace;
    },
    setSelectedSemanticModel: (state, action: { payload: { semanticModels: SemanticModel[] } }) => {
      state.selectedSemanticModel = action.payload.semanticModels;
    },
    setSelectedReports: (state, action: { payload: { reports: Report[] } }) => {
      state.selectedReports = action.payload.reports;
    },
  },
  /**
   *
   * @param builder This function is used to define the extra reducers for the powerBISlice.
   * It handles the pending, fulfilled, and rejected states of the async thunks for fetching
   */
  extraReducers: (builder) => {
    // workspace
    builder.addCase(fetchWorkspaces.pending, (state) => {
      state.fetchingWorkspaces = true;
      state.errorInWorkspace = null;
    });
    builder.addCase(fetchWorkspaces.rejected, (state, action) => {
      state.fetchingWorkspaces = false;
      state.errorInWorkspace = typeof action.payload === "string" ? action.payload : "An unknown error occurred";
    });
    builder.addCase(fetchWorkspaces.fulfilled, (state, action) => {
      state.fetchingWorkspaces = false;
      state.workspaces = action.payload.workspaces;
      state.errorInWorkspace = null;
    });

    // semanticModel
    builder.addCase(fetchSemanticModels.pending, (state) => {
      state.fetchingSemanticModels = true;
      state.errorInSemanticModel = null;
    });
    builder.addCase(fetchSemanticModels.rejected, (state, action) => {
      state.fetchingSemanticModels = false;
      state.errorInSemanticModel = action.payload ?? "An unknown error occurred";
    });
    builder.addCase(fetchSemanticModels.fulfilled, (state, action) => {
      state.fetchingSemanticModels = false;
      state.semanticModels = action.payload.semanticModels;
      state.errorInSemanticModel = null;
    });

    // report
    builder.addCase(fetchReports.pending, (state) => {
      state.fetchingReports = true;
      state.errorInReport = null;
    });
    builder.addCase(fetchReports.rejected, (state, action) => {
      state.fetchingReports = false;
      state.errorInReport = action.payload ?? "An unknown error occurred";
    });
    builder.addCase(fetchReports.fulfilled, (state, action) => {
      state.fetchingReports = false;
      state.reports = action.payload.report;
      state.errorInReport = null;
    });
  },
});

export const { setSelectedWorkspace, setSelectedSemanticModel, setSelectedReports } = powerBISlice.actions;
export default powerBISlice.reducer;
