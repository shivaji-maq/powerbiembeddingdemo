/* eslint-disable @typescript-eslint/no-explicit-any */
import { describe, it, expect, beforeEach, vi } from "vitest";
import reducer, { setSelectedWorkspace, setSelectedSemanticModel, setSelectedReports } from "../powerBISlice";

// Mock all external dependencies (mock the actual module path used in your slice)
vi.mock("../../../configs/msalInstance", () => ({
    getAccessToken: vi.fn().mockResolvedValue("mock-token"),
}));

// Instead of mocking a missing lib, mock the actual functions used in the slice directly on the module
vi.mock("../powerBISlice", async () => {
    const actual = await vi.importActual<any>("../powerBISlice");
    return {
        ...actual,
        // These are the async functions used in the slice
        fetchUserWorkspaces: vi.fn().mockResolvedValue([{ id: "ws1", name: "Workspace 1" }]),
        fetchSemanticModelsOfWorkspace: vi.fn().mockResolvedValue([{ id: "ds1", name: "SemanticModel 1" }]),
        fetchReportsForWorkspace: vi.fn().mockResolvedValue([{ id: "r1", name: "Report 1", semanticModelId: "ds1", type: "Report" }]),
    };
});

describe("powerBISlice reducers", () => {
    let initialState: any;
    beforeEach(() => {
        initialState = {
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
        vi.clearAllMocks();
    });

    it("setSelectedWorkspace sets selected workspace", () => {
        const ws = { id: "ws1", name: "Workspace 1", isOnDedicatedCapacity: false, isReadOnly: false };
        const state = reducer(initialState, setSelectedWorkspace({ workspace: ws }));
        expect(state.selectedWorkSpace).toEqual(ws);
    });

    it("setSelectedSemanticModel sets selected semanticModels", () => {
        const ds = [{ id: "ds1", name: "SemanticModel 1", workspaceId: "ws1" }];
        const state = reducer(initialState, setSelectedSemanticModel({ semanticModels: ds }));
        expect(state.selectedSemanticModel).toEqual(ds);
    });

    it("setSelectedReports sets selected reports", () => {
        const reports = [{ id: "r1", name: "Report 1", datasetId: "ds1", workspaceId: "ws1", embedUrl: "emurl" }];
        const state = reducer(initialState, setSelectedReports({ reports }));
        expect(state.selectedReports).toEqual(reports);
    });
});
