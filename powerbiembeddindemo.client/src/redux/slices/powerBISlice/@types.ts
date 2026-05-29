import type { Report, SemanticModel, WorkSpace } from "../../../lib/powerbiLib/@types";

export interface PowerBISliceState {
    fetchingWorkspaces: boolean;
    workspaces: WorkSpace[];
    selectedWorkSpace: WorkSpace | null;
    errorInWorkspace: string | null;
    fetchingSemanticModels: boolean;
    semanticModels: SemanticModel[];
    selectedSemanticModel: SemanticModel[];
    errorInSemanticModel: string | null;
    fetchingReports: boolean;
    reports: Report[];
    selectedReports: Report[];
    errorInReport: string | null;
}