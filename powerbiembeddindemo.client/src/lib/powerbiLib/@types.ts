/**
 * Represents a Power BI workspace.
 *
 * @property id The unique identifier of the workspace.
 * @property name The display name of the workspace.
 * @property isOnDedicatedCapacity Indicates if the workspace is on dedicated capacity.
 * @property isReadOnly Indicates if the workspace is read-only.
 */
export interface WorkSpace {
    /** The unique identifier of the workspace. */
    id: string;
    /** The display name of the workspace. */
    name: string;
    /** Indicates if the workspace is on dedicated capacity. */
    isOnDedicatedCapacity: boolean;
    /** Indicates if the workspace is read-only. */
    isReadOnly: boolean;
}

/**
 * Represents a Power BI semantic model (dataset).
 *
 * @property id The unique identifier of the semantic model.
 * @property name The display name of the semantic model.
 * @property workspaceId The ID of the workspace this model belongs to. See {@link WorkSpace}.
 */
export interface SemanticModel {
    /** The unique identifier of the semantic model. */
    id: string;
    /** The display name of the semantic model. */
    name: string;
    /** The ID of the workspace this model belongs to. See {@link WorkSpace}. */
    workspaceId: string;
}

/**
 * Represents a Power BI report.
 *
 * @property id The unique identifier of the report.
 * @property name The display name of the report.
 * @property datasetId The ID of the dataset (semantic model) used by the report. See {@link SemanticModel}.
 * @property workspaceId The ID of the workspace this report belongs to. See {@link WorkSpace}.
 */
export interface Report {
    /** The unique identifier of the report. */
    id: string;
    /** The display name of the report. */
    name: string;
    /** workspace id of dataset */
    datasetWorkspaceId?: string;
    /** The ID of the dataset (semantic model) used by the report. See {@link SemanticModel}. */
    datasetId: string;
    /** The ID of the workspace this report belongs to. See {@link WorkSpace}. */
    workspaceId: string;

    embedUrl: string;
}

export interface ArtifactItem extends SemanticModel {
    type: string; // 'Report' or 'Semantic Model'
    datasetId?: string; // Optional, only for reports
    datasetWorkspaceId?: string; // Optional only for reports
}
