import { useState, useCallback, useMemo } from "react";
import {
  visualTypeToDataRoles,
  schemas,
  propertyToSelector,
  visualTypeProperties,
} from "./visualConfig";
import type { DatasetField } from "../../lib/powerbiLib";
import "./VisualCreatorModal.css";

interface VisualCreatorModalProps {
  isOpen: boolean;
  onClose: () => void;
  onCreateVisual: (visualState: VisualCreatorState) => void;
  report: any;
  authoringPage: any;
  editingVisual?: any;
  datasetFields: DatasetField[];
}

export interface VisualCreatorState {
  visualType: string | null;
  dataRoles: Record<string, any>;
  dataFieldsCount: number;
  properties: {
    legend: boolean;
    xAxis: boolean;
    yAxis: boolean;
    title: boolean;
    titleText: string | null;
    titleAlign: string | null;
  };
  newVisual: any;
}

const initialState: VisualCreatorState = {
  visualType: null,
  dataRoles: {},
  dataFieldsCount: 0,
  properties: {
    legend: true,
    xAxis: true,
    yAxis: true,
    title: true,
    titleText: null,
    titleAlign: null,
  },
  newVisual: null,
};

export default function VisualCreatorModal({
  isOpen,
  onClose,
  onCreateVisual,
  report,
  authoringPage,
  editingVisual,
  datasetFields,
}: VisualCreatorModalProps) {
  const [state, setState] = useState<VisualCreatorState>({ ...initialState });
  const [customTitle, setCustomTitle] = useState("");
  const [titleAlignment, setTitleAlignment] = useState("left");

  const columnFields = useMemo(() => datasetFields.filter((f) => f.type === "column"), [datasetFields]);
  const measureFields = useMemo(() => datasetFields.filter((f) => f.type === "measure"), [datasetFields]);

  const getFieldsForRole = useCallback(
    (dataRoleLabel: string, dataRoleName: string): DatasetField[] => {
      // Keep field choices compatible with role expectations to avoid API addDataField failures.
      if (dataRoleName === "Y" || dataRoleLabel === "Values") {
        return [...measureFields];
      }
      if (dataRoleName === "Category" || dataRoleName === "Series" || dataRoleLabel === "Axis" || dataRoleLabel === "Legend") {
        return [...columnFields];
      }
      if (dataRoleName === "Tooltips") {
        return [...measureFields, ...columnFields];
      }
      return [...columnFields, ...measureFields];
    },
    [columnFields, measureFields]
  );

  const getVisualLayout = () => ({
    width: 1240,
    height: 680,
    x: 10,
    y: 10,
    displayState: { mode: 0 },
  });

  const buildDataFieldTarget = (field: DatasetField) => {
    if (field.type === "measure") {
      return { measure: field.name, table: field.table, schema: "http://powerbi.com/product/schema#measure" };
    }
    return { column: field.name, table: field.table, schema: "http://powerbi.com/product/schema#column" };
  };

  const handleVisualTypeChange = useCallback(
    async (visualTypeDisplayName: string) => {
      if (!report || !authoringPage) return;

      const visualInfo = visualTypeToDataRoles.find((v) => v.displayName === visualTypeDisplayName);
      if (!visualInfo) return;

      const visualType = visualInfo.name;

      if (state.visualType === visualType) return;

      if (state.newVisual) {
        try {
          await authoringPage.deleteVisual(state.newVisual.name);
        } catch (e) {
          console.error("Error deleting visual:", e);
        }
      }

      try {
        const result = await authoringPage.createVisual(visualType, getVisualLayout());
        const visual = result.visual || result;

        visual.setProperty(propertyToSelector("titleSize"), { schema: schemas.property, value: 25 });
        visual.setProperty(propertyToSelector("titleColor"), { schema: schemas.property, value: "#000" });

        if (visualType === "pieChart") {
          visual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: true });
        }

        setState((prev) => ({
          ...prev,
          visualType,
          newVisual: visual,
          dataRoles: {},
          dataFieldsCount: 0,
          properties: { ...initialState.properties },
        }));
      } catch (e) {
        console.error("Error creating visual:", e);
      }
    },
    [report, authoringPage, state.visualType, state.newVisual]
  );

  const handleDataFieldChange = useCallback(
    async (dataRoleName: string, fieldIdentifier: string) => {
      if (!state.newVisual) return;
      if (!dataRoleName) return;

      const isReset = !fieldIdentifier || fieldIdentifier === "";

      try {
        if (isReset) {
          if (state.dataRoles[dataRoleName]) {
            await state.newVisual.removeDataField(dataRoleName, 0);
            setState((prev) => ({
              ...prev,
              dataRoles: { ...prev.dataRoles, [dataRoleName]: null },
              dataFieldsCount: prev.dataFieldsCount - 1,
            }));
          }
          return;
        }

        const [table, name, type] = fieldIdentifier.split("||");

        // Defensive check: value role should always be a measure in this flow.
        if (dataRoleName === "Y" && type !== "measure") {
          return;
        }

        const dataFieldTarget = buildDataFieldTarget({ table, name, type: type as "column" | "measure" });

        if (state.dataRoles[dataRoleName]) {
          await state.newVisual.removeDataField(dataRoleName, 0);
          await state.newVisual.addDataField(dataRoleName, dataFieldTarget);
          setState((prev) => ({
            ...prev,
            dataRoles: { ...prev.dataRoles, [dataRoleName]: dataFieldTarget },
          }));
        } else {
          await state.newVisual.addDataField(dataRoleName, dataFieldTarget);
          const newCount = state.dataFieldsCount + 1;
          setState((prev) => ({
            ...prev,
            dataRoles: { ...prev.dataRoles, [dataRoleName]: dataFieldTarget },
            dataFieldsCount: newCount,
          }));
        }
      } catch (e) {
        console.error(`Error updating data field for role ${dataRoleName}:`, e);
      }
    },
    [state.newVisual, state.dataRoles, state.dataFieldsCount]
  );

  const handlePropertyToggle = useCallback(
    (propertyName: string, value: boolean) => {
      if (!state.newVisual) return;

      setState((prev) => ({
        ...prev,
        properties: { ...prev.properties, [propertyName]: value },
      }));

      state.newVisual.setProperty(propertyToSelector(propertyName), {
        schema: schemas.property,
        value,
      });
    },
    [state.newVisual]
  );

  const handleTitleTextChange = useCallback(
    (text: string) => {
      setCustomTitle(text);
      if (!state.newVisual) return;

      if (text === "") {
        state.newVisual.resetProperty(propertyToSelector("titleText"));
      } else {
        state.newVisual.setProperty(propertyToSelector("titleText"), {
          schema: schemas.property,
          value: text,
        });
      }
    },
    [state.newVisual]
  );

  const handleTitleAlignChange = useCallback(
    (direction: string) => {
      setTitleAlignment(direction);
      if (!state.newVisual) return;

      state.newVisual.setProperty(propertyToSelector("titleAlign"), {
        schema: schemas.property,
        value: direction,
      });

      setState((prev) => ({
        ...prev,
        properties: { ...prev.properties, titleAlign: direction },
      }));
    },
    [state.newVisual]
  );

  const buildFinalState = (): VisualCreatorState => ({
    ...state,
    properties: {
      ...state.properties,
      titleText: customTitle || null,
      titleAlign: titleAlignment,
    },
  });

  const handleCreate = () => {
    onCreateVisual(buildFinalState());
    handleClose();
  };

  const handleClose = () => {
    if (state.newVisual && authoringPage) {
      authoringPage.deleteVisual(state.newVisual.name).catch(() => {
        // Visual may already be gone — ignore
      });
    }
    setState({ ...initialState });
    setCustomTitle("");
    setTitleAlignment("left");
    onClose();
  };

  const currentVisualInfo = visualTypeToDataRoles.find((v) => v.name === state.visualType);
  const availableProperties = state.visualType ? visualTypeProperties[state.visualType] || [] : [];
  const canCreate = state.dataFieldsCount >= 2;

  if (!isOpen) return null;

  return (
    <div className="qvc-modal-overlay" onClick={handleClose}>
      <div className="qvc-modal" onClick={(e) => e.stopPropagation()}>
        {/* Header */}
        <div className="qvc-modal-header">
          <h2>Create quick visual</h2>
          <button className="qvc-close-btn" onClick={handleClose} aria-label="Close dialog">
            &times;
          </button>
        </div>

        {/* Body */}
        <div className="qvc-modal-body">
          <div className="qvc-options-panel">
            {/* Visual Type Selection */}
            <div className="qvc-section">
              <span className="qvc-section-title">Choose the visual type</span>
              <select
                className="qvc-select"
                value={state.visualType ? visualTypeToDataRoles.find((v) => v.name === state.visualType)?.displayName || "" : ""}
                onChange={(e) => handleVisualTypeChange(e.target.value)}
              >
                <option value="">Select visual type</option>
                {visualTypeToDataRoles.map((vt) => (
                  <option key={vt.name} value={vt.displayName}>
                    {vt.displayName}
                  </option>
                ))}
              </select>
            </div>

            {/* Data Fields Selection */}
            <div className={`qvc-section ${!state.visualType ? "qvc-disabled" : ""}`}>
              <span className="qvc-section-title">Set your fields</span>
              <div className="qvc-fields-container">
                {currentVisualInfo?.dataRoles.map((dataRole, idx) => {
                  const dataRoleName = currentVisualInfo.dataRoleNames[idx];
                  if (!dataRoleName) return null;

                  const fieldsForRole = getFieldsForRole(dataRole, dataRoleName);

                  return (
                    <div key={dataRoleName} className="qvc-field-row">
                      <span className="qvc-field-label">{dataRole}</span>
                      <select
                        className="qvc-select qvc-field-select"
                        disabled={!state.visualType}
                        onChange={(e) => handleDataFieldChange(dataRoleName, e.target.value)}
                        defaultValue=""
                      >
                        <option value="">Select {dataRole}</option>
                        {fieldsForRole.map((field) => (
                          <option key={`${field.table}.${field.name}`} value={`${field.table}||${field.name}||${field.type}`}>
                            {field.name} ({field.table})
                          </option>
                        ))}
                      </select>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Properties */}
            <div className={`qvc-section ${!canCreate ? "qvc-disabled" : ""}`}>
              <span className="qvc-section-title">Format your visual</span>
              <div className="qvc-properties-container">
                {/* Legend Toggle */}
                <div className={`qvc-toggle-row ${!availableProperties.includes("legend") ? "qvc-toggle-disabled" : ""}`}>
                  <span>Legend</span>
                  <label className="qvc-switch">
                    <input
                      type="checkbox"
                      checked={state.properties.legend}
                      disabled={!canCreate || !availableProperties.includes("legend")}
                      onChange={(e) => handlePropertyToggle("legend", e.target.checked)}
                    />
                    <span className="qvc-slider"></span>
                  </label>
                </div>

                {/* X Axis Toggle */}
                <div className={`qvc-toggle-row ${!availableProperties.includes("xAxis") ? "qvc-toggle-disabled" : ""}`}>
                  <span>Category Axis</span>
                  <label className="qvc-switch">
                    <input
                      type="checkbox"
                      checked={state.properties.xAxis}
                      disabled={!canCreate || !availableProperties.includes("xAxis")}
                      onChange={(e) => handlePropertyToggle("xAxis", e.target.checked)}
                    />
                    <span className="qvc-slider"></span>
                  </label>
                </div>

                {/* Y Axis Toggle */}
                <div className={`qvc-toggle-row ${!availableProperties.includes("yAxis") ? "qvc-toggle-disabled" : ""}`}>
                  <span>Value Axis</span>
                  <label className="qvc-switch">
                    <input
                      type="checkbox"
                      checked={state.properties.yAxis}
                      disabled={!canCreate || !availableProperties.includes("yAxis")}
                      onChange={(e) => handlePropertyToggle("yAxis", e.target.checked)}
                    />
                    <span className="qvc-slider"></span>
                  </label>
                </div>

                {/* Title Toggle */}
                <div className="qvc-toggle-row">
                  <span>Title</span>
                  <label className="qvc-switch">
                    <input
                      type="checkbox"
                      checked={state.properties.title}
                      disabled={!canCreate}
                      onChange={(e) => handlePropertyToggle("title", e.target.checked)}
                    />
                    <span className="qvc-slider"></span>
                  </label>
                </div>

                {/* Custom Title */}
                <div className="qvc-title-input-row">
                  <input
                    type="text"
                    className="qvc-title-input"
                    placeholder="Type your personalized title"
                    value={customTitle}
                    disabled={!canCreate || !state.properties.title}
                    onChange={(e) => handleTitleTextChange(e.target.value)}
                  />
                  <button
                    className="qvc-erase-btn"
                    disabled={!canCreate || !state.properties.title}
                    onClick={() => handleTitleTextChange("")}
                    aria-label="Clear title"
                  >
                    ✕
                  </button>
                </div>

                {/* Title Alignment */}
                <div className="qvc-alignment-row">
                  <span>Title alignment</span>
                  <div className="qvc-alignment-buttons">
                    <button
                      className={`qvc-align-btn ${titleAlignment === "left" ? "qvc-align-active" : ""}`}
                      disabled={!canCreate || !state.properties.title}
                      onClick={() => handleTitleAlignChange("left")}
                      aria-label="Left align"
                    >
                      ≡
                    </button>
                    <button
                      className={`qvc-align-btn ${titleAlignment === "center" ? "qvc-align-active" : ""}`}
                      disabled={!canCreate || !state.properties.title}
                      onClick={() => handleTitleAlignChange("center")}
                      aria-label="Center align"
                    >
                      ≡
                    </button>
                    <button
                      className={`qvc-align-btn ${titleAlignment === "right" ? "qvc-align-active" : ""}`}
                      disabled={!canCreate || !state.properties.title}
                      onClick={() => handleTitleAlignChange("right")}
                      aria-label="Right align"
                    >
                      ≡
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="qvc-modal-footer">
          <button className="qvc-create-btn" disabled={!canCreate} onClick={handleCreate}>
            Create
          </button>
        </div>
      </div>
    </div>
  );
}
