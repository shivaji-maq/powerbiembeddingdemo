/* eslint-disable @typescript-eslint/no-explicit-any */
// components/ReportEditor/ReportEditor.tsx
import React, { useState } from "react";
import { setReportMode } from "../../lib/powerbiLib/personalization";
import "./ReportEditor.css";

interface ReportEditorProps {
  reportRef: any;
  reportId: string;
  userId: string;
  workspaceId: string;
  allowEdit?: boolean;
  onSave?: () => void;
  onCancel?: () => void;
  onModeChange?: (mode: "view" | "edit") => void;
}

export const ReportEditor: React.FC<ReportEditorProps> = ({ reportRef, allowEdit = true, onSave, onCancel, onModeChange }) => {
  const [isEditMode, setIsEditMode] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleToggleEdit = async () => {
    try {
      setError(null);

      if (!isEditMode) {
        console.log("[ReportEditor] Attempting to enable edit mode. allowEdit:", allowEdit);
        try {
          await setReportMode(reportRef.current, "edit");
          setIsEditMode(true);
          onModeChange?.("edit");
          console.log("[ReportEditor] Edit mode enabled successfully");
        } catch (err) {
          const message = err instanceof Error ? err.message : "Failed to enable edit mode";
          console.error("[ReportEditor] Error enabling edit mode:", message);

          // Check if it's a permission error
          if (message.toLowerCase().includes("insufficient") || message.toLowerCase().includes("permission") || message.toLowerCase().includes("edit rights")) {
            setError("Edit mode is blocked by Power BI permissions. Ensure this account has edit rights in the workspace/report.");
          } else {
            setError(message);
          }
        }
      } else {
        await setReportMode(reportRef.current, "view");
        setIsEditMode(false);
        onModeChange?.("view");
        console.log("[ReportEditor] Switched to view mode");
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : "Failed to toggle edit mode";
      console.error("[ReportEditor] Toggle edit error:", message);
      setError(message);
    }
  };

  const handleSave = async () => {
    setIsSaving(true);
    setError(null);
    try {
      // Save report directly to Power BI service
      if (!reportRef.current || typeof reportRef.current.save !== "function") {
        throw new Error("Report is not ready to save.");
      }
      await reportRef.current.save();
      console.log("[ReportEditor] Report saved to Power BI service successfully");

      // Switch back to view mode (changes are already saved to service)
      await setReportMode(reportRef.current, "view");
      setIsEditMode(false);
      onModeChange?.("view");
      onSave?.();
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : "Failed to save report";
      console.error("[ReportEditor] Save error:", errorMsg);
      setError(errorMsg);
    } finally {
      setIsSaving(false);
    }
  };

  const handleCancel = async () => {
    try {
      await setReportMode(reportRef.current, "view");
    } catch (err) {
      console.warn("[ReportEditor] Error switching to view mode on cancel:", err);
    }
    setIsEditMode(false);
    setError(null);
    onModeChange?.("view");
    onCancel?.();
  };

  return (
    <div className="report-editor-toolbar">
      {error && <div className="error-message">{error}</div>}

      {!isEditMode ? (
        <button onClick={handleToggleEdit} className="btn btn-edit" disabled={isSaving} title="Switch report to edit mode">
          Edit Report
        </button>
      ) : (
        <div className="edit-mode-controls">
          <button onClick={handleSave} disabled={isSaving} className="btn btn-save">
            {isSaving ? "Saving..." : "Save Changes"}
          </button>
          <button onClick={handleCancel} className="btn btn-cancel">
            Cancel
          </button>
        </div>
      )}
    </div>
  );
};
