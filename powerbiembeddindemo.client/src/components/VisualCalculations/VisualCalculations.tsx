// components/VisualCalculations/VisualCalculations.tsx
import React, { useEffect, useState } from "react";
import { getVisualCalculations } from "../../lib/powerbiLib/personalization";
import "./VisualCalculations.css";

interface VisualDetail {
  id: string;
  name: string;
  type: string;
  dataPoints: number;
}

interface VisualCalculationsProps {
  reportRef: any;
  pageName: string;
}

export const VisualCalculations: React.FC<VisualCalculationsProps> = ({ reportRef, pageName }) => {
  const [calculations, setCalculations] = useState<VisualDetail[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const loadCalculations = async () => {
      if (!reportRef?.current || !pageName) return;

      setLoading(true);
      setError(null);

      try {
        const calcs = await getVisualCalculations(reportRef.current, pageName);
        setCalculations(calcs as any);
      } catch (err) {
        setError(err instanceof Error ? err.message : "Failed to load visual calculations");
      } finally {
        setLoading(false);
      }
    };

    loadCalculations();
  }, [reportRef, pageName]);

  if (loading) {
    return <div className="visual-calculations loading">Loading calculations...</div>;
  }

  if (error) {
    return <div className="visual-calculations error">Error: {error}</div>;
  }

  return (
    <div className="visual-calculations">
      <div className="visual-calculations-header">
        <h3>📊 Visual Calculations & Details</h3>
        <p className="visual-count">{calculations.length} visuals on this page</p>
      </div>

      {calculations.length === 0 ? (
        <p className="no-visuals">No visuals found on this page</p>
      ) : (
        <div className="calculations-grid">
          {calculations.map((calc) => (
            <div key={calc.id} className="calculation-card">
              <div className="calculation-header">
                <h4>{calc.name}</h4>
                <span className="visual-type">{calc.type}</span>
              </div>
              <div className="calculation-details">
                <div className="detail-item">
                  <span className="detail-label">Visual ID:</span>
                  <span className="detail-value">{calc.id}</span>
                </div>
                <div className="detail-item">
                  <span className="detail-label">Data Points:</span>
                  <span className="detail-value">{calc.dataPoints}</span>
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
