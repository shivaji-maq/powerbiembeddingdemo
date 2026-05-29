import { useEffect, useRef, useState } from "react";
import { models } from "powerbi-client";
import "./App.css";
import logo from "./assets/logo.svg";
import { useAuth } from "./hooks/useAuth";
import { useDispatch, useSelector } from "react-redux";
import { fetchReports, fetchWorkspaces } from "./redux/slices/powerBISlice/powerBISlice";
import { PersonalizedEditableReport } from "./components/PersonalizedEditableReport/PersonalizedEditableReport";
import { QuickVisualCreator } from "./components/QuickVisualCreator";
import { getAccessToken } from "./configs/msalInstance";
import type { RootState, AppDispatch } from "./redux/store";

interface WorkSpace {
  id: string;
  name: string;
  isReadOnly?: boolean;
}

interface Report {
  id: string;
  name: string;
  embedUrl: string;
  datasetId?: string;
}

interface SelectedData {
  dataPoints: Array<{
    identity: Array<{
      target: { column: string };
      equals: string;
    }>;
  }>;
}

function App() {
  const [selectedReport, setSelectedReport] = useState<Report | null>(null);
  const [selectedWorkspace, setSelectedWorkspace] = useState<WorkSpace | null>(null);
  const [accessKey, setAccessKey] = useState("");
  const [wsFilter, setWsFilter] = useState("");
  const [rFilter, setRFilter] = useState("");
  const [authError, setAuthError] = useState<string>("");
  const [mode, setMode] = useState<"viewer" | "creator">("viewer");
  const { isAuthenticated, user, account, error: authHookError } = useAuth();
  const { workspaces, fetchingWorkspaces, fetchingReports, reports, errorInWorkspace } = useSelector(
    (state: RootState) => state.powerBI
  );
  const dispatch = useDispatch<AppDispatch>();

  const reportRef = useRef<any>(null);
  const quickVisualCreatorRef = useRef<any>(null);

  useEffect(() => {
    console.log("App mounted, isAuthenticated:", isAuthenticated);
    console.log("User:", user);
  }, []);

  useEffect(() => {
    console.log("Auth state changed:", { isAuthenticated, user, account, authHookError });
    console.log("Power BI state:", { workspaces: workspaces?.length, reports: reports?.length, fetchingWorkspaces, fetchingReports, errorInWorkspace });
    const fun = async () => {
      if (isAuthenticated) {
        try {
          const ak = await getAccessToken();
          console.log("Access token acquired:", !!ak);
          if (ak) {
            setAccessKey(ak);
            console.log("Dispatching fetchWorkspaces with token");
            dispatch(fetchWorkspaces(ak));
          } else {
            const error = "Power BI access token was not acquired.";
            console.error(error);
            setAuthError(error);
          }
        } catch (err) {
          const error = err instanceof Error ? err.message : "Unknown auth error";
          console.error("Auth error:", error);
          setAuthError(error);
        }
      } else {
        console.log("User not authenticated. Auth state:", { user, account, authHookError });
      }
    };
    void fun();
  }, [dispatch, isAuthenticated, user, account, authHookError]);

  const settings = {
    extensions: [
      {
        command: {
          name: "addComment",
          title: "Add Comment",
          extend: {
            visualOptionsMenu: {
              title: "Add Comment",
              menuLocation: models.MenuLocation.Top,
            },
          },
        },
      },
      {
        command: {
          name: "changeVisual",
          title: "Change visual",
          extend: {
            visualOptionsMenu: {
              title: "Change visual",
              menuLocation: models.MenuLocation.Top,
            },
          },
        },
      },
      {
        command: {
          name: "createQuickVisual",
          title: "Create quick visual",
          extend: {
            visualOptionsMenu: {
              title: "Create quick visual",
              menuLocation: models.MenuLocation.Top,
            },
          },
        },
      },
    ],
  };

  const addComment = async () => {
    const selectedDataStr = window.localStorage.getItem("selectedData");
    if (!selectedDataStr) {
      alert("No data point selected.");
      return;
    }

    let id: string | undefined;
    const selectedData: SelectedData = JSON.parse(selectedDataStr);
    selectedData.dataPoints[0].identity.forEach((iden) => {
      if (iden.target.column === "Id") id = iden.equals;
    });
    const comment = prompt("Please add a comment for sales id: " + id);

    if (!comment) {
      alert("No comment provided.");
      return;
    }

    fetch("https://localhost:44301/api/v1/Auth/UpdateComment", {
      method: "post",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        id: id,
        comment: comment,
      }),
    })
      .then(async () => {
        alert(`Saved comment: "${comment}" for the corresponding data.`);
        if (reportRef.current && typeof reportRef.current.reload === "function") {
          await reportRef.current.reload();
        }
      })
      .catch((err: Error) => {
        alert("error: " + err.message);
      });
  };

  const SelectADataPoint = (data: any): void => {
    window.localStorage.setItem("selectedData", JSON.stringify(data, null, 2));
    // setSelectedData(data);
  };

  const handleCommandTriggered = (event: any): void => {
    const commandDetails = event?.detail || {};
    if (commandDetails.command === "addComment" || commandDetails.command === "Add Comment") {
      void addComment();
    }
  };

  const handleDataSelected = (event: any): void => {
    const data = event?.detail;
    SelectADataPoint(data);
  };

  const addCustomFunctionsInEmbeddedReport = (report: any): void => {
    // Keep existing handlers and only replace our own callback references.
    report.off("commandTriggered", handleCommandTriggered);
    report.on("commandTriggered", handleCommandTriggered);

    report.off("dataSelected", handleDataSelected);
    report.on("dataSelected", handleDataSelected);
  };

  return (
    isAuthenticated && !authError && !authHookError ? (
      <div className="app-root">
        {/* Header */}
        <header className="app-header">
          <div className="header-left">
            <img src={logo} alt="Logo" className="app-logo" />
            <span className="app-title">Power BI Report Embedder</span>
          </div>

          <div>
            <p>{user?.displayName}</p>
            <p>{user?.email}</p>
          </div>
        </header>

        <div className="app-main">
          {/* Sidebar for workspace/report selection (User Mode) */}

          <aside className="sidebar">
            <div className="sidebar-section">
              <input className="sidebar-input" type="text" value={wsFilter} onChange={(e) => setWsFilter(e.target.value)} placeholder="Filter workspace" />
              <div className="sidebar-list">
                {fetchingWorkspaces ? (
                  <>Loading...</>
                ) : (
                  workspaces
                    .filter((x: WorkSpace) => x.name.toLowerCase().includes(wsFilter.toLowerCase()))
                    .filter((x: WorkSpace) => !x.name.toLowerCase().includes("pepsi"))
                    .map((w: WorkSpace) => (
                      <div className="sidebar-list-item" key={w.id}>
                        <span className="sidebar-list-title" title={w.name}>
                          {w.name}
                        </span>
                        <button className="mini-btn" onClick={() => {
                          setSelectedWorkspace(w);
                          void dispatch(fetchReports({ workspaceId: w.id, accessToken: accessKey }));
                        }}>
                          Fetch Reports
                        </button>
                      </div>
                    ))
                )}
              </div>
            </div>
            <div className="sidebar-section">
              <input className="sidebar-input" type="text" value={rFilter} onChange={(e) => setRFilter(e.target.value)} placeholder="Filter reports" />
              <div className="sidebar-list">
                {fetchingReports ? (
                  <>Loading...</>
                ) : (
                  reports
                    .filter((x: Report) => x.name.toLowerCase().includes(rFilter.toLowerCase()))
                    .filter((x: Report) => !x.name.toLowerCase().includes("pepsi"))
                    .map((w: Report) => (
                      <div className="sidebar-list-item" key={w.id}>
                        <span className="sidebar-list-title" title={w.name}>
                          {w.name}
                        </span>
                        <button className="mini-btn" onClick={() => setSelectedReport(w)}>
                          Embed
                        </button>
                      </div>
                    ))
                )}
              </div>
            </div>
          </aside>

          {/* Main content area */}
          <main className="main-content">
            {selectedReport && selectedWorkspace ? (
              <>
                {mode === "creator" ? (
                  <>
                    {/* Toggle button row for creator mode */}
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                      <button
                        className="primary-btn"
                        onClick={() => quickVisualCreatorRef.current?.openModal()}
                        style={{ background: "#1a237e", color: "#fff" }}
                      >
                        + Create Visual
                      </button>
                      <button
                        className="primary-btn"
                        onClick={() => setMode("viewer")}
                      >
                        Report Viewer
                      </button>
                    </div>
                    <QuickVisualCreator
                      ref={quickVisualCreatorRef}
                      accessToken={accessKey}
                      embedUrl={selectedReport.embedUrl}
                      reportId={selectedReport.id}
                      datasetId={selectedReport.datasetId || ""}
                      workspaceId={selectedWorkspace.id}
                      tokenType="Aad"
                    />
                  </>
                ) : (
                  <PersonalizedEditableReport
                    key={`${selectedWorkspace.id}_${selectedReport.id}`}
                    reportRef={reportRef}
                    reportId={selectedReport.id}
                    embedUrl={selectedReport.embedUrl}
                    embedReportEventHandlers={new Map()}
                    reportSettings={settings}
                    accessToken={accessKey}
                    userId={user?.id || user?.userPrincipalName || user?.email || ""}
                    workspaceId={selectedWorkspace.id}
                    allowEdit={!selectedWorkspace.isReadOnly}
                    onReportLoadReportAttachmentFunction={(report: any) => {
                      addCustomFunctionsInEmbeddedReport(report);
                      reportRef.current = report;
                    }}
                    toggleButton={
                      <button
                        className="primary-btn"
                        onClick={() => setMode("creator")}
                      >
                        Quick Visual Creator
                      </button>
                    }
                  />
                )}
              </>
            ) : (
              <div className="empty-state">Select a report to embed</div>
            )}
          </main>
        </div>
      </div>
    ) : authError || authHookError ? (
      <div style={{ padding: "20px", color: "red", fontFamily: "Arial" }}>
        <h2>Authentication Error</h2>
        <p>{authError || authHookError}</p>
        <p>Please check the browser console for more details.</p>
      </div>
    ) : errorInWorkspace ? (
      <div style={{ padding: "20px", color: "red", fontFamily: "Arial" }}>
        <h2>Error Loading Workspaces</h2>
        <p>{errorInWorkspace}</p>
        <p>Please check the browser console for more details.</p>
      </div>
    ) : isAuthenticated && fetchingWorkspaces ? (
      <div style={{ padding: "20px", fontFamily: "Arial", textAlign: "center" }}>
        <h2>Loading Power BI workspaces...</h2>
        <p>Please wait while we fetch your workspaces and reports.</p>
      </div>
    ) : (
      <div style={{ padding: "20px", fontFamily: "Arial", textAlign: "center" }}>
        <p>Initializing authentication...</p>
        <p>If this page doesn't load, please check the browser console for errors.</p>
      </div>
    )
  );
}

export default App;
