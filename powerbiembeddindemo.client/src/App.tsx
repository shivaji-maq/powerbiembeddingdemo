import { useEffect, useRef, useState } from "react";
import { models } from "powerbi-client";
import "./App.css";
import logo from "./assets/logo.svg";
import { useAuth } from "./hooks/useAuth";
import { useDispatch, useSelector } from "react-redux";
import { fetchReports, fetchWorkspaces } from "./redux/slices/powerBISlice/powerBISlice";
import { removeBookmark } from "./redux/slices/bookmarkSlice/bookmarkSlice";
import { PersonalizedEditableReport } from "./components/PersonalizedEditableReport/PersonalizedEditableReport";
import { QuickVisualCreator, LayoutCustomizer } from "./components/QuickVisualCreator";
import type { LayoutCustomizerHandle } from "./components/QuickVisualCreator";
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
  const [selectedBookmarkId, setSelectedBookmarkId] = useState("");
  const [layoutReport, setLayoutReport] = useState<any>(null);
  const [layoutPage, setLayoutPage] = useState<any>(null);
  const bookmarks = useSelector((state: RootState) => state.bookmarks.bookmarks);
  const { isAuthenticated, user, account, error: authHookError } = useAuth();
  const { workspaces, fetchingWorkspaces, fetchingReports, reports, errorInWorkspace } = useSelector(
    (state: RootState) => state.powerBI
  );
  const dispatch = useDispatch<AppDispatch>();

  const reportRef = useRef<any>(null);
  const quickVisualCreatorRef = useRef<any>(null);
  const layoutCustomizerRef = useRef<LayoutCustomizerHandle>(null);

  useEffect(() => {
    console.log("App mounted, isAuthenticated:", isAuthenticated);
    console.log("User:", user);
  }, []);

  useEffect(() => {
    const fun = async () => {
      if (isAuthenticated) {
        try {
          const ak = await getAccessToken();
          if (ak) {
            setAccessKey(ak);
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
                    {/* Toolbar row for creator mode */}
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, gap: 8, flexWrap: "wrap" }}>
                      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                        <button
                          className="primary-btn"
                          onClick={() => quickVisualCreatorRef.current?.openModal()}
                          style={{ background: "#1a237e", color: "#fff" }}
                        >
                          + Create Visual
                        </button>

                        {/* Save View = manual bookmark of last visual */}
                        <button
                          className="primary-btn"
                          style={{ background: "#2e7d55", color: "#fff", fontWeight: 600 }}
                          onClick={() => quickVisualCreatorRef.current?.bookmarkLastVisual("")}
                        >
                          Save View
                        </button>

                        {/* Bookmark selector + delete + load */}
                        <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                          <select
                            className="sidebar-input"
                            style={{ margin: 0, minWidth: 180, maxWidth: 280, height: 34 }}
                            value={selectedBookmarkId}
                            onChange={(e) => setSelectedBookmarkId(e.target.value)}
                          >
                            <option value="" disabled>Select bookmark</option>
                            {bookmarks.map((b) => (
                              <option key={b.id} value={b.id}>{b.name}</option>
                            ))}
                          </select>
                          <button
                            title="Delete bookmark"
                            disabled={!selectedBookmarkId}
                            style={{
                              width: 34, height: 34, border: "1.5px solid #e57373", borderRadius: 6,
                              background: "#fff", color: "#e53935", cursor: selectedBookmarkId ? "pointer" : "not-allowed",
                              fontSize: "1rem", display: "flex", alignItems: "center", justifyContent: "center",
                              opacity: selectedBookmarkId ? 1 : 0.4,
                            }}
                            onClick={() => {
                              dispatch(removeBookmark(selectedBookmarkId));
                              setSelectedBookmarkId("");
                            }}
                          >
                            ✕
                          </button>
                          <button
                            title="Load bookmark"
                            disabled={!selectedBookmarkId}
                            style={{
                              width: 34, height: 34, border: "1.5px solid #43a047", borderRadius: 6,
                              background: "#fff", color: "#2e7d32", cursor: selectedBookmarkId ? "pointer" : "not-allowed",
                              fontSize: "1.1rem", display: "flex", alignItems: "center", justifyContent: "center",
                              opacity: selectedBookmarkId ? 1 : 0.4,
                            }}
                            onClick={() => {
                              const bm = bookmarks.find((b) => b.id === selectedBookmarkId);
                              if (bm) quickVisualCreatorRef.current?.loadBookmark(bm);
                            }}
                          >
                            ↻
                          </button>
                        </div>

                      </div>
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
                  <>
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
                    onReportReady={(report, page) => {
                      setLayoutReport((prev: any) => (prev === report ? prev : report));
                      setLayoutPage((prev: any) =>
                        prev?.name === page?.name ? prev : page
                      );
                    }}
                    toggleButton={
                      <button
                        className="primary-btn"
                        onClick={() => setMode("creator")}
                      >
                        Quick Visual Creator
                      </button>
                    }
                    layoutControls={
                      layoutReport && layoutPage ? (
                        <LayoutCustomizer ref={layoutCustomizerRef} report={layoutReport} page={layoutPage} />
                      ) : null
                    }
                    layoutCustomizerRef={layoutCustomizerRef}
                  />
                  </>
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
