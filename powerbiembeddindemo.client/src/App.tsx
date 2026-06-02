/* eslint-disable @typescript-eslint/no-explicit-any */
import { useEffect, useMemo, useRef, useState } from "react";
import { models } from "powerbi-client";
import "./App.css";
import logo from "./assets/logo.png";
import { useAuth } from "./hooks/useAuth";
import { useDispatch, useSelector } from "react-redux";
import { fetchWorkspaces } from "./redux/slices/powerBISlice/powerBISlice";
import { removeBookmark } from "./redux/slices/bookmarkSlice/bookmarkSlice";
import { PersonalizedEditableReport } from "./components/PersonalizedEditableReport/PersonalizedEditableReport";
import { QuickVisualCreator, LayoutCustomizer } from "./components/QuickVisualCreator";
import type { LayoutCustomizerHandle } from "./components/QuickVisualCreator";
import { getAccessToken } from "./configs/msalInstance";
import type { RootState, AppDispatch } from "./redux/store";
import { reportsToEmbed, type ReportToEmbed } from "./constants/reports";
import { POWER_BI_API_CONST } from "./lib/powerbiLib";
import axios from "axios";

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

const GLOBAL_DATE_MIN = "1999-01-01";
const GLOBAL_DATE_MAX = "2026-12-31";
const GLOBAL_DATE_FILTER_STORAGE_KEY = "globalDateFilterRange";
const DATE_INPUT_VALUE_PATTERN = /^\d{4}-\d{2}-\d{2}$/;

interface GlobalDateRange {
  from: string;
  to: string;
}

type GlobalDateFilterTarget = NonNullable<ReportToEmbed["globalDateFilter"]>;

const normalizeDateInputValue = (value: string | undefined, fallback: string) => {
  if (!value || !DATE_INPUT_VALUE_PATTERN.test(value)) {
    return fallback;
  }

  if (value < GLOBAL_DATE_MIN) {
    return GLOBAL_DATE_MIN;
  }

  if (value > GLOBAL_DATE_MAX) {
    return GLOBAL_DATE_MAX;
  }

  return value;
};

const getInitialGlobalDateRange = (): GlobalDateRange => {
  if (typeof window === "undefined") {
    return { from: GLOBAL_DATE_MIN, to: GLOBAL_DATE_MAX };
  }

  try {
    const storedRange = window.localStorage.getItem(GLOBAL_DATE_FILTER_STORAGE_KEY);
    if (!storedRange) {
      return { from: GLOBAL_DATE_MIN, to: GLOBAL_DATE_MAX };
    }

    const parsedRange = JSON.parse(storedRange) as Partial<GlobalDateRange>;
    return {
      from: normalizeDateInputValue(parsedRange.from, GLOBAL_DATE_MIN),
      to: normalizeDateInputValue(parsedRange.to, GLOBAL_DATE_MAX),
    };
  } catch {
    return { from: GLOBAL_DATE_MIN, to: GLOBAL_DATE_MAX };
  }
};

const toStartOfDayIso = (dateValue: string) => `${dateValue}T00:00:00.000Z`;
const toEndOfDayIso = (dateValue: string) => `${dateValue}T23:59:59.999Z`;

const buildGlobalDateFilters = (target: GlobalDateFilterTarget | undefined, dateRange: GlobalDateRange): models.ReportLevelFilters[] => {
  if (!target || !dateRange.from || !dateRange.to) {
    return [];
  }

  const from = dateRange.from <= dateRange.to ? dateRange.from : dateRange.to;
  const to = dateRange.from <= dateRange.to ? dateRange.to : dateRange.from;

  return [
    {
      $schema: "http://powerbi.com/product/schema#advanced",
      target,
      filterType: models.FilterType.Advanced,
      logicalOperator: "And",
      conditions: [
        {
          operator: "GreaterThanOrEqual",
          value: toStartOfDayIso(from),
        },
        {
          operator: "LessThan",
          value: toEndOfDayIso(to),
        },
      ],
      displaySettings: {
        displayName: "Global Date",
      },
    },
  ];
};

const hasSameFilterTarget = (filter: any, targetFilter: any) => {
  const filterTarget = filter?.target;
  const target = targetFilter?.target;

  return !!filterTarget && !!target && filterTarget.table === target.table && filterTarget.column === target.column;
};

const mergeGlobalDateFilters = (currentFilters: models.IFilter[], globalFilters: models.ReportLevelFilters[]) => {
  if (globalFilters.length === 0) {
    return currentFilters;
  }

  return [...currentFilters.filter((filter) => !globalFilters.some((globalFilter) => hasSameFilterTarget(filter, globalFilter))), ...globalFilters];
};

// interface SelectedData {
//   dataPoints: Array<{
//     identity: Array<{
//       target: { column: string };
//       equals: string;
//     }>;
//   }>;
// }

function App() {
  const [selectedReportRaw, setSelectedReportRaw] = useState<ReportToEmbed | null>(null);
  const [selectedReport, setSelectedReport] = useState<Report | null>(null);
  const [culture, setCulture] = useState<string>("en-US");
  const [selectedWorkspace, setSelectedWorkspace] = useState<WorkSpace | null>(null);
  const [accessKey, setAccessKey] = useState("");
  // const [wsFilter, setWsFilter] = useState("");
  // const [rFilter, setRFilter] = useState("");
  const [authError, setAuthError] = useState<string>("");
  const [mode, setMode] = useState<"viewer" | "creator">("viewer");
  const [selectedBookmarkId, setSelectedBookmarkId] = useState("");
  const [layoutReport, setLayoutReport] = useState<any>(null);
  const [layoutPage, setLayoutPage] = useState<any>(null);
  const [globalDateRange, setGlobalDateRange] = useState<GlobalDateRange>(getInitialGlobalDateRange);
  const [appliedGlobalDateRange, setAppliedGlobalDateRange] = useState<GlobalDateRange>(getInitialGlobalDateRange);
  const [globalDateStatus, setGlobalDateStatus] = useState("");
  const bookmarks = useSelector((state: RootState) => state.bookmarks.bookmarks);
  const { isAuthenticated, user, account, error: authHookError } = useAuth();
  const {
    workspaces,
    fetchingWorkspaces,
    // fetchingReports, reports,
    errorInWorkspace,
  } = useSelector((state: RootState) => state.powerBI);
  const dispatch = useDispatch<AppDispatch>();

  const reportRef = useRef<any>(null);
  const quickVisualCreatorRef = useRef<any>(null);
  const layoutCustomizerRef = useRef<LayoutCustomizerHandle>(null);
  const fromDateInputRef = useRef<HTMLInputElement | null>(null);
  const toDateInputRef = useRef<HTMLInputElement | null>(null);

  const globalDateFilters = useMemo<models.ReportLevelFilters[]>(() => {
    return buildGlobalDateFilters(selectedReportRaw?.globalDateFilter, appliedGlobalDateRange);
  }, [appliedGlobalDateRange.from, appliedGlobalDateRange.to, selectedReportRaw?.globalDateFilter]);

  // useEffect(() => {
  //   console.log("App mounted, isAuthenticated:", isAuthenticated);
  //   console.log("User:", user);
  // }, []);

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

  useEffect(() => {
    window.localStorage.setItem(GLOBAL_DATE_FILTER_STORAGE_KEY, JSON.stringify(appliedGlobalDateRange));
  }, [appliedGlobalDateRange]);

  const settings = {
    settings: {
      localeSettings: {
        language: culture, // Sets the display language (e.g., Spanish)
        formatLocale: culture, // Sets regional format (e.g., currency/dates)
      },
    },
    extensions: [
      // {
      //   command: {
      //     name: "addComment",
      //     title: "Add Comment",
      //     extend: {
      //       visualOptionsMenu: {
      //         title: "Add Comment",
      //         menuLocation: models.MenuLocation.Top,
      //       },
      //     },
      //   },
      // },
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

  // const addComment = async () => {
  //   const selectedDataStr = window.localStorage.getItem("selectedData");
  //   if (!selectedDataStr) {
  //     alert("No data point selected.");
  //     return;
  //   }

  //   let id: string | undefined;
  //   const selectedData: SelectedData = JSON.parse(selectedDataStr);
  //   selectedData.dataPoints[0].identity.forEach((iden) => {
  //     if (iden.target.column === "Id") id = iden.equals;
  //   });
  //   const comment = prompt("Please add a comment for sales id: " + id);

  //   if (!comment) {
  //     alert("No comment provided.");
  //     return;
  //   }

  //   fetch("https://localhost:44301/api/v1/Auth/UpdateComment", {
  //     method: "post",
  //     headers: {
  //       "Content-Type": "application/json",
  //     },
  //     body: JSON.stringify({
  //       id: id,
  //       comment: comment,
  //     }),
  //   })
  //     .then(async () => {
  //       alert(`Saved comment: "${comment}" for the corresponding data.`);
  //       if (reportRef.current && typeof reportRef.current.reload === "function") {
  //         await reportRef.current.reload();
  //       }
  //     })
  //     .catch((err: Error) => {
  //       alert("error: " + err.message);
  //     });
  // };

  const SelectADataPoint = (data: any): void => {
    window.localStorage.setItem("selectedData", JSON.stringify(data, null, 2));
    // setSelectedData(data);
  };

  // const handleCommandTriggered = (event: any): void => {
  //   // const commandDetails = event?.detail || {};
  //   // if (commandDetails.command === "addComment" || commandDetails.command === "Add Comment") {
  //   //   void addComment();
  //   // }
  // };

  const handleDataSelected = (event: any): void => {
    const data = event?.detail;
    SelectADataPoint(data);
  };

  const addCustomFunctionsInEmbeddedReport = (report: any): void => {
    // Keep existing handlers and only replace our own callback references.
    // report.off("commandTriggered", handleCommandTriggered);
    // report.on("commandTriggered", handleCommandTriggered);

    report.off("dataSelected", handleDataSelected);
    report.on("dataSelected", handleDataSelected);
  };
  const mainContentDivRef = useRef<HTMLDivElement | null>(null);

  const openDatePicker = (input: HTMLInputElement | null) => {
    if (!input) {
      return;
    }

    if (typeof input.showPicker === "function") {
      input.showPicker();
      return;
    }

    input.focus();
  };

  const applyGlobalDateFilter = async () => {
    const nextAppliedRange = {
      from: normalizeDateInputValue(globalDateRange.from, GLOBAL_DATE_MIN),
      to: normalizeDateInputValue(globalDateRange.to, GLOBAL_DATE_MAX),
    };

    setAppliedGlobalDateRange(nextAppliedRange);

    const filters = buildGlobalDateFilters(selectedReportRaw?.globalDateFilter, nextAppliedRange);
    if (!selectedReportRaw?.globalDateFilter || filters.length === 0) {
      setGlobalDateStatus("Select Contoso or Competitive Marketing to apply this filter.");
      return;
    }

    if (!reportRef.current || typeof reportRef.current.setFilters !== "function") {
      setGlobalDateStatus("Report is still loading. Try Apply again in a moment.");
      return;
    }

    try {
      let currentFilters: models.IFilter[] = [];
      try {
        currentFilters = (await reportRef.current.getFilters?.()) || [];
      } catch (readError) {
        console.warn("Unable to read existing filters before applying global date filter", readError);
      }

      await reportRef.current.setFilters(mergeGlobalDateFilters(currentFilters, filters));

      try {
        const appliedFilters = await reportRef.current.getFilters?.();
        console.log("Applied global date filter from Apply button", {
          requestedFilters: filters,
          appliedFilters,
        });
      } catch (readError) {
        console.warn("Global date filter applied, but unable to read filters afterward", readError);
      }

      setGlobalDateStatus("Date filter applied.");
    } catch (error) {
      console.warn("Unable to apply global date filter from Apply button", error);
      setGlobalDateStatus("Could not apply date filter. Check console for details.");
    }
  };

  const resetGlobalDateFilter = () => {
    const resetRange = { from: GLOBAL_DATE_MIN, to: GLOBAL_DATE_MAX };
    setGlobalDateRange(resetRange);
    setAppliedGlobalDateRange(resetRange);
    setGlobalDateStatus("");
  };

  return isAuthenticated && !authError && !authHookError ? (
    <div className="app-root">
      {/* Header */}
      <header className="app-header">
        <div className="header-left">
          <img src={logo} alt="Logo" className="app-logo" />
          <span className="app-title">Power BI Report Embedding Portal</span>
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
            <h3 className="sidebar-title">Select Reports</h3>
            {reportsToEmbed.map((reportDetail) => (
              <div className="sidebar-list-item" key={reportDetail.reportId}>
                <span className="sidebar-list-title" title={reportDetail.name}>
                  {reportDetail.name}
                </span>
                <button
                  className="mini-btn"
                  onClick={async () => {
                    setSelectedReportRaw(reportDetail);
                    if (reportDetail.isSecureEmbedded) {
                      setTimeout(() => {
                        if (mainContentDivRef.current && reportDetail.embeddingIframe) {
                          mainContentDivRef.current.innerHTML = reportDetail.embeddingIframe.replaceAll("selectLanguage", culture);
                        }
                      }, 1000);
                      return;
                    }
                    const workspace = workspaces.filter((x) => x.id == reportDetail.workspaceId);
                    if (!workspace[0]) {
                      alert("user don't have workspace access");
                      return;
                    }
                    setSelectedWorkspace(workspace[0]);

                    const reportUrl = `${POWER_BI_API_CONST.GROUP_BASE_URL}/${reportDetail.workspaceId}/reports`;

                    const response = await axios.get(reportUrl, {
                      headers: {
                        Authorization: `Bearer ${accessKey}`,
                        "Content-Type": "application/json",
                      },
                    });

                    const reportData: Report[] = response.data?.value;
                    const report = reportData.find((x) => x.id == reportDetail.reportId);
                    if (report) setSelectedReport(report);
                    else alert("report not found");
                  }}
                >
                  Embed
                </button>
              </div>
            ))}
            {/* <input className="sidebar-input" type="text" value={wsFilter} onChange={(e) => setWsFilter(e.target.value)} placeholder="Filter workspace" />
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
                      <button
                        className="mini-btn"
                        onClick={() => {
                          setSelectedWorkspace(w);
                          void dispatch(fetchReports({ workspaceId: w.id, accessToken: accessKey }));
                        }}
                      >
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
            </div> */}
          </div>
          <div className="sidebar-section global-date-filter">
            <h3 className="sidebar-title">Global Date Filter</h3>
            <div className="global-date-fields">
              <label className="global-date-field">
                <span>From</span>
                <div className="date-picker-row">
                  <input
                    ref={fromDateInputRef}
                    className="global-date-input"
                    type="date"
                    min={GLOBAL_DATE_MIN}
                    max={GLOBAL_DATE_MAX}
                    value={globalDateRange.from}
                    onChange={(e) => {
                      setGlobalDateRange((currentRange) => ({
                        ...currentRange,
                        from: normalizeDateInputValue(e.target.value, currentRange.from),
                      }));
                    }}
                  />
                  <button type="button" className="date-picker-button" onClick={() => openDatePicker(fromDateInputRef.current)}>
                    Pick
                  </button>
                </div>
              </label>
              <label className="global-date-field">
                <span>To</span>
                <div className="date-picker-row">
                  <input
                    ref={toDateInputRef}
                    className="global-date-input"
                    type="date"
                    min={GLOBAL_DATE_MIN}
                    max={GLOBAL_DATE_MAX}
                    value={globalDateRange.to}
                    onChange={(e) => {
                      setGlobalDateRange((currentRange) => ({
                        ...currentRange,
                        to: normalizeDateInputValue(e.target.value, currentRange.to),
                      }));
                    }}
                  />
                  <button type="button" className="date-picker-button" onClick={() => openDatePicker(toDateInputRef.current)}>
                    Pick
                  </button>
                </div>
              </label>
            </div>
            <div className="global-date-actions">
              <button type="button" className="mini-btn" onClick={() => void applyGlobalDateFilter()}>
                Apply
              </button>
              <button type="button" className="mini-btn" onClick={resetGlobalDateFilter}>
                Reset
              </button>
            </div>
            {selectedReportRaw?.globalDateFilter ? (
              <p className="global-date-filter-status">
                Filtering {selectedReportRaw.globalDateFilter.table}.{selectedReportRaw.globalDateFilter.column}
              </p>
            ) : (
              <p className="global-date-filter-status">Applies to Contoso and Competitive Marketing reports.</p>
            )}
            {globalDateStatus && <p className="global-date-filter-status">{globalDateStatus}</p>}
          </div>
        </aside>

        {/* Main content area */}
        <main className="main-content">
          {selectedReportRaw?.isSecureEmbedded && (
            <div className="main-content-child">
              <div className="select-culture-div">
                <label>Select Culture</label>
                <select
                  value={culture}
                  onChange={(e) => {
                    setCulture(e.target.value);
                    setTimeout(() => {
                      if (mainContentDivRef.current && selectedReportRaw.embeddingIframe) {
                        mainContentDivRef.current.innerHTML = selectedReportRaw.embeddingIframe.replaceAll("selectLanguage", e.target.value);
                      }
                    }, 1000);
                  }}
                >
                  <option value="en-US">English</option>
                  <option value="fr-FR">French</option>
                  <option value="es-ES">Spanish</option>
                </select>
              </div>
              <div className="main-content-child-inner" ref={mainContentDivRef}></div>
            </div>
          )}
          {selectedReportRaw?.isSecureEmbedded == false ? (
            selectedReport && selectedWorkspace ? (
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
                            <option value="" disabled>
                              Select bookmark
                            </option>
                            {bookmarks.map((b) => (
                              <option key={b.id} value={b.id}>
                                {b.name}
                              </option>
                            ))}
                          </select>
                          <button
                            title="Delete bookmark"
                            disabled={!selectedBookmarkId}
                            style={{
                              width: 34,
                              height: 34,
                              border: "1.5px solid #e57373",
                              borderRadius: 6,
                              background: "#fff",
                              color: "#e53935",
                              cursor: selectedBookmarkId ? "pointer" : "not-allowed",
                              fontSize: "1rem",
                              display: "flex",
                              alignItems: "center",
                              justifyContent: "center",
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
                              width: 34,
                              height: 34,
                              border: "1.5px solid #43a047",
                              borderRadius: 6,
                              background: "#fff",
                              color: "#2e7d32",
                              cursor: selectedBookmarkId ? "pointer" : "not-allowed",
                              fontSize: "1.1rem",
                              display: "flex",
                              alignItems: "center",
                              justifyContent: "center",
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
                      <button className="primary-btn" onClick={() => setMode("viewer")}>
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
                      reportFilters={globalDateFilters}
                      accessToken={accessKey}
                      userId={user?.id || user?.userPrincipalName || user?.email || ""}
                      workspaceId={selectedWorkspace.id}
                      allowEdit={!selectedWorkspace.isReadOnly}
                      onReportLoadReportAttachmentFunction={(report: any) => {
                        addCustomFunctionsInEmbeddedReport(report);
                        reportRef.current = report;
                      }}
                      onReportReady={(report, page) => {
                        setLayoutReport((prev: any) => (prev !== report ? report : prev));
                        setLayoutPage((prev: any) => (prev !== page ? page : prev));
                      }}
                      toggleButton={
                        <button className="primary-btn" onClick={() => setMode("creator")}>
                          Quick Visual Creator
                        </button>
                      }
                      layoutControls={
                        layoutReport && layoutPage ? <LayoutCustomizer ref={layoutCustomizerRef} report={layoutReport} page={layoutPage} /> : null
                      }
                      layoutCustomizerRef={layoutCustomizerRef}
                    />
                  </>
                )}
              </>
            ) : (
              <div className="empty-state">Select a report to embed</div>
            )
          ) : null}
          {!selectedReportRaw && <div className="empty-state">Select a report to embed</div>}
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
  );
}

export default App;
