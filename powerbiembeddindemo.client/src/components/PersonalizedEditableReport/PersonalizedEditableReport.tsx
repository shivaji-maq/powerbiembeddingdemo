import React, { useRef, useState, useEffect, useCallback, useMemo } from "react";
import { models } from "powerbi-client";
import type { EventHandler } from "powerbi-client-react";
import EmbedReport from "../EmbeddedReport/EmbeddedReport";
import { ReportEditor } from "../ReportEditor/ReportEditor";
import { usePersonalization } from "../../hooks/usePersonalization";
import { applyPersonalizedFilters, getReportPersonalizationState} from "../../lib/powerbiLib/personalization";
// @ts-ignore -- CSS side-effect imports are handled by Vite at runtime.
import "./PersonalizedEditableReport.css";

interface PersonalizedEditableReportProps {
  reportId: string;
  embedUrl: string;
  workspaceId: string;
  userId: string;
  accessToken: string;
  tokenType?: "Aad" | "Embed";
  allowEdit?: boolean;
  reportRef?: React.MutableRefObject<any>;
  embedReportEventHandlers?: Map<string, EventHandler>;
  reportSettings?: any;
  onReportLoadReportAttachmentFunction?: (report: any) => void;
  toggleButton?: React.ReactNode;
  layoutControls?: React.ReactNode;
  onReportReady?: (report: any, page: any) => void;
  layoutCustomizerRef?: React.RefObject<any>;
}

interface BookmarkProfile {
  id: string;
  name: string;
  state?: string;
  filtersJson?: string;
  activePage?: string;
  bookmarkStateJson?: string;
  createdAt: string;
  updatedAt: string;
  layoutState?: any;
}

const SAVED_BOOKMARK_PREFIX = "saved:";
const REPORT_BOOKMARK_PREFIX = "report:";
const ORIGINAL_REPORT_SELECTION_ID = "original:view";

type QuickVisualMode = "create" | "change";
type QuickVisualProperty =
  | "title"
  | "xAxis"
  | "yAxis"
  | "legend"
  | "titleText"
  | "titleAlign"
  | "titleSize"
  | "titleColor";

interface QuickVisualOption {
  name: string;
  label: string;
  dataRoleNames: string[];
  properties: Array<"legend" | "xAxis" | "yAxis">;
}

interface QuickVisualTarget {
  name: string;
  title: string;
  type: string;
  layout?: any;
}

interface QuickVisualFieldOption {
  key: string;
  label: string;
  target: any;
}

interface SessionQuickVisualLayout {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface SessionQuickVisualProperties {
  showTitle: boolean;
  showLegend: boolean;
  showXAxis: boolean;
  showYAxis: boolean;
  titleText: string;
  titleAlign: string;
}

interface SessionQuickVisualProfile {
  id: string;
  visualType: string;
  roleTargets: Record<string, any>;
  layout: SessionQuickVisualLayout;
  properties: SessionQuickVisualProperties;
  runtimeVisualName?: string;
  createdAt: string;
}

type QuickVisualRoleSelections = Record<string, string>;
type CapturedBookmarkUpsertMode = "saveView" | "syncApplied";

const QUICK_VISUAL_SCHEMAS = {
  property: "http://powerbi.com/product/schema#property",
};

const QUICK_VISUAL_ROLE_LABELS: Record<string, string> = {
  Category: "X-axis",
  Y: "Y-axis",
  Series: "Legend",
  Tooltips: "Tooltip",
  Values: "Values",
  Axis: "Axis",
  Legend: "Legend",
};

const QUICK_VISUAL_MIN_REQUIRED_FIELDS = 2;

const QUICK_VISUAL_OPTIONS: QuickVisualOption[] = [
  {
    name: "columnChart",
    label: "Column chart",
    dataRoleNames: ["Category", "Y", "Tooltips"],
    properties: ["xAxis", "yAxis"],
  },
  {
    name: "areaChart",
    label: "Area chart",
    dataRoleNames: ["Category", "Series", "Y"],
    properties: ["legend", "xAxis", "yAxis"],
  },
  {
    name: "barChart",
    label: "Bar chart",
    dataRoleNames: ["Category", "Y", "Tooltips"],
    properties: ["xAxis", "yAxis"],
  },
  {
    name: "pieChart",
    label: "Pie chart",
    dataRoleNames: ["Category", "Y", "Tooltips"],
    properties: ["legend"],
  },
  {
    name: "lineChart",
    label: "Line chart",
    dataRoleNames: ["Category", "Series", "Y"],
    properties: ["legend", "xAxis", "yAxis"],
  },
];

const QUICK_VISUAL_DEFAULT_ROLE_NAMES = Array.from(
  new Set(QUICK_VISUAL_OPTIONS.flatMap((option) => option.dataRoleNames))
);

const getQuickVisualOption = (visualType: string): QuickVisualOption => {
  return (
    QUICK_VISUAL_OPTIONS.find((option) => option.name === visualType) ||
    QUICK_VISUAL_OPTIONS[0]
  );
};

const toQuickVisualPropertySelector = (propertyName: QuickVisualProperty) => {
  switch (propertyName) {
    case "title":
      return { objectName: "title", propertyName: "visible" };
    case "xAxis":
      return { objectName: "categoryAxis", propertyName: "visible" };
    case "yAxis":
      return { objectName: "valueAxis", propertyName: "visible" };
    case "legend":
      return { objectName: "legend", propertyName: "visible" };
    case "titleText":
      return { objectName: "title", propertyName: "titleText" };
    case "titleAlign":
      return { objectName: "title", propertyName: "alignment" };
    case "titleSize":
      return { objectName: "title", propertyName: "textSize" };
    case "titleColor":
      return { objectName: "title", propertyName: "fontColor" };
    default:
      return null;
  }
};

const getQuickVisualFieldKey = (field: any) => {
  if (!field || typeof field !== "object") {
    return "";
  }

  if (typeof field.table === "string" && typeof field.column === "string") {
    return `${field.table}|column|${field.column}`;
  }

  if (typeof field.table === "string" && typeof field.measure === "string") {
    return `${field.table}|measure|${field.measure}`;
  }

  if (
    typeof field.table === "string" &&
    typeof field.hierarchy === "string" &&
    typeof field.hierarchyLevel === "string"
  ) {
    return `${field.table}|hierarchy|${field.hierarchy}|${field.hierarchyLevel}`;
  }

  try {
    return JSON.stringify(field);
  } catch {
    return "";
  }
};

const getQuickVisualFieldLabel = (field: any) => {
  if (!field || typeof field !== "object") {
    return "Unknown field";
  }

  if (field.label) {
    return String(field.label);
  }

  if (field.displayName) {
    return String(field.displayName);
  }

  if (field.measure) {
    return String(field.measure);
  }

  if (field.column) {
    return String(field.column);
  }

  if (field.hierarchyLevel) {
    return String(field.hierarchyLevel);
  }

  if (field.queryName) {
    return String(field.queryName);
  }

  return "Field";
};

const parseQueryNameToFieldTarget = (
  queryName: string
): { table: string; column: string } | null => {
  const trimmed = String(queryName || "").trim();
  if (!trimmed) {
    return null;
  }

  const bracketMatch = trimmed.match(/^'?([^'\[]+)'?\[(.+)\]$/);
  if (bracketMatch) {
    return {
      table: bracketMatch[1].trim(),
      column: bracketMatch[2].trim(),
    };
  }

  const dotMatch = trimmed.match(/^'?([^'.]+)'?\.(.+)$/);
  if (dotMatch) {
    const table = dotMatch[1].trim();
    const fieldName = dotMatch[2].trim().split(".").pop() || "";
    if (table && fieldName) {
      return {
        table,
        column: fieldName,
      };
    }
  }

  return null;
};

const normalizeQuickVisualFieldTarget = (
  field: any
): {
  table: string;
  column?: string;
  measure?: string;
  hierarchy?: string;
  hierarchyLevel?: string;
} | null => {
  if (!field || typeof field !== "object") {
    return null;
  }

  if (field.target && typeof field.target === "object") {
    const normalizedTarget = normalizeQuickVisualFieldTarget(field.target);
    if (normalizedTarget) {
      return normalizedTarget;
    }
  }

  const fromQueryName =
    typeof field.queryName === "string"
      ? parseQueryNameToFieldTarget(field.queryName)
      : null;

  const entityName =
    field?.entity ||
    field?.Entity ||
    field?.tableName ||
    field?.source?.entity ||
    field?.sourceRef?.entity ||
    field?.SourceRef?.Entity ||
    field?.expression?.source?.entity ||
    field?.expression?.sourceRef?.entity ||
    field?.Expression?.SourceRef?.Entity ||
    field?.Column?.Expression?.SourceRef?.Entity ||
    field?.Measure?.Expression?.SourceRef?.Entity;

  const propertyName =
    field?.property ||
    field?.Property ||
    field?.Column?.Property ||
    field?.Measure?.Property ||
    field?.name ||
    field?.displayName;

  if (typeof entityName === "string" && typeof propertyName === "string") {
    return {
      table: entityName,
      column: propertyName,
    };
  }

  if (typeof entityName === "string" && typeof field.column === "string") {
    return {
      table: entityName,
      column: field.column,
    };
  }

  if (typeof entityName === "string" && typeof field.measure === "string") {
    return {
      table: entityName,
      measure: field.measure,
    };
  }

  if (typeof field.table !== "string") {
    if (fromQueryName) {
      return fromQueryName;
    }
    return null;
  }

  if (typeof field.column === "string") {
    return {
      table: field.table,
      column: field.column,
    };
  }

  if (typeof field.measure === "string") {
    return {
      table: field.table,
      measure: field.measure,
    };
  }

  if (
    typeof field.hierarchy === "string" &&
    typeof field.hierarchyLevel === "string"
  ) {
    return {
      table: field.table,
      hierarchy: field.hierarchy,
      hierarchyLevel: field.hierarchyLevel,
    };
  }

  if (fromQueryName) {
    return fromQueryName;
  }

  return null;
};

const collectQuickVisualFieldTargets = (
  value: any,
  onTarget: (target: any) => void,
  depth = 0,
  visited: WeakSet<object> = new WeakSet<object>()
) => {
  if (value == null || depth > 8) {
    return;
  }

  if (Array.isArray(value)) {
    value.forEach((item) => {
      collectQuickVisualFieldTargets(item, onTarget, depth + 1, visited);
    });
    return;
  }

  if (typeof value !== "object") {
    return;
  }

  const normalized = normalizeQuickVisualFieldTarget(value);
  if (normalized) {
    onTarget(normalized);
  }

  if (visited.has(value)) {
    return;
  }
  visited.add(value);

  Object.values(value).forEach((item) => {
    collectQuickVisualFieldTargets(item, onTarget, depth + 1, visited);
  });
};

const addQuickVisualFieldOption = (
  optionByKey: Map<string, QuickVisualFieldOption>,
  rawField: any
) => {
  const normalizedField = normalizeQuickVisualFieldTarget(rawField);
  if (!normalizedField) {
    return;
  }

  const key = getQuickVisualFieldKey(normalizedField);
  if (!key || optionByKey.has(key)) {
    return;
  }

  optionByKey.set(key, {
    key,
    label:
      getQuickVisualFieldLabel(normalizedField) ||
      getQuickVisualFieldLabel(rawField),
    target: normalizedField,
  });
};

const createBookmarkId = () => {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }

  return `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
};

const toSavedBookmarkSelectionId = (bookmarkId: string) =>
  `${SAVED_BOOKMARK_PREFIX}${bookmarkId}`;

const toReportBookmarkSelectionId = (bookmarkName: string) =>
  `${REPORT_BOOKMARK_PREFIX}${bookmarkName}`;

const getSavedBookmarkIdFromSelection = (selectionId: string) =>
  selectionId.startsWith(SAVED_BOOKMARK_PREFIX)
    ? selectionId.slice(SAVED_BOOKMARK_PREFIX.length)
    : "";

const getReportBookmarkNameFromSelection = (selectionId: string) =>
  selectionId.startsWith(REPORT_BOOKMARK_PREFIX)
    ? selectionId.slice(REPORT_BOOKMARK_PREFIX.length)
    : "";

const isOriginalReportSelection = (selectionId: string) =>
  selectionId === ORIGINAL_REPORT_SELECTION_ID;

const parseSafeDate = (value?: string | null) => {
  if (!value) {
    return null;
  }

  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const flattenReportBookmarks = (
  bookmarks: models.IReportBookmark[]
): models.IReportBookmark[] => {
  const flattened: models.IReportBookmark[] = [];
  const stack = Array.isArray(bookmarks) ? [...bookmarks] : [];

  while (stack.length > 0) {
    const bookmark = stack.shift();
    if (!bookmark || typeof bookmark.name !== "string") {
      continue;
    }

    const children = Array.isArray(bookmark.children) ? bookmark.children : [];
    const canApplyBookmark =
      (typeof bookmark.state === "string" && bookmark.state.length > 0) ||
      children.length === 0;

    if (canApplyBookmark) {
      flattened.push({
        name: bookmark.name,
        displayName: bookmark.displayName || bookmark.name,
        state: bookmark.state,
      });
    }

    if (children.length > 0) {
      stack.unshift(...children);
    }
  }

  const deduped = new Map<string, models.IReportBookmark>();
  flattened.forEach((bookmark) => {
    if (!deduped.has(bookmark.name)) {
      deduped.set(bookmark.name, bookmark);
    }
  });

  return Array.from(deduped.values());
};

const getBookmarkNameFromAppliedEvent = (event: any) => {
  const candidates = [
    event?.detail?.bookmark?.displayName,
    event?.detail?.bookmark?.displayname,
    event?.detail?.bookmark?.display_name,
    event?.detail?.bookmark?.title,
    event?.detail?.bookmark?.name,
    event?.detail?.bookmark?.bookmarkName,
    event?.detail?.bookmarkName,
    event?.detail?.displayName,
    event?.detail?.name,
    event?.detail?.title,
  ];

  const match = candidates.find(
    (value) => typeof value === "string" && value.trim()
  );

  return typeof match === "string" ? match.trim() : "";
};

const isLikelyOpaqueBookmarkIdentifier = (value: string) => {
  const trimmed = value.trim();
  if (!trimmed) {
    return false;
  }

  const looksLikeGuid =
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(
      trimmed
    );
  
  return looksLikeGuid;
};

const toReadableBookmarkName = (
  ...candidates: Array<string | null | undefined>
) => {
  for (const candidate of candidates) {
    if (typeof candidate !== "string") {
      continue;
    }

    const trimmed = candidate.trim();
    if (!trimmed || isLikelyOpaqueBookmarkIdentifier(trimmed)) {
      continue;
    }

    return trimmed;
  }

  const opaqueCandidate = candidates.find(
    (candidate) => typeof candidate === "string" && candidate.trim()
  );

  if (typeof opaqueCandidate === "string") {
    const compact = opaqueCandidate.trim().replace(/-/g, "").slice(0, 6);
    return compact ? `Personal bookmark ${compact}` : "Personal bookmark";
  }

  return "Personal bookmark";
};

const toUniqueBookmarkName = (
  preferredName: string,
  existingBookmarks: BookmarkProfile[],
  preserveBookmarkId?: string
) => {
  const baseName = preferredName.trim() || "Saved view";
  const normalize = (value: string) => value.trim().toLowerCase();
  const existingNames = new Set(
    existingBookmarks
      .filter((bookmark) => bookmark.id !== preserveBookmarkId)
      .map((bookmark) => normalize(bookmark.name))
  );

  if (!existingNames.has(normalize(baseName))) {
    return baseName;
  }

  let suffix = 2;
  while (existingNames.has(normalize(`${baseName} (${suffix})`))) {
    suffix += 1;
  }

  return `${baseName} (${suffix})`;
};

const toNormalizedBookmarkProfile = (bookmark: any): BookmarkProfile | null => {
  if (!bookmark || typeof bookmark !== "object") {
    return null;
  }

  const id =
    typeof bookmark.id === "string" && bookmark.id
      ? bookmark.id
      : createBookmarkId();

  const rawBookmarkName =
    typeof bookmark.name === "string" ? bookmark.name.trim() : "";
  const bookmarkName = rawBookmarkName
    ? toReadableBookmarkName(rawBookmarkName)
    : `Saved view ${id.slice(0, 6)}`;

  const bookmarkState =
    typeof bookmark.state === "string"
      ? bookmark.state
      : typeof bookmark.bookmarkStateJson === "string"
        ? bookmark.bookmarkStateJson
        : "";

  if (!bookmarkState) {
    return null;
  }

  return {
    ...bookmark,
    id,
    name: bookmarkName,
    state: bookmarkState,
    bookmarkStateJson: bookmarkState,
    createdAt:
      typeof bookmark.createdAt === "string"
        ? bookmark.createdAt
        : new Date().toISOString(),
    updatedAt:
      typeof bookmark.updatedAt === "string"
        ? bookmark.updatedAt
        : new Date().toISOString(),
  } as BookmarkProfile;
};

const normalizeAndDedupeBookmarkProfiles = (bookmarks: any): BookmarkProfile[] => {
  if (!Array.isArray(bookmarks)) {
    return [];
  }

  const normalizedBookmarks = bookmarks
    .map((bookmark) => toNormalizedBookmarkProfile(bookmark))
    .filter(Boolean) as BookmarkProfile[];

  const sortedByRecentUpdate = [...normalizedBookmarks].sort((a, b) => {
    const left = parseSafeDate(a.updatedAt)?.getTime() ?? 0;
    const right = parseSafeDate(b.updatedAt)?.getTime() ?? 0;
    return right - left;
  });

  const seenState = new Set<string>();
  const seenName = new Set<string>();
  const deduped: BookmarkProfile[] = [];

  sortedByRecentUpdate.forEach((bookmark) => {
    const stateKey = (bookmark.state || bookmark.bookmarkStateJson || "").trim();
    const nameKey = bookmark.name.trim().toLowerCase();

    if (!stateKey) {
      return;
    }

    if (seenState.has(stateKey)) {
      return;
    }

    if (nameKey && seenName.has(nameKey)) {
      return;
    }

    seenState.add(stateKey);
    if (nameKey) {
      seenName.add(nameKey);
    }

    deduped.push({
      ...bookmark,
      state: stateKey,
      bookmarkStateJson: stateKey,
    });
  });

  return deduped;
};

const wait = (ms: number) =>
  new Promise<void>((resolve) => window.setTimeout(resolve, ms));

const isReportNotReadyError = (error: unknown) => {
  const message = String((error as any)?.message || "").toLowerCase();
  const detailed = String((error as any)?.detailedMessage || "").toLowerCase();
  return (
    message.includes("reportisnotready") ||
    detailed.includes("report is not ready")
  );
};

export const PersonalizedEditableReport: React.FC<
  PersonalizedEditableReportProps
> = ({
  reportId,
  embedUrl,
  workspaceId,
  userId,
  accessToken,
  tokenType,
  allowEdit = true,
  reportRef: externalReportRef,
  embedReportEventHandlers,
  reportSettings: customReportSettings,
  onReportLoadReportAttachmentFunction,
  toggleButton,
  layoutControls,
  onReportReady,
  layoutCustomizerRef,
}) => {
  const internalReportRef = useRef<any>(null);
  const authoringReportRef = useRef<any>(null);
  // FIX: Track the dedicated blank authoring page separately (showcase uses pages[1])
  const authoringPageRef = useRef<any>(null);
  const reportRef = externalReportRef || internalReportRef;
  const customReportSettingsObject =
    (customReportSettings || {}) as Record<string, any>;
  const customQuickVisualFieldCatalog = Array.isArray(
    customReportSettingsObject.quickVisualFieldCatalog
  )
    ? customReportSettingsObject.quickVisualFieldCatalog
    : [];
  const embedReportSettingsOverrides = { ...customReportSettingsObject };
  delete embedReportSettingsOverrides.quickVisualFieldCatalog;

  const [_currentPage, setCurrentPage] = useState<string>("");
  const [autoSaveEnabled, setAutoSaveEnabled] = useState(true);
  const [lastSaved, setLastSaved] = useState<Date | null>(null);
  const [saveStatus, setSaveStatus] = useState<
    "idle" | "saving" | "saved" | "error"
  >("idle");
  const [isReportLoaded, setIsReportLoaded] = useState(false);
  const [isAuthoringReportLoaded, setIsAuthoringReportLoaded] = useState(false);
  const [authoringEmbedError, setAuthoringEmbedError] = useState<string | null>(
    null
  );
  const [autoSaveRevision, setAutoSaveRevision] = useState(0);
  const [bookmarkProfiles, setBookmarkProfiles] = useState<BookmarkProfile[]>(
    []
  );
  const [reportBookmarks, setReportBookmarks] = useState<
    models.IReportBookmark[]
  >([]);
  const [selectedBookmarkId, setSelectedBookmarkId] = useState<string>("");
  const [bookmarkStatus, setBookmarkStatus] = useState<string | null>(null);
  const [isBookmarkModalOpen, setIsBookmarkModalOpen] = useState(false);
  const [bookmarkNameInput, setBookmarkNameInput] = useState("");
  const [isSavingBookmark, setIsSavingBookmark] = useState(false);
  const [isQuickVisualModalOpen, setIsQuickVisualModalOpen] = useState(false);
  const [quickVisualMode, setQuickVisualMode] =
    useState<QuickVisualMode>("create");
  const [quickVisualType, setQuickVisualType] = useState("columnChart");
  const [quickVisualTargetVisualName, setQuickVisualTargetVisualName] =
    useState("");
  const [quickVisualFieldOptions, setQuickVisualFieldOptions] = useState<
    QuickVisualFieldOption[]
  >([]);
  const [quickVisualRoleSelections, setQuickVisualRoleSelections] =
    useState<QuickVisualRoleSelections>({});
  const [quickVisualTitle, setQuickVisualTitle] = useState("");
  const [quickVisualTitleAlign, setQuickVisualTitleAlign] = useState(
    "left"
  );
  const [quickVisualShowTitle, setQuickVisualShowTitle] = useState(true);
  const [quickVisualShowLegend, setQuickVisualShowLegend] = useState(true);
  const [quickVisualShowXAxis, setQuickVisualShowXAxis] = useState(true);
  const [quickVisualShowYAxis, setQuickVisualShowYAxis] = useState(true);
  const [isQuickVisualApplying, setIsQuickVisualApplying] = useState(false);
  const [quickVisualTargets, setQuickVisualTargets] = useState<
    QuickVisualTarget[]
  >([]);
  const [quickVisualStatus, setQuickVisualStatus] = useState<string | null>(
    null
  );
  const [isHydratingPersonalization, setIsHydratingPersonalization] =
    useState(true);
  const autoSaveInFlightRef = useRef(false);
  const suppressAutoSaveEventsRef = useRef(0);
  const hasHydratedRef = useRef(false);
  const quickVisualSessionSyncTimerRef = useRef<number | null>(null);
  const bookmarkStatusTimerRef = useRef<number | null>(null);
  const quickVisualStatusTimerRef = useRef<number | null>(null);
  // FIX: Track the in-progress preview visual on the authoring blank page
  const authoringPreviewVisualRef = useRef<any>(null);
  const bookmarksStorageKey = `pbi_bookmarks_${userId}_${reportId}`;
  const selectedBookmarkStorageKey = `${bookmarksStorageKey}_selected`;
  const quickVisualFieldCatalogStorageKey =
    `pbi_quick_visual_fields_${reportId}`;
  const quickVisualSessionStorageKey =
    `pbi_quick_visual_session_${userId}_${reportId}`;
  const originalReportStateStorageKey =
    `pbi_original_state_${userId}_${reportId}`;

  const { savePersonalization, getPersonalization, loading } =
    usePersonalization();
  const isReportLoadedRef = useRef(isReportLoaded);
  const autoSaveEnabledRef = useRef(autoSaveEnabled);
  const isHydratingPersonalizationRef = useRef(isHydratingPersonalization);
  const getCurrentPersonalizationPayloadRef =
    useRef<() => Promise<any>>(async () => null);
  const savePersonalizationRef = useRef(savePersonalization);
  const bookmarkProfilesRef = useRef<BookmarkProfile[]>(bookmarkProfiles);
  const hasReplayedSessionQuickVisualsRef = useRef(false);
  const originalReportStateRef = useRef("");

  useEffect(() => {
    isReportLoadedRef.current = isReportLoaded;
  }, [isReportLoaded]);

  useEffect(() => {
    autoSaveEnabledRef.current = autoSaveEnabled;
  }, [autoSaveEnabled]);

  useEffect(() => {
    isHydratingPersonalizationRef.current = isHydratingPersonalization;
  }, [isHydratingPersonalization]);

  useEffect(() => {
    savePersonalizationRef.current = savePersonalization;
  }, [savePersonalization]);

  useEffect(() => {
    bookmarkProfilesRef.current = bookmarkProfiles;
  }, [bookmarkProfiles]);

  const showBookmarkStatus = useCallback((message: string | null) => {
    if (bookmarkStatusTimerRef.current) {
      window.clearTimeout(bookmarkStatusTimerRef.current);
      bookmarkStatusTimerRef.current = null;
    }

    setBookmarkStatus(message);
    if (message) {
      bookmarkStatusTimerRef.current = window.setTimeout(() => {
        setBookmarkStatus(null);
      }, 3000);
    }
  }, []);

  const showQuickVisualStatus = useCallback((message: string | null) => {
    if (quickVisualStatusTimerRef.current) {
      window.clearTimeout(quickVisualStatusTimerRef.current);
      quickVisualStatusTimerRef.current = null;
    }

    setQuickVisualStatus(message);
    if (message) {
      quickVisualStatusTimerRef.current = window.setTimeout(() => {
        setQuickVisualStatus(null);
      }, 4000);
    }
  }, []);

  const toAuthoringEmbedErrorMessage = useCallback((event?: any) => {
    const candidateMessages = [
      event?.detail?.message,
      event?.detail?.detailedMessage,
      event?.detail?.error?.message,
      event?.message,
    ];

    const rawMessage = candidateMessages.find(
      (candidate) => typeof candidate === "string" && candidate.trim()
    );

    const message =
      typeof rawMessage === "string"
        ? rawMessage.trim()
        : "Authoring report failed to initialize.";

    const normalizedMessage = message.toLowerCase();
    if (
      normalizedMessage.includes("insufficient") ||
      normalizedMessage.includes("forbidden") ||
      normalizedMessage.includes("access denied") ||
      normalizedMessage.includes("unauthorized") ||
      normalizedMessage.includes("403")
    ) {
      return "Authoring report failed to initialize with edit access. Ensure embed token allows edit and the user has Member/Contributor/Admin role.";
    }

    return message;
  }, []);

  const getOriginalReportState = useCallback(() => {
    if (originalReportStateRef.current) {
      return originalReportStateRef.current;
    }

    try {
      const storedState =
        window.sessionStorage.getItem(originalReportStateStorageKey) || "";
      if (storedState) {
        originalReportStateRef.current = storedState;
      }
      return storedState;
    } catch {
      return "";
    }
  }, [originalReportStateStorageKey]);

  const captureOriginalReportStateIfMissing = useCallback(async () => {
    const existingOriginalState = getOriginalReportState();
    if (existingOriginalState) {
      return true;
    }

    if (
      !reportRef.current ||
      typeof reportRef.current?.bookmarksManager?.capture !== "function"
    ) {
      return false;
    }

    try {
      const captured = await reportRef.current.bookmarksManager.capture({
        personalizeVisuals: true,
      });
      const bookmarkState = captured?.state || "";
      if (!bookmarkState) {
        return false;
      }

      originalReportStateRef.current = bookmarkState;
      try {
        window.sessionStorage.setItem(
          originalReportStateStorageKey,
          bookmarkState
        );
      } catch { }

      return true;
    } catch (error) {
      console.warn("Unable to capture original report state", error);
      return false;
    }
  }, [getOriginalReportState, originalReportStateStorageKey, reportRef]);

  const restoreOriginalReportState = useCallback(async () => {
    if (!reportRef.current) {
      return false;
    }

    const originalState = getOriginalReportState();

    try {
      if (
        originalState &&
        typeof reportRef.current?.bookmarksManager?.applyState === "function"
      ) {
        suppressAutoSaveEventsRef.current = Math.max(
          suppressAutoSaveEventsRef.current,
          2
        );
        await reportRef.current.bookmarksManager.applyState(originalState);
      } else if (typeof reportRef.current.reload === "function") {
        await reportRef.current.reload();
      } else {
        return false;
      }

      const activePage = await reportRef.current?.getActivePage?.();
      if (activePage?.name) {
        setCurrentPage(activePage.name);
      }

      if (
        reportRef.current &&
        authoringReportRef.current &&
        typeof reportRef.current?.bookmarksManager?.capture === "function" &&
        typeof authoringReportRef.current?.bookmarksManager?.applyState ===
          "function"
      ) {
        try {
          const captured = await reportRef.current.bookmarksManager.capture({
            personalizeVisuals: true,
          });
          const bookmarkState = captured?.state || "";
          if (bookmarkState) {
            await authoringReportRef.current.bookmarksManager.applyState(
              bookmarkState
            );
          }
        } catch (syncError) {
          console.warn(
            "Unable to sync visible report state into authoring report",
            syncError
          );
        }
      }

      return true;
    } catch (error) {
      console.warn("Unable to restore original report state", error);
      return false;
    }
  }, [getOriginalReportState, reportRef]);

  const syncVisibleStateToAuthoring = useCallback(async () => {
    if (!reportRef.current || !authoringReportRef.current) {
      return false;
    }

    if (
      typeof reportRef.current?.bookmarksManager?.capture !== "function" ||
      typeof authoringReportRef.current?.bookmarksManager?.applyState !==
        "function"
    ) {
      return false;
    }

    try {
      const captured = await reportRef.current.bookmarksManager.capture({
        personalizeVisuals: true,
      });
      const bookmarkState = captured?.state || "";

      if (!bookmarkState) {
        return false;
      }

      await authoringReportRef.current.bookmarksManager.applyState(bookmarkState);
      return true;
    } catch (error) {
      console.warn("Unable to sync visible report state into authoring report", error);
      return false;
    }
  }, [reportRef]);

  const syncAuthoringStateToVisible = useCallback(async () => {
    if (!reportRef.current || !authoringReportRef.current) {
      return false;
    }

    if (
      typeof authoringReportRef.current?.bookmarksManager?.capture !== "function" ||
      typeof reportRef.current?.bookmarksManager?.applyState !== "function"
    ) {
      return false;
    }

    try {
      const captured = await authoringReportRef.current.bookmarksManager.capture({
        personalizeVisuals: true,
      });
      const bookmarkState = captured?.state || "";

      if (!bookmarkState) {
        return false;
      }

      suppressAutoSaveEventsRef.current = Math.max(
        suppressAutoSaveEventsRef.current,
        2
      );
      await reportRef.current.bookmarksManager.applyState(bookmarkState);

      const activePage = await reportRef.current?.getActivePage?.();
      if (activePage?.name) {
        setCurrentPage(activePage.name);
      }

      return true;
    } catch (error) {
      console.warn("Unable to sync authoring report state into visible report", error);
      return false;
    }
  }, [reportRef]);

  useEffect(() => {
    hasHydratedRef.current = false;
    hasReplayedSessionQuickVisualsRef.current = false;
    originalReportStateRef.current = "";
    authoringReportRef.current = null;
    // FIX: Reset authoring page ref on report/user change
    authoringPageRef.current = null;
    authoringPreviewVisualRef.current = null;
    setIsReportLoaded(false);
    setIsAuthoringReportLoaded(false);
    setAuthoringEmbedError(null);
    setIsHydratingPersonalization(true);
    setSaveStatus("idle");
    setBookmarkStatus(null);
    setReportBookmarks([]);
    setIsBookmarkModalOpen(false);
    setBookmarkNameInput("");
    setIsQuickVisualModalOpen(false);
    setQuickVisualMode("create");
    setQuickVisualTitle("");
    setQuickVisualType("columnChart");
    setQuickVisualTargetVisualName("");
    setQuickVisualTargets([]);
    setQuickVisualFieldOptions([]);
    setQuickVisualRoleSelections({});
    setQuickVisualShowTitle(true);
    setQuickVisualShowLegend(true);
    setQuickVisualShowXAxis(true);
    setQuickVisualShowYAxis(true);
    setQuickVisualStatus(null);
    setAutoSaveRevision(0);
    setCurrentPage("");
  }, [reportId, userId]);

  useEffect(() => {
    return () => {
      if (bookmarkStatusTimerRef.current) {
        window.clearTimeout(bookmarkStatusTimerRef.current);
      }
      if (quickVisualStatusTimerRef.current) {
        window.clearTimeout(quickVisualStatusTimerRef.current);
      }
      if (quickVisualSessionSyncTimerRef.current) {
        window.clearTimeout(quickVisualSessionSyncTimerRef.current);
      }
    };
  }, []);

  useEffect(() => {
    if (
      !isReportLoaded ||
      !isAuthoringReportLoaded ||
      isHydratingPersonalization
    ) {
      return;
    }

    void syncVisibleStateToAuthoring();
  }, [
    isAuthoringReportLoaded,
    isHydratingPersonalization,
    isReportLoaded,
    syncVisibleStateToAuthoring,
  ]);

  const persistBookmarks = useCallback(
    (nextBookmarks: BookmarkProfile[]) => {
      const finalBookmarksToSave = normalizeAndDedupeBookmarkProfiles(
        nextBookmarks
      );

      setBookmarkProfiles(finalBookmarksToSave);
      bookmarkProfilesRef.current = finalBookmarksToSave;
      window.localStorage.setItem(
        bookmarksStorageKey,
        JSON.stringify(finalBookmarksToSave)
      );
    },
    [bookmarksStorageKey]
  );

  const upsertCapturedBookmark = useCallback(
    (
      bookmarkName: string,
      bookmarkStateJson: string,
      selectAfterUpsert = false,
      mode: CapturedBookmarkUpsertMode = "syncApplied",
      layoutState?: any
    ) => {
      if (!bookmarkStateJson) {
        return null;
      }

      const existingBookmarks = bookmarkProfilesRef.current;

      const normalizedInputName =
        typeof bookmarkName === "string" ? bookmarkName.trim() : "";
      const canMatchByName =
        !!normalizedInputName &&
        !isLikelyOpaqueBookmarkIdentifier(normalizedInputName);

      const existingBookmarkByState = existingBookmarks.find(
        (bookmark) =>
          bookmark.state === bookmarkStateJson ||
          bookmark.bookmarkStateJson === bookmarkStateJson
      );

      const existingBookmarkByName = canMatchByName
        ? existingBookmarks.find(
            (bookmark) =>
              bookmark.name.toLowerCase() === normalizedInputName.toLowerCase()
          )
        : null;

      const existingBookmark =
        mode === "syncApplied"
          ? existingBookmarkByState || existingBookmarkByName
          : null;

      const preferredName = toReadableBookmarkName(
        normalizedInputName,
        existingBookmark?.name,
        existingBookmarkByState?.name,
        existingBookmarkByName?.name
      );

      const normalizedName = toUniqueBookmarkName(
        preferredName,
        existingBookmarks,
        existingBookmark?.id
      );

      const now = new Date().toISOString();

      const nextBookmark: BookmarkProfile = {
        id: existingBookmark?.id || createBookmarkId(),
        name: normalizedName,
        state: bookmarkStateJson,
        bookmarkStateJson,
        createdAt: existingBookmark?.createdAt || now,
        updatedAt: now,
        layoutState: layoutState ?? existingBookmark?.layoutState,
      };

      const nextBookmarks = [
        nextBookmark,
        ...existingBookmarks.filter((bookmark) => bookmark.id !== nextBookmark.id),
      ];

      persistBookmarks(nextBookmarks);

      if (selectAfterUpsert) {
        const nextSelectionId = toSavedBookmarkSelectionId(nextBookmark.id);
        setSelectedBookmarkId(nextSelectionId);
        window.localStorage.setItem(selectedBookmarkStorageKey, nextSelectionId);
      }

      return nextBookmark;
    },
    [
      persistBookmarks,
      selectedBookmarkStorageKey,
      setSelectedBookmarkId,
    ]
  );

  const applyBookmarkProfile = useCallback(
    async (bookmark: BookmarkProfile) => {
      if (!reportRef.current) {
        return;
      }

      suppressAutoSaveEventsRef.current = 3;

      try {
        let appliedBookmarkState = false;
        const bookmarkState = bookmark.state || bookmark.bookmarkStateJson;

        if (
          bookmarkState &&
          reportRef.current?.bookmarksManager?.applyState
        ) {
          await reportRef.current.bookmarksManager.applyState(
            bookmarkState
          );
          appliedBookmarkState = true;
        }

        if (!appliedBookmarkState && bookmark.filtersJson) {
          const filters = JSON.parse(bookmark.filtersJson);
          await applyPersonalizedFilters(reportRef.current, filters);
        }

        if (!appliedBookmarkState && bookmark.activePage && reportRef.current.getPages) {
          const pages = await reportRef.current.getPages();
          const page = pages.find((p: any) => p?.name === bookmark.activePage);
          if (page && typeof page.setActive === "function") {
            await page.setActive();
          }
        }

        const activePage = await reportRef.current?.getActivePage?.();
        if (activePage?.name) {
          setCurrentPage(activePage.name);
        } else if (bookmark.activePage) {
          setCurrentPage(bookmark.activePage);
        }

        const safeLastSaved = parseSafeDate(bookmark.updatedAt);
        if (safeLastSaved) {
          setLastSaved(safeLastSaved);
        }
      } catch (error) {
        console.error("Error applying bookmark profile:", error);
      }
    },
    [reportRef]
  );

  const runWhenReportReady = useCallback(
    async <T,>(action: () => Promise<T>, retries = 5): Promise<T> => {
      let lastError: unknown;

      for (let attempt = 0; attempt <= retries; attempt += 1) {
        try {
          return await action();
        } catch (error) {
          lastError = error;
          if (!isReportNotReadyError(error) || attempt === retries) {
            throw error;
          }

          await wait(350 * (attempt + 1));
        }
      }

      throw lastError;
    },
    []
  );

  const loadReportBookmarks = useCallback(
    async (retryForNonEmpty = false): Promise<models.IReportBookmark[]> => {
      if (
        !isReportLoaded ||
        !reportRef.current ||
        typeof reportRef.current?.bookmarksManager?.getBookmarks !== "function"
      ) {
        setReportBookmarks([]);
        return [];
      }

      try {
        const maxAttempts = retryForNonEmpty ? 4 : 1;
        let flattenedBookmarks: models.IReportBookmark[] = [];

        for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
          const bookmarks = await runWhenReportReady(async () => {
            return await reportRef.current.bookmarksManager.getBookmarks();
          });
          flattenedBookmarks = flattenReportBookmarks(
            Array.isArray(bookmarks) ? bookmarks : []
          );
          if (flattenedBookmarks.length > 0 || attempt === maxAttempts - 1) {
            setReportBookmarks(flattenedBookmarks);
            return flattenedBookmarks;
          }

          await wait(400 * (attempt + 1));
        }

        setReportBookmarks(flattenedBookmarks);
        return flattenedBookmarks;
      } catch (error) {
        console.warn("Unable to load report-defined bookmarks", error);
        setReportBookmarks([]);
        return [];
      }
    },
    [isReportLoaded, reportRef, runWhenReportReady]
  );

  useEffect(() => {
    try {
      const stored = window.localStorage.getItem(bookmarksStorageKey);
      const parsedBookmarks = stored
        ? (JSON.parse(stored) as BookmarkProfile[])
        : [];

      const normalizedBookmarks = normalizeAndDedupeBookmarkProfiles(
        parsedBookmarks
      );

      if (stored) {
        window.localStorage.setItem(
          bookmarksStorageKey,
          JSON.stringify(normalizedBookmarks)
        );
      }

      setBookmarkProfiles(normalizedBookmarks);
      bookmarkProfilesRef.current = normalizedBookmarks;

      const storedSelectedBookmarkRaw =
        window.localStorage.getItem(selectedBookmarkStorageKey) || "";

      const storedSelectedBookmark =
        !storedSelectedBookmarkRaw.startsWith(SAVED_BOOKMARK_PREFIX) &&
        !storedSelectedBookmarkRaw.startsWith(REPORT_BOOKMARK_PREFIX) &&
        normalizedBookmarks.some(
          (bookmark) => bookmark.id === storedSelectedBookmarkRaw
        )
          ? toSavedBookmarkSelectionId(storedSelectedBookmarkRaw)
          : storedSelectedBookmarkRaw;

      const savedBookmarkId = getSavedBookmarkIdFromSelection(
        storedSelectedBookmark
      );
      const hasSelectedSavedBookmark = !!savedBookmarkId
        ? normalizedBookmarks.some((bookmark) => bookmark.id === savedBookmarkId)
        : false;

      const isReportBookmarkSelection =
        !!getReportBookmarkNameFromSelection(storedSelectedBookmark);
      const isOriginalSelection = isOriginalReportSelection(
        storedSelectedBookmark
      );

      if (hasSelectedSavedBookmark || isReportBookmarkSelection || isOriginalSelection) {
        setSelectedBookmarkId(storedSelectedBookmark);
      } else {
        setSelectedBookmarkId("");
        window.localStorage.removeItem(selectedBookmarkStorageKey);
      }

      if (!stored) {
        showBookmarkStatus(null);
      }
    } catch (error) {
      console.warn("Failed loading bookmark profiles", error);
      setBookmarkProfiles([]);
      setSelectedBookmarkId("");
      window.localStorage.removeItem(selectedBookmarkStorageKey);
      showBookmarkStatus(null);
    }
  }, [bookmarksStorageKey, selectedBookmarkStorageKey, showBookmarkStatus]);

  useEffect(() => {
    let isCancelled = false;

    void (async () => {
      if (isCancelled) {
        return;
      }

      await loadReportBookmarks(true);
    })();

    return () => {
      isCancelled = true;
    };
  }, [loadReportBookmarks]);

  useEffect(() => {
    if (!isReportLoaded || !reportRef.current) {
      return;
    }

    if (hasHydratedRef.current) {
      return;
    }

    hasHydratedRef.current = true;
    setIsHydratingPersonalization(true);

    const loadPersonalization = async () => {
      try {
        const saved = await getPersonalization(userId, reportId);
        if (saved && reportRef.current) {
          suppressAutoSaveEventsRef.current = 2;

          let appliedSavedState = false;

          if (
            saved.settingsJson &&
            reportRef.current?.bookmarksManager?.applyState
          ) {
            try {
              const parsedSettings = JSON.parse(saved.settingsJson);
              const bookmarkStateFromSettings =
                typeof parsedSettings === "string"
                  ? parsedSettings
                  : parsedSettings?.bookmarkState;

              if (bookmarkStateFromSettings) {
                await runWhenReportReady(async () => {
                  await reportRef.current.bookmarksManager.applyState(
                    bookmarkStateFromSettings
                  );
                  return true;
                });
                appliedSavedState = true;
              }
            } catch (stateError) {
              console.warn("Failed applying saved bookmark state", stateError);
            }
          }

          if (!appliedSavedState && saved.filtersJson) {
            const filters = JSON.parse(saved.filtersJson);
            await runWhenReportReady(async () => {
              await applyPersonalizedFilters(reportRef.current, filters);
              return true;
            });
          }

          if (
            !appliedSavedState &&
            saved.activePage &&
            reportRef.current.getPages
          ) {
            const pages = await runWhenReportReady(async () => {
              return await reportRef.current.getPages();
            });
            const page = pages.find((p: any) => p.name === saved.activePage);
            if (page && typeof page.setActive === "function") {
              await runWhenReportReady(async () => {
                await page.setActive();
                return true;
              });
              setCurrentPage(saved.activePage);
            }
          }

          const safeLastSaved = parseSafeDate(saved.updatedAt);
          if (safeLastSaved) {
            setLastSaved(safeLastSaved);
          }
        }
      } catch (error) {
        console.error("Error loading personalization:", error);
      }

      let localBookmarks: BookmarkProfile[] = [];
      try {
        const rawBookmarks = window.localStorage.getItem(bookmarksStorageKey);
        const parsedBookmarks = rawBookmarks ? JSON.parse(rawBookmarks) : [];
        localBookmarks = normalizeAndDedupeBookmarkProfiles(parsedBookmarks);

        window.localStorage.setItem(
          bookmarksStorageKey,
          JSON.stringify(localBookmarks)
        );
      } catch {
        localBookmarks = [];
      }

      let reportDefinedBookmarks: models.IReportBookmark[] = [];
      reportDefinedBookmarks = await loadReportBookmarks(true);

      setBookmarkProfiles(localBookmarks);
      setReportBookmarks(reportDefinedBookmarks);

      const preferredBookmarkIdRaw = window.localStorage.getItem(
        selectedBookmarkStorageKey
      ) || "";
      const preferredBookmarkId =
        !preferredBookmarkIdRaw.startsWith(SAVED_BOOKMARK_PREFIX) &&
        !preferredBookmarkIdRaw.startsWith(REPORT_BOOKMARK_PREFIX) &&
        localBookmarks.some((bookmark) => bookmark.id === preferredBookmarkIdRaw)
          ? toSavedBookmarkSelectionId(preferredBookmarkIdRaw)
          : preferredBookmarkIdRaw;
      const preferredSavedBookmarkId = getSavedBookmarkIdFromSelection(
        preferredBookmarkId
      );
      const preferredReportBookmarkName = getReportBookmarkNameFromSelection(
        preferredBookmarkId
      );
      const prefersOriginalSelection = isOriginalReportSelection(
        preferredBookmarkId
      );

      const preferredSavedBookmark = preferredSavedBookmarkId
        ? localBookmarks.find((bookmark) => bookmark.id === preferredSavedBookmarkId)
        : null;
      const preferredReportBookmark = preferredReportBookmarkName
        ? reportDefinedBookmarks.find(
            (bookmark) => bookmark.name === preferredReportBookmarkName
          )
        : null;

      const fallbackSelectionId = localBookmarks[0]?.id
        ? toSavedBookmarkSelectionId(localBookmarks[0].id)
        : reportDefinedBookmarks[0]?.name
          ? toReportBookmarkSelectionId(reportDefinedBookmarks[0].name)
          : "";

      const nextSelectionId = preferredSavedBookmark
        ? toSavedBookmarkSelectionId(preferredSavedBookmark.id)
        : preferredReportBookmark
          ? toReportBookmarkSelectionId(preferredReportBookmark.name)
          : prefersOriginalSelection
            ? ORIGINAL_REPORT_SELECTION_ID
            : fallbackSelectionId;

      if (nextSelectionId) {
        setSelectedBookmarkId(nextSelectionId);
        window.localStorage.setItem(selectedBookmarkStorageKey, nextSelectionId);

        const selectedSavedBookmarkId = getSavedBookmarkIdFromSelection(
          nextSelectionId
        );
        const selectedReportBookmarkName = getReportBookmarkNameFromSelection(
          nextSelectionId
        );

        if (selectedSavedBookmarkId) {
          const selectedSavedBookmark = localBookmarks.find(
            (bookmark) => bookmark.id === selectedSavedBookmarkId
          );
          if (selectedSavedBookmark) {
            await applyBookmarkProfile(selectedSavedBookmark);
          }
        } else if (
          selectedReportBookmarkName &&
          typeof reportRef.current?.bookmarksManager?.apply === "function"
        ) {
          await reportRef.current.bookmarksManager.apply(selectedReportBookmarkName);
          const activePage = await reportRef.current?.getActivePage?.();
          if (activePage?.name) {
            setCurrentPage(activePage.name);
          }
        } else if (isOriginalReportSelection(nextSelectionId)) {
          await restoreOriginalReportState();
        }
      } else {
        setSelectedBookmarkId("");
        window.localStorage.removeItem(selectedBookmarkStorageKey);
        showBookmarkStatus(null);
      }
      setIsHydratingPersonalization(false);
    };

    loadPersonalization().catch((error) => {
      console.error("Error hydrating personalization state:", error);
      setIsHydratingPersonalization(false);
    });
  }, [
    isReportLoaded,
    reportRef,
    userId,
    reportId,
    getPersonalization,
    bookmarksStorageKey,
    applyBookmarkProfile,
    runWhenReportReady,
    selectedBookmarkStorageKey,
    showBookmarkStatus,
    loadReportBookmarks,
    restoreOriginalReportState,
  ]);

  const getCurrentPersonalizationPayload = useCallback(async () => {
    if (!reportRef.current) {
      return null;
    }

    const state = await getReportPersonalizationState(reportRef.current);
    let bookmarkState = "";

    if (reportRef.current?.bookmarksManager?.capture) {
      try {
        const captured = await reportRef.current.bookmarksManager.capture({
          personalizeVisuals: true,
        });
        bookmarkState = captured?.state || "";
      } catch (error) {
        console.warn("Unable to capture bookmark state for autosave", error);
      }
    }

    return {
      userId,
      reportId,
      workspaceId,
      filtersJson: JSON.stringify(state.filters),
      bookmarksJson: JSON.stringify(state.bookmarks),
      activePage: state.currentPage,
      settingsJson: JSON.stringify({
        schemaVersion: 1,
        bookmarkState,
      }),
    };
  }, [reportRef, userId, reportId, workspaceId]);

  useEffect(() => {
    getCurrentPersonalizationPayloadRef.current = getCurrentPersonalizationPayload;
  }, [getCurrentPersonalizationPayload]);

  const persistCurrentPersonalizationSnapshot = useCallback(async () => {
    if (!reportRef.current || isHydratingPersonalizationRef.current) {
      return false;
    }

    try {
      const payload = await getCurrentPersonalizationPayloadRef.current();
      if (!payload) {
        return false;
      }

      await savePersonalizationRef.current(payload);
      return true;
    } catch (error) {
      console.warn("Failed to persist personalization snapshot", error);
      return false;
    }
  }, [reportRef]);

  const readSessionQuickVisualProfiles = useCallback(() => {
    try {
      const stored = window.sessionStorage.getItem(quickVisualSessionStorageKey);
      const parsed = stored ? JSON.parse(stored) : [];
      if (!Array.isArray(parsed)) {
        return [] as SessionQuickVisualProfile[];
      }

      return parsed
        .map((entry: any) => {
          if (!entry || typeof entry !== "object") {
            return null;
          }

          if (typeof entry.visualType !== "string" || !entry.visualType) {
            return null;
          }

          const layout = entry.layout || {};
          const width = Number(layout.width);
          const height = Number(layout.height);
          const normalizedLayout: SessionQuickVisualLayout = {
            x: Number(layout.x) || 24,
            y: Number(layout.y) || 24,
            width: width > 0 ? width : 460,
            height: height > 0 ? height : 300,
          };

          const properties = entry.properties || {};
          const normalizedProperties: SessionQuickVisualProperties = {
            showTitle: Boolean(properties.showTitle),
            showLegend: Boolean(properties.showLegend),
            showXAxis: Boolean(properties.showXAxis),
            showYAxis: Boolean(properties.showYAxis),
            titleText:
              typeof properties.titleText === "string"
                ? properties.titleText
                : "",
            titleAlign:
              typeof properties.titleAlign === "string" && properties.titleAlign
                ? properties.titleAlign
                : "left",
          };

          const roleTargets =
            entry.roleTargets && typeof entry.roleTargets === "object"
              ? entry.roleTargets
              : {};

          return {
            id:
              typeof entry.id === "string" && entry.id
                ? entry.id
                : createBookmarkId(),
            visualType: entry.visualType,
            roleTargets,
            layout: normalizedLayout,
            properties: normalizedProperties,
            runtimeVisualName:
              typeof entry.runtimeVisualName === "string" &&
              entry.runtimeVisualName
                ? entry.runtimeVisualName
                : undefined,
            createdAt:
              typeof entry.createdAt === "string"
                ? entry.createdAt
                : new Date().toISOString(),
          } as SessionQuickVisualProfile;
        })
        .filter(Boolean) as SessionQuickVisualProfile[];
    } catch {
      return [] as SessionQuickVisualProfile[];
    }
  }, [quickVisualSessionStorageKey]);

  const writeSessionQuickVisualProfiles = useCallback(
    (profiles: SessionQuickVisualProfile[]) => {
      window.sessionStorage.setItem(
        quickVisualSessionStorageKey,
        JSON.stringify(profiles)
      );
    },
    [quickVisualSessionStorageKey]
  );

  const appendSessionQuickVisualProfile = useCallback(
    (profile: SessionQuickVisualProfile) => {
      const existingProfiles = readSessionQuickVisualProfiles();
      const nextProfiles = [...existingProfiles, profile];
      writeSessionQuickVisualProfiles(nextProfiles);
    },
    [readSessionQuickVisualProfiles, writeSessionQuickVisualProfiles]
  );

  const syncSessionQuickVisualProfilesFromPage = useCallback(async () => {
    if (!allowEdit || !authoringReportRef.current) {
      return false;
    }

    const profiles = readSessionQuickVisualProfiles();
    if (!profiles.length) {
      return false;
    }

    // FIX: Use the dedicated authoring page ref instead of getActivePage()
    // to avoid reading layout from the wrong report page.
    const activePage = authoringPageRef.current;
    if (!activePage || typeof activePage.getVisuals !== "function") {
      return false;
    }

    let visuals: any[] = [];
    try {
      visuals = await activePage.getVisuals();
    } catch {
      return false;
    }

    const visualByName = new Map<string, any>(
      (visuals || [])
        .filter((visual: any) => visual?.name)
        .map((visual: any) => [visual.name, visual] as [string, any])
    );

    let hasChanges = false;
    const nextProfiles = profiles.map((profile) => {
      if (!profile.runtimeVisualName) {
        return profile;
      }

      const matchedVisual = visualByName.get(profile.runtimeVisualName);
      if (!matchedVisual) {
        return profile;
      }

      const layout = matchedVisual.layout || {};
      const nextLayout: SessionQuickVisualLayout = {
        x: Number(layout.x) || profile.layout.x,
        y: Number(layout.y) || profile.layout.y,
        width:
          Number(layout.width) > 0 ? Number(layout.width) : profile.layout.width,
        height:
          Number(layout.height) > 0
            ? Number(layout.height)
            : profile.layout.height,
      };

      const layoutChanged =
        nextLayout.x !== profile.layout.x ||
        nextLayout.y !== profile.layout.y ||
        nextLayout.width !== profile.layout.width ||
        nextLayout.height !== profile.layout.height;

      if (!layoutChanged) {
        return profile;
      }

      hasChanges = true;
      return {
        ...profile,
        layout: nextLayout,
      };
    });

    if (hasChanges) {
      writeSessionQuickVisualProfiles(nextProfiles);
    }

    return hasChanges;
  }, [
    allowEdit,
    readSessionQuickVisualProfiles,
    writeSessionQuickVisualProfiles,
  ]);

  const scheduleSessionQuickVisualProfileSync = useCallback(() => {
    if (quickVisualSessionSyncTimerRef.current) {
      window.clearTimeout(quickVisualSessionSyncTimerRef.current);
    }

    quickVisualSessionSyncTimerRef.current = window.setTimeout(() => {
      void syncSessionQuickVisualProfilesFromPage();
      quickVisualSessionSyncTimerRef.current = null;
    }, 300);
  }, [syncSessionQuickVisualProfilesFromPage]);

  useEffect(() => {
    if (
      !autoSaveEnabled ||
      !reportRef.current ||
      autoSaveRevision === 0 ||
      isHydratingPersonalization
    ) {
      return;
    }

    const saveTimer = window.setTimeout(async () => {
      if (autoSaveInFlightRef.current) {
        return;
      }

      autoSaveInFlightRef.current = true;

      try {
        setSaveStatus("saving");
        const payload = await getCurrentPersonalizationPayload();
        if (!payload) {
          setSaveStatus("idle");
          return;
        }
        await savePersonalization(payload);

        setSaveStatus("saved");
        setLastSaved(new Date());
        window.setTimeout(() => setSaveStatus("idle"), 1500);
      } catch (error) {
        console.error("Error auto-saving personalization:", error);
        setSaveStatus("error");
      } finally {
        autoSaveInFlightRef.current = false;
      }
    }, 1000);

    return () => window.clearTimeout(saveTimer);
  }, [
    autoSaveEnabled,
    autoSaveRevision,
    isHydratingPersonalization,
    reportRef,
    userId,
    reportId,
    workspaceId,
    savePersonalization,
    getCurrentPersonalizationPayload,
  ]);

  useEffect(() => {
    return () => {
      if (
        !isReportLoadedRef.current ||
        !autoSaveEnabledRef.current ||
        isHydratingPersonalizationRef.current ||
        !reportRef.current ||
        autoSaveInFlightRef.current
      ) {
        return;
      }

      autoSaveInFlightRef.current = true;
      void (async () => {
        try {
          const payload = await getCurrentPersonalizationPayloadRef.current();
          if (payload) {
            await savePersonalizationRef.current(payload);
          }
        } catch (error) {
          console.warn("Failed to flush personalization on cleanup", error);
        } finally {
          autoSaveInFlightRef.current = false;
        }
      })();
    };
  }, [
    isReportLoaded,
    autoSaveEnabled,
    reportRef,
    getCurrentPersonalizationPayload,
    savePersonalization,
  ]);

  const triggerAutoSaveRevision = useCallback(() => {
    if (!isReportLoaded || isHydratingPersonalization) {
      return;
    }

    if (suppressAutoSaveEventsRef.current > 0) {
      suppressAutoSaveEventsRef.current -= 1;
      return;
    }

    setAutoSaveRevision((prev) => prev + 1);
  }, [isReportLoaded, isHydratingPersonalization]);

  const ensureAuthoringReportReady = useCallback(
    async (retries = 5) => {
      if (authoringEmbedError) {
        return false;
      }

      for (let attempt = 0; attempt <= retries; attempt += 1) {
        const authoringReport = authoringReportRef.current;
        if (!authoringReport) {
          await wait(250 * (attempt + 1));
          continue;
        }

        try {
          if (typeof authoringReport.getPages === "function") {
            const pages = await authoringReport.getPages();
            if (Array.isArray(pages) && pages.length > 0) {
              return true;
            }
          }

          if (typeof authoringReport.getActivePage === "function") {
            const activePage = await authoringReport.getActivePage();
            if (activePage) {
              return true;
            }
          }
        } catch (error) {
          if (!isReportNotReadyError(error) && attempt === retries) {
            console.warn("Authoring report readiness check failed", error);
          }
        }

        await wait(250 * (attempt + 1));
      }

      return false;
    },
    [authoringEmbedError]
  );

  // FIX: Always use the dedicated blank authoring page (pages[1] per the showcase pattern).
  // Falls back to getActivePage only if pages[1] hasn't been initialised yet.
  const getActivePageForQuickVisual = useCallback(async () => {
    // Primary: use the cached blank authoring page reference
    if (authoringPageRef.current) {
      return authoringPageRef.current;
    }

    const isAuthoringReady = await ensureAuthoringReportReady();
    if (!isAuthoringReady || !authoringReportRef.current) {
      return null;
    }

    // Try to obtain pages[1] (the blank authoring page) from the authoring report
    if (typeof authoringReportRef.current.getPages === "function") {
      try {
        const pages = await authoringReportRef.current.getPages();
        if (Array.isArray(pages) && pages.length > 1) {
          // pages[1] is the empty/blank page used for visual authoring per the showcase
          const blankPage = pages[1];
          await blankPage.setActive();
          authoringPageRef.current = blankPage;
          return blankPage;
        }
        // Single-page report: use pages[0] as fallback
        if (Array.isArray(pages) && pages.length > 0) {
          authoringPageRef.current = pages[0];
          return pages[0];
        }
      } catch (error) {
        console.warn("Unable to get authoring pages", error);
      }
    }

    // Last resort: getActivePage
    if (typeof authoringReportRef.current.getActivePage === "function") {
      try {
        const activePage = await authoringReportRef.current.getActivePage();
        if (activePage) {
          authoringPageRef.current = activePage;
          return activePage;
        }
      } catch { }
    }

    return null;
  }, [ensureAuthoringReportReady]);

  const canAccessQuickVisualAPIs = useCallback(async () => {
    if (authoringEmbedError) {
      return false;
    }

    const authoringReport = authoringReportRef.current;
    if (!authoringReport) {
      return false;
    }

    try {
      const page = await getActivePageForQuickVisual();
      if (!page) {
        return false;
      }

      return typeof page.createVisual === "function";
    } catch {
      return false;
    }
  }, [authoringEmbedError, getActivePageForQuickVisual]);

  const getQuickVisualAuthoringAvailabilityMessage = useCallback(
    async (mode: QuickVisualMode) => {
      if (!allowEdit) {
        return "Quick visual editing requires workspace edit permission (Member/Contributor/Admin).";
      }

      if (authoringEmbedError) {
        return authoringEmbedError;
      }

      const canAccessAPIs = await canAccessQuickVisualAPIs();
      if (!canAccessAPIs) {
        return "Quick visual authoring instance is still loading. Try again in a moment.";
      }

      if (typeof authoringReportRef.current.getMode === "function") {
        try {
          const currentMode = await authoringReportRef.current.getMode();
          const isEditMode =
            currentMode === models.ViewMode.Edit ||
            String(currentMode).toLowerCase() === "edit";

          if (!isEditMode) {
            return "Quick visual editing is only available in Edit mode.";
          }
        } catch { }
      }

      const activePage = await getActivePageForQuickVisual();
      if (!activePage) {
        return "Active page is not available for quick visual edits.";
      }

      if (mode === "create" && typeof activePage.createVisual !== "function") {
        return "Create visual is unavailable in this embed context. Ensure report-authoring APIs are enabled and your token/user has edit rights.";
      }

      return null;
    },
    [
      allowEdit,
      authoringEmbedError,
      canAccessQuickVisualAPIs,
      getActivePageForQuickVisual,
    ]
  );

  const getQuickVisualPages = useCallback(async () => {
    if (
      authoringReportRef.current &&
      typeof authoringReportRef.current.getPages === "function"
    ) {
      try {
        const pages = await authoringReportRef.current.getPages();
        if (Array.isArray(pages) && pages.length) {
          return pages;
        }
      } catch (error) {
        console.warn("Unable to list report pages for quick visual fields", error);
      }
    }

    const activePage = await getActivePageForQuickVisual();
    return activePage ? [activePage] : [];
  }, [getActivePageForQuickVisual]);

  const refreshQuickVisualTargets = useCallback(async () => {
    // FIX: For "change" mode, query visuals on the BASE report's active page,
    // not the blank authoring page (which only contains the preview visual).
    let targetPage: any = null;
    if (reportRef.current && typeof reportRef.current.getActivePage === "function") {
      try {
        targetPage = await reportRef.current.getActivePage();
      } catch { }
    }

    if (!targetPage || typeof targetPage.getVisuals !== "function") {
      setQuickVisualTargets([]);
      return [] as QuickVisualTarget[];
    }

    const visuals = await targetPage.getVisuals();
    const targets = (visuals || [])
      .filter((visual: any) => visual?.name)
      .map((visual: any) => ({
        name: visual.name,
        title: visual.title || visual.name,
        type: visual.type || "",
        layout: visual.layout,
      })) as QuickVisualTarget[];

    setQuickVisualTargets(targets);
    return targets;
  }, [reportRef]);

  const refreshQuickVisualFieldOptions = useCallback(async () => {
    const optionByKey = new Map<string, QuickVisualFieldOption>();

    const collectFromUnknownValue = (value: any) => {
      collectQuickVisualFieldTargets(value, (target) => {
        addQuickVisualFieldOption(optionByKey, target);
      });
    };

    const collectFromFieldCatalog = (catalog: any[]) => {
      if (!Array.isArray(catalog)) {
        return;
      }

      catalog.forEach((entry) => {
        addQuickVisualFieldOption(optionByKey, entry);
        if (entry?.target && typeof entry.target === "object") {
          addQuickVisualFieldOption(optionByKey, entry.target);
        }
      });
    };

    collectFromFieldCatalog(customQuickVisualFieldCatalog);
    try {
      const savedCatalog = window.localStorage.getItem(
        quickVisualFieldCatalogStorageKey
      );
      if (savedCatalog) {
        const parsedCatalog = JSON.parse(savedCatalog);
        collectFromFieldCatalog(parsedCatalog);
      }
    } catch { }

    const pages = await getQuickVisualPages();

    for (const page of pages) {
      if (!page || typeof page.getVisuals !== "function") {
        continue;
      }

      if (typeof page.getFilters === "function") {
        try {
          const pageFilters = await page.getFilters();
          collectFromUnknownValue(pageFilters);
        } catch { }
      }

      let visuals: any[] = [];
      try {
        visuals = await page.getVisuals();
      } catch {
        continue;
      }

      for (const visual of visuals || []) {
        if (!visual) {
          continue;
        }

        if (typeof visual.getFilters === "function") {
          try {
            const visualFilters = await visual.getFilters();
            collectFromUnknownValue(visualFilters);
          } catch { }
        }

        if (typeof visual.getDataFields !== "function") {
          continue;
        }

        const roleNames = new Set<string>(QUICK_VISUAL_DEFAULT_ROLE_NAMES);
        if (typeof visual.getCapabilities === "function") {
          try {
            const capabilities = await visual.getCapabilities();
            const dataRoles = Array.isArray(capabilities?.dataRoles)
              ? capabilities.dataRoles
              : [];

            dataRoles.forEach((role: any) => {
              if (typeof role?.name === "string" && role.name) {
                roleNames.add(role.name);
              }
              if (typeof role?.displayName === "string" && role.displayName) {
                roleNames.add(role.displayName);
              }
            });
          } catch { }
        }

        for (const roleName of roleNames) {
          try {
            const fields = await visual.getDataFields(roleName);
            if (!Array.isArray(fields)) {
              continue;
            }

            for (const field of fields) {
              addQuickVisualFieldOption(optionByKey, field);
            }
          } catch { }
        }
      }
    }

    if (
      authoringReportRef.current &&
      typeof authoringReportRef.current.getFilters === "function"
    ) {
      try {
        const reportFilters = await authoringReportRef.current.getFilters();
        collectFromUnknownValue(reportFilters);
      } catch { }
    }

    try {
      const selectedData = window.localStorage.getItem("selectedData");
      if (selectedData) {
        collectFromUnknownValue(JSON.parse(selectedData));
      }
    } catch { }

    const options = Array.from(optionByKey.values()).sort((a, b) =>
      a.label.localeCompare(b.label)
    );
    setQuickVisualFieldOptions(options);
    return options;
  }, [
    customQuickVisualFieldCatalog,
    getQuickVisualPages,
    quickVisualFieldCatalogStorageKey,
  ]);

  const getRoleSelectionsFromVisual = useCallback(
    async (
      visual: any,
      visualType: string,
      fieldOptions: QuickVisualFieldOption[]
    ) => {
      const selections: QuickVisualRoleSelections = {};
      if (!visual || typeof visual.getDataFields !== "function") {
        return selections;
      }

      const roleNames = getQuickVisualOption(visualType).dataRoleNames;
      const optionKeys = new Set(fieldOptions.map((option) => option.key));

      for (const roleName of roleNames) {
        try {
          const dataFields = await visual.getDataFields(roleName);
          if (!Array.isArray(dataFields) || !dataFields[0]) {
            continue;
          }

          const key = getQuickVisualFieldKey(dataFields[0]);
          if (key && optionKeys.has(key)) {
            selections[roleName] = key;
          }
        } catch { }
      }

      return selections;
    },
    []
  );

  const syncQuickVisualSelectionsFromVisual = useCallback(
    async (visualName: string, visualType: string) => {
      if (!visualName) {
        setQuickVisualRoleSelections({});
        return;
      }

      try {
        // FIX: For "change" mode, read the visual from the BASE report (where the real
        // report visuals live), not from the blank authoring page.
        let targetPage: any = null;
        if (reportRef.current && typeof reportRef.current.getActivePage === "function") {
          try {
            targetPage = await reportRef.current.getActivePage();
          } catch { }
        }

        if (!targetPage || typeof targetPage.getVisuals !== "function") {
          setQuickVisualRoleSelections({});
          return;
        }

        const visuals = await targetPage.getVisuals();
        const matchedVisual = (visuals || []).find(
          (visual: any) => visual?.name === visualName
        );
        const selections = await getRoleSelectionsFromVisual(
          matchedVisual,
          visualType,
          quickVisualFieldOptions
        );
        setQuickVisualRoleSelections(selections);
      } catch (error) {
        console.warn("Unable to read selected visual fields", error);
        setQuickVisualRoleSelections({});
      }
    },
    [reportRef, getRoleSelectionsFromVisual, quickVisualFieldOptions]
  );

  const openQuickVisualModal = useCallback(
    async (mode: QuickVisualMode, preferredVisualName?: string) => {
      const authoringAvailabilityMessage =
        await getQuickVisualAuthoringAvailabilityMessage(mode);
      if (authoringAvailabilityMessage) {
        showQuickVisualStatus(authoringAvailabilityMessage);
        return;
      }

      try {
        await syncVisibleStateToAuthoring();

        const targets = await refreshQuickVisualTargets();
        const fieldOptions = await refreshQuickVisualFieldOptions();

        if (mode === "change" && !targets.length) {
          showQuickVisualStatus(
            "No editable visuals were found on the active page."
          );
          return;
        }

        if (!fieldOptions.length) {
          showQuickVisualStatus(
            "No usable fields were found. Open a report page with data-bound visuals, or provide quickVisualFieldCatalog entries for missing fields."
          );
          return;
        }

        const defaultVisual =
          (preferredVisualName
            ? targets.find((visual) => visual.name === preferredVisualName)
            : null) || targets[0];

        const defaultType =
          defaultVisual?.type &&
          QUICK_VISUAL_OPTIONS.some((option) => option.name === defaultVisual.type)
            ? defaultVisual.type
            : "columnChart";

        const option = getQuickVisualOption(defaultType);

        setQuickVisualMode(mode);
        setQuickVisualType(defaultType);
        setQuickVisualTitle("");
        setQuickVisualTitleAlign("left");
        setQuickVisualShowTitle(true);
        setQuickVisualShowLegend(option.properties.includes("legend"));
        setQuickVisualShowXAxis(option.properties.includes("xAxis"));
        setQuickVisualShowYAxis(option.properties.includes("yAxis"));
        setQuickVisualStatus(null);

        const nextTargetVisualName =
          mode === "change" ? defaultVisual?.name || "" : "";
        setQuickVisualTargetVisualName(nextTargetVisualName);

        let defaultSelections: QuickVisualRoleSelections = {};
        if (mode === "change" && defaultVisual?.name) {
          // FIX: Read role selections from the base report visual, not authoring page
          let targetPage: any = null;
          if (reportRef.current && typeof reportRef.current.getActivePage === "function") {
            try {
              targetPage = await reportRef.current.getActivePage();
            } catch { }
          }
          if (targetPage && typeof targetPage.getVisuals === "function") {
            const pageVisuals = await targetPage.getVisuals();
            const matchedVisual = (pageVisuals || []).find(
              (visual: any) => visual?.name === defaultVisual.name
            );
            defaultSelections = await getRoleSelectionsFromVisual(
              matchedVisual,
              defaultType,
              fieldOptions
            );
          }
        }

        setQuickVisualRoleSelections(defaultSelections);

        setIsQuickVisualModalOpen(true);
      } catch (error) {
        console.error("Unable to open quick visual modal", error);
        showQuickVisualStatus(
          "Unable to open quick visual editor for this report."
        );
      }
    },
    [
      getQuickVisualAuthoringAvailabilityMessage,
      reportRef,
      getRoleSelectionsFromVisual,
      refreshQuickVisualFieldOptions,
      refreshQuickVisualTargets,
      showQuickVisualStatus,
      syncVisibleStateToAuthoring,
    ]
  );

  const closeQuickVisualModal = () => {
    if (isQuickVisualApplying) {
      return;
    }

    setIsQuickVisualModalOpen(false);
  };

  const applyQuickVisualProperty = useCallback(
    async (visual: any, propertyName: QuickVisualProperty, value: any) => {
      if (!visual || typeof visual.setProperty !== "function") {
        return;
      }

      const selector = toQuickVisualPropertySelector(propertyName);
      if (!selector) {
        return;
      }

      try {
        await visual.setProperty(selector, {
          schema: QUICK_VISUAL_SCHEMAS.property,
          value,
        });
      } catch (error) {
        console.warn(`Unable to set visual property ${propertyName}`, error);
      }
    },
    []
  );

  const applyQuickVisualProperties = useCallback(
    async (visual: any, visualType: string) => {
      const option = getQuickVisualOption(visualType);
      const supported = new Set(option.properties);

      if (supported.has("legend")) {
        await applyQuickVisualProperty(visual, "legend", quickVisualShowLegend);
      }

      if (supported.has("xAxis")) {
        await applyQuickVisualProperty(visual, "xAxis", quickVisualShowXAxis);
      }

      if (supported.has("yAxis")) {
        await applyQuickVisualProperty(visual, "yAxis", quickVisualShowYAxis);
      }

      await applyQuickVisualProperty(visual, "title", quickVisualShowTitle);
      await applyQuickVisualProperty(visual, "titleSize", 13);
      await applyQuickVisualProperty(visual, "titleColor", "#000000");

      if (!quickVisualShowTitle) {
        if (typeof visual?.resetProperty === "function") {
          const titleTextSelector = toQuickVisualPropertySelector("titleText");
          if (titleTextSelector) {
            try {
              await visual.resetProperty(titleTextSelector);
            } catch (error) {
              console.warn("Unable to reset visual title", error);
            }
          }
        }
        return;
      }

      if (quickVisualTitle.trim()) {
        await applyQuickVisualProperty(
          visual,
          "titleText",
          quickVisualTitle.trim()
        );
      } else if (typeof visual?.resetProperty === "function") {
        const titleTextSelector = toQuickVisualPropertySelector("titleText");
        if (titleTextSelector) {
          try {
            await visual.resetProperty(titleTextSelector);
          } catch (error) {
            console.warn("Unable to reset visual title", error);
          }
        }
      }

      await applyQuickVisualProperty(visual, "titleAlign", quickVisualTitleAlign);
    },
    [
      applyQuickVisualProperty,
      quickVisualShowLegend,
      quickVisualShowTitle,
      quickVisualShowXAxis,
      quickVisualShowYAxis,
      quickVisualTitle,
      quickVisualTitleAlign,
    ]
  );

  // FIX: Remove all active data fields from a visual before changing type,
  // matching the showcase pattern (removeAllActiveDataRoles → changeType).
  const removeAllActiveDataRoles = useCallback(
    async (visual: any, visualType: string) => {
      if (!visual || typeof visual.getDataFields !== "function") {
        return;
      }

      const roleNames = getQuickVisualOption(visualType).dataRoleNames;

      for (const roleName of roleNames) {
        try {
          const existingFields = await visual.getDataFields(roleName);
          if (!Array.isArray(existingFields)) {
            continue;
          }

          // Remove all fields by repeatedly removing index 0
          for (let idx = 0; idx < existingFields.length; idx += 1) {
            if (typeof visual.removeDataField === "function") {
              await visual.removeDataField(roleName, 0);
            }
          }
        } catch (error) {
          console.warn(`Unable to remove data role ${roleName} before changeType`, error);
        }
      }
    },
    []
  );

  const applyQuickVisualDataRoles = useCallback(
    async (
      visual: any,
      visualType: string,
      roleSelections: QuickVisualRoleSelections
    ) => {
      if (!visual) {
        return;
      }

      const roleNames = getQuickVisualOption(visualType).dataRoleNames;

      for (const roleName of roleNames) {
        if (
          typeof visual.getDataFields === "function" &&
          typeof visual.removeDataField === "function"
        ) {
          try {
            const existingFields = await visual.getDataFields(roleName);
            if (Array.isArray(existingFields)) {
              for (let idx = 0; idx < existingFields.length; idx += 1) {
                await visual.removeDataField(roleName, 0);
              }
            }
          } catch (error) {
            console.warn(`Unable to clear data role ${roleName}`, error);
          }
        }

        const selectedFieldKey = roleSelections[roleName];
        if (!selectedFieldKey || typeof visual.addDataField !== "function") {
          continue;
        }

        const selectedField = quickVisualFieldOptions.find(
          (option) => option.key === selectedFieldKey
        );
        if (!selectedField) {
          continue;
        }

        try {
          await visual.addDataField(roleName, selectedField.target);
        } catch (error) {
          console.warn(`Unable to apply data role ${roleName}`, error);
        }
      }
    },
    [quickVisualFieldOptions]
  );

  const applyQuickVisualDataRolesFromTargets = useCallback(
    async (
      visual: any,
      visualType: string,
      roleTargets: Record<string, any>
    ) => {
      if (!visual) {
        return;
      }

      const roleNames = getQuickVisualOption(visualType).dataRoleNames;

      for (const roleName of roleNames) {
        if (
          typeof visual.getDataFields === "function" &&
          typeof visual.removeDataField === "function"
        ) {
          try {
            const existingFields = await visual.getDataFields(roleName);
            if (Array.isArray(existingFields)) {
              for (let idx = 0; idx < existingFields.length; idx += 1) {
                await visual.removeDataField(roleName, 0);
              }
            }
          } catch { }
        }

        const target = roleTargets?.[roleName];
        if (!target || typeof visual.addDataField !== "function") {
          continue;
        }

        try {
          await visual.addDataField(roleName, target);
        } catch (error) {
          console.warn(`Unable to replay data role ${roleName}`, error);
        }
      }
    },
    []
  );

  const applyQuickVisualPropertiesFromProfile = useCallback(
    async (
      visual: any,
      visualType: string,
      properties: SessionQuickVisualProperties
    ) => {
      const option = getQuickVisualOption(visualType);
      const supported = new Set(option.properties);

      if (supported.has("legend")) {
        await applyQuickVisualProperty(visual, "legend", properties.showLegend);
      }

      if (supported.has("xAxis")) {
        await applyQuickVisualProperty(visual, "xAxis", properties.showXAxis);
      }

      if (supported.has("yAxis")) {
        await applyQuickVisualProperty(visual, "yAxis", properties.showYAxis);
      }

      await applyQuickVisualProperty(visual, "title", properties.showTitle);
      await applyQuickVisualProperty(visual, "titleSize", 13);
      await applyQuickVisualProperty(visual, "titleColor", "#000000");

      if (!properties.showTitle) {
        if (typeof visual?.resetProperty === "function") {
          const selector = toQuickVisualPropertySelector("titleText");
          if (selector) {
            try {
              await visual.resetProperty(selector);
            } catch { }
          }
        }
        return;
      }

      const titleText = properties.titleText?.trim?.() || "";
      if (titleText) {
        await applyQuickVisualProperty(visual, "titleText", titleText);
      }

      await applyQuickVisualProperty(
        visual,
        "titleAlign",
        properties.titleAlign || "left"
      );
    },
    [applyQuickVisualProperty]
  );

  const replaySessionQuickVisuals = useCallback(async () => {
    if (!allowEdit || !authoringReportRef.current) {
      return 0;
    }

    const profiles = readSessionQuickVisualProfiles();
    if (!profiles.length) {
      return 0;
    }

    // FIX: Use the dedicated blank authoring page
    const activePage = await getActivePageForQuickVisual();
    if (!activePage || typeof activePage.createVisual !== "function") {
      return 0;
    }

    let restoredCount = 0;
    const nextProfiles: SessionQuickVisualProfile[] = [];

    for (const profile of profiles) {
      try {
        // FIX: createVisual returns ICreateVisualResponse; extract .visual from the response
        const createVisualResponse = await activePage.createVisual(profile.visualType, {
          x: profile.layout.x,
          y: profile.layout.y,
          width: profile.layout.width,
          height: profile.layout.height,
          displayState: {
            mode: models.VisualContainerDisplayMode.Visible,
          },
        });

        // Per the official API: createVisual returns { visual: IVisual }
        const visual: any = createVisualResponse?.visual ?? createVisualResponse;
        if (!visual) {
          nextProfiles.push(profile);
          continue;
        }

        await applyQuickVisualDataRolesFromTargets(
          visual,
          profile.visualType,
          profile.roleTargets
        );
        await applyQuickVisualPropertiesFromProfile(
          visual,
          profile.visualType,
          profile.properties
        );

        const visualLayout = visual.layout || {};
        nextProfiles.push({
          ...profile,
          runtimeVisualName:
            typeof visual.name === "string" && visual.name
              ? visual.name
              : profile.runtimeVisualName,
          layout: {
            x: Number(visualLayout.x) || profile.layout.x,
            y: Number(visualLayout.y) || profile.layout.y,
            width:
              Number(visualLayout.width) > 0
                ? Number(visualLayout.width)
                : profile.layout.width,
            height:
              Number(visualLayout.height) > 0
                ? Number(visualLayout.height)
                : profile.layout.height,
          },
        });

        restoredCount += 1;
      } catch (error) {
        console.warn("Unable to restore a session quick visual", error);
        nextProfiles.push(profile);
      }
    }

    writeSessionQuickVisualProfiles(nextProfiles);

    if (restoredCount > 0) {
      await syncAuthoringStateToVisible();
      await refreshQuickVisualTargets();
      await persistCurrentPersonalizationSnapshot();
    }

    return restoredCount;
  }, [
    allowEdit,
    applyQuickVisualDataRolesFromTargets,
    applyQuickVisualPropertiesFromProfile,
    getActivePageForQuickVisual,
    persistCurrentPersonalizationSnapshot,
    readSessionQuickVisualProfiles,
    refreshQuickVisualTargets,
    syncAuthoringStateToVisible,
    writeSessionQuickVisualProfiles,
  ]);

  useEffect(() => {
    if (
      !isReportLoaded ||
      !isAuthoringReportLoaded ||
      isHydratingPersonalization ||
      hasReplayedSessionQuickVisualsRef.current
    ) {
      return;
    }

    hasReplayedSessionQuickVisualsRef.current = true;

    void (async () => {
      const restoredCount = await replaySessionQuickVisuals();
      if (restoredCount > 0) {
        showQuickVisualStatus(
          `Restored ${restoredCount} session visual${
            restoredCount === 1 ? "" : "s"
          }.`
        );
      }
    })();
  }, [
    isAuthoringReportLoaded,
    isHydratingPersonalization,
    isReportLoaded,
    replaySessionQuickVisuals,
    showQuickVisualStatus,
  ]);

  const handleApplyQuickVisual = async () => {
    const authoringAvailabilityMessage =
      await getQuickVisualAuthoringAvailabilityMessage(quickVisualMode);
    if (authoringAvailabilityMessage) {
      showQuickVisualStatus(authoringAvailabilityMessage);
      return;
    }

    const roleNames = getQuickVisualOption(quickVisualType).dataRoleNames;
    const selectedRoleCount = roleNames.filter(
      (roleName) => !!quickVisualRoleSelections[roleName]
    ).length;
    const minimumRequiredRoles = Math.min(
      QUICK_VISUAL_MIN_REQUIRED_FIELDS,
      roleNames.length
    );

    if (selectedRoleCount < minimumRequiredRoles) {
      showQuickVisualStatus(
        `Select at least ${minimumRequiredRoles} fields before creating the visual.`
      );
      return;
    }

    setIsQuickVisualApplying(true);

    try {
      await syncVisibleStateToAuthoring();

      if (quickVisualMode === "change") {
        // FIX: "change" mode operates on the BASE report's active page visuals
        // (not the blank authoring page). The authoring report is used only for
        // visual creation; changeType is performed directly on the base report visual.
        if (!reportRef.current) {
          throw new Error("Report is not available.");
        }

        const basePage = await reportRef.current.getActivePage?.();
        if (!basePage || typeof basePage.getVisuals !== "function") {
          throw new Error("Active page is not available for visual changes.");
        }

        const pageVisuals = await basePage.getVisuals();
        const visualByName = new Map<string, any>(
          (pageVisuals || []).map(
            (visual: any) => [visual.name, visual] as [string, any]
          )
        );

        const targetVisual: any = quickVisualTargetVisualName
          ? visualByName.get(quickVisualTargetVisualName)
          : null;

        if (!targetVisual) {
          throw new Error("Select a visual to change.");
        }

        if (targetVisual.type !== quickVisualType) {
          if (typeof targetVisual.changeType !== "function") {
            throw new Error(
              "Change visual type is unavailable in this embed context. Enable report-authoring APIs for full quick visual support."
            );
          }

          // FIX: Remove all active data roles BEFORE changing type, per the showcase pattern.
          // This prevents stale field bindings from the old visual type carrying over.
          await removeAllActiveDataRoles(targetVisual, targetVisual.type);

          await targetVisual.changeType(quickVisualType);
        }

        await applyQuickVisualDataRoles(
          targetVisual,
          quickVisualType,
          quickVisualRoleSelections
        );
        await applyQuickVisualProperties(targetVisual, quickVisualType);
        scheduleSessionQuickVisualProfileSync();
        showQuickVisualStatus(`Updated visual "${targetVisual.title || targetVisual.name}".`);

      } else {
        // "create" mode: create a new visual on the dedicated blank authoring page
        // FIX: Use getActivePageForQuickVisual() which resolves to the blank authoring page
        const authoringPage = await getActivePageForQuickVisual();

        if (!authoringPage || typeof authoringPage.createVisual !== "function") {
          throw new Error(
            "Create visual is unavailable in this embed context. Enable report-authoring APIs for full quick visual support."
          );
        }

        // FIX: Derive layout from base report visuals for a sensible default position
        let layout = {
          x: 24,
          y: 24,
          width: 460,
          height: 300,
          displayState: {
            mode: models.VisualContainerDisplayMode.Visible,
          },
        };

        if (reportRef.current) {
          try {
            const basePage = await reportRef.current.getActivePage?.();
            if (basePage && typeof basePage.getVisuals === "function") {
              const baseVisuals = await basePage.getVisuals();
              const seedVisual = (baseVisuals || [])[0];
              if (seedVisual?.layout) {
                const sourceLayout = seedVisual.layout;
                layout = {
                  x: Math.max(8, Number(sourceLayout.x || 0) + 24),
                  y: Math.max(8, Number(sourceLayout.y || 0) + 24),
                  width: Number(sourceLayout.width) > 0 ? Number(sourceLayout.width) : 460,
                  height: Number(sourceLayout.height) > 0 ? Number(sourceLayout.height) : 300,
                  displayState: {
                    mode: models.VisualContainerDisplayMode.Visible,
                  },
                };
              }
            }
          } catch { /* non-fatal: use default layout */ }
        }

        // FIX: createVisual returns ICreateVisualResponse — always extract .visual from the response
        const createVisualResponse = await authoringPage.createVisual(quickVisualType, layout);
        const newVisual: any = createVisualResponse?.visual ?? createVisualResponse;

        if (!newVisual) {
          throw new Error(
            "Unable to create visual. Try another visual type or role selection."
          );
        }

        // FIX: Store preview visual reference so it can be cleaned up or re-used
        authoringPreviewVisualRef.current = newVisual;

        // FIX: Enable pie chart legend by default, per the showcase
        if (quickVisualType === "pieChart") {
          await applyQuickVisualProperty(newVisual, "legend", true);
        }

        // FIX: Disable legend for column and bar charts by default, per the showcase
        if (quickVisualType === "columnChart" || quickVisualType === "barChart") {
          await applyQuickVisualProperty(newVisual, "legend", false);
        }

        await applyQuickVisualDataRoles(
          newVisual,
          quickVisualType,
          quickVisualRoleSelections
        );
        await applyQuickVisualProperties(newVisual, quickVisualType);

        const roleTargetsSnapshot: Record<string, any> = {};
        roleNames.forEach((roleName) => {
          const selectedFieldKey = quickVisualRoleSelections[roleName];
          if (!selectedFieldKey) {
            return;
          }

          const selectedField = quickVisualFieldOptions.find(
            (option) => option.key === selectedFieldKey
          );
          if (selectedField?.target) {
            roleTargetsSnapshot[roleName] = selectedField.target;
          }
        });

        // FIX: Use newVisual.name (not newVisual?.name) — .name is defined on the visual object
        const runtimeVisualName: string | undefined =
          typeof newVisual.name === "string" && newVisual.name
            ? newVisual.name
            : undefined;

        appendSessionQuickVisualProfile({
          id: createBookmarkId(),
          visualType: quickVisualType,
          roleTargets: roleTargetsSnapshot,
          layout: {
            x: layout.x,
            y: layout.y,
            width: layout.width,
            height: layout.height,
          },
          properties: {
            showTitle: quickVisualShowTitle,
            showLegend: quickVisualShowLegend,
            showXAxis: quickVisualShowXAxis,
            showYAxis: quickVisualShowYAxis,
            titleText: quickVisualTitle.trim(),
            titleAlign: quickVisualTitleAlign,
          },
          runtimeVisualName,
          createdAt: new Date().toISOString(),
        });

        showQuickVisualStatus("Created a new quick visual.");
      }

      // Sync the authoring report state back into the visible report
      await syncAuthoringStateToVisible();

      setIsQuickVisualModalOpen(false);
      const persisted = await persistCurrentPersonalizationSnapshot();
      if (persisted) {
        setSaveStatus("saved");
        setLastSaved(new Date());
        window.setTimeout(() => setSaveStatus("idle"), 1500);
      }
      triggerAutoSaveRevision();
      await refreshQuickVisualTargets();
    } catch (error) {
      console.error("Error applying quick visual changes", error);
      
      let message = "Failed to apply quick visual changes.";
      
      if (error instanceof Error) {
        const errorMsg = error.message.toLowerCase();
        
        if (errorMsg.includes("unavailable in this embed context")) {
          message = `⚠️ Edit access denied. Your user account doesn't have edit permissions in this workspace. Contact your Power BI admin to be added as a Member or Admin.`;
        } else if (errorMsg.includes("create visual")) {
          message = `Report authoring APIs not enabled. Ensure you have edit permissions and the workspace allows visual creation.`;
        } else if (errorMsg.includes("change visual")) {
          message = `Cannot change visual type. You may not have edit permissions for this report.`;
        } else if (errorMsg.includes("unable to create visual")) {
          message = `Visual creation failed. Check that the field selections are valid and try another visual type.`;
        } else {
          message = error.message;
        }
      }
      
      showQuickVisualStatus(message);
    } finally {
      setIsQuickVisualApplying(false);
    }
  };

  const openSaveBookmarkModal = () => {
    if (!reportRef.current) {
      showBookmarkStatus("Report is still loading. Try again in a moment.");
      return;
    }

    // Generate a suggested default name
    const nextNumber = bookmarkProfiles.length + 1;
    setBookmarkNameInput(`My View ${nextNumber}`);
    setIsBookmarkModalOpen(true);
  };

  const closeSaveBookmarkModal = () => {
    if (isSavingBookmark) {
      return;
    }

    setIsBookmarkModalOpen(false);
    setBookmarkNameInput("");
  };

  const handleSaveBookmark = async () => {
    const bookmarkName = bookmarkNameInput.trim();
    if (!bookmarkName) {
      showBookmarkStatus("Bookmark name is required.");
      return;
    }

    if (!reportRef.current) {
      showBookmarkStatus("Report is still loading. Try again in a moment.");
      return;
    }

    setIsSavingBookmark(true);

    try {
      setSaveStatus("saving");
      let bookmarkStateJson = "";

      await syncSessionQuickVisualProfilesFromPage();

      if (reportRef.current?.bookmarksManager?.capture) {
        const captured = await reportRef.current.bookmarksManager.capture({
          personalizeVisuals: true,
        });
        bookmarkStateJson = captured?.state || "";
      }

      if (!bookmarkStateJson) {
        throw new Error("Unable to capture the current report view.");
      }

      // Capture current layout customizer state
      const layoutState = layoutCustomizerRef?.current?.getLayoutState?.() ?? undefined;

      const nextBookmark = upsertCapturedBookmark(
        bookmarkName,
        bookmarkStateJson,
        false,
        "saveView",
        layoutState
      );
      if (!nextBookmark) {
        throw new Error("Unable to save captured bookmark state.");
      }

      const nextSelectionId = toSavedBookmarkSelectionId(nextBookmark.id);
      setSelectedBookmarkId(nextSelectionId);
      window.localStorage.setItem(selectedBookmarkStorageKey, nextSelectionId);

      showBookmarkStatus(
        `✓ Saved view "${bookmarkName}" in your personal library.`
      );

      setSaveStatus("saved");
      setLastSaved(new Date(nextBookmark.updatedAt));
      window.setTimeout(() => setSaveStatus("idle"), 1500);
      setIsBookmarkModalOpen(false);
      setBookmarkNameInput("");
    } catch (error) {
      console.error("Error saving bookmark profile:", error);
      setSaveStatus("error");
      showBookmarkStatus("Failed to save view.");
    } finally {
      setIsSavingBookmark(false);
    }
  };

  const loadBookmarkById = async (bookmarkIdToLoad: string) => {
    if (!bookmarkIdToLoad) {
      return;
    }

    if (isOriginalReportSelection(bookmarkIdToLoad)) {
      const restored = await restoreOriginalReportState();
      if (!restored) {
        showBookmarkStatus("Unable to restore original report view right now.");
        return;
      }

      // Reset layout customizer to default (all visuals selected, default layout)
      if (layoutCustomizerRef?.current?.resetToDefault) {
        layoutCustomizerRef.current.resetToDefault();
      }

      window.localStorage.setItem(selectedBookmarkStorageKey, bookmarkIdToLoad);
      showBookmarkStatus("Loaded original report view.");
      triggerAutoSaveRevision();
      return;
    }

    await captureOriginalReportStateIfMissing();

    const selectedSavedBookmarkId = getSavedBookmarkIdFromSelection(
      bookmarkIdToLoad
    );
    const selectedReportBookmarkName = getReportBookmarkNameFromSelection(
      bookmarkIdToLoad
    );

    if (selectedSavedBookmarkId) {
      const selectedBookmark = bookmarkProfiles.find(
        (bookmark) => bookmark.id === selectedSavedBookmarkId
      );
      if (!selectedBookmark) {
        showBookmarkStatus("Selected bookmark was not found.");
        return;
      }

      // Apply PBI bookmark state first (filters, slicers, etc.)
      await applyBookmarkProfile(selectedBookmark);
      await syncVisibleStateToAuthoring();

      // Then restore layout customizer state (applies custom layout on top)
      if (selectedBookmark.layoutState && layoutCustomizerRef?.current?.setLayoutState) {
        layoutCustomizerRef.current.setLayoutState(selectedBookmark.layoutState);
      } else if (layoutCustomizerRef?.current?.resetToDefault) {
        // No layout state saved — reset to default
        layoutCustomizerRef.current.resetToDefault();
      }

      window.localStorage.setItem(selectedBookmarkStorageKey, bookmarkIdToLoad);
      showBookmarkStatus(`Loaded bookmark "${selectedBookmark.name}"`);
      return;
    }

    if (
      selectedReportBookmarkName &&
      typeof reportRef.current?.bookmarksManager?.apply === "function"
    ) {
      await reportRef.current.bookmarksManager.apply(selectedReportBookmarkName);
      const selectedReportBookmark = reportBookmarks.find(
        (bookmark) => bookmark.name === selectedReportBookmarkName
      );
      const bookmarkLabel =
        selectedReportBookmark?.displayName || selectedReportBookmarkName;
      window.localStorage.setItem(selectedBookmarkStorageKey, bookmarkIdToLoad);
      showBookmarkStatus(`Loaded bookmark "${bookmarkLabel}"`);
      const activePage = await reportRef.current?.getActivePage?.();
      if (activePage?.name) {
        setCurrentPage(activePage.name);
      }
      await syncVisibleStateToAuthoring();
      return;
    }

    showBookmarkStatus("Selected bookmark was not found.");
  };

  const handleDeleteSelectedBookmark = () => {
    if (!selectedBookmarkId) {
      return;
    }

    if (isOriginalReportSelection(selectedBookmarkId)) {
      showBookmarkStatus("Original report view cannot be deleted.");
      return;
    }

    const selectedSavedBookmarkId = getSavedBookmarkIdFromSelection(
      selectedBookmarkId
    );
    if (!selectedSavedBookmarkId) {
      showBookmarkStatus("Report bookmarks cannot be deleted from the app.");
      return;
    }

    const selectedBookmark = bookmarkProfiles.find(
      (bookmark) => bookmark.id === selectedSavedBookmarkId
    );

    const nextBookmarks = bookmarkProfiles.filter(
      (bookmark) => bookmark.id !== selectedSavedBookmarkId
    );
    persistBookmarks(nextBookmarks);

    const fallbackSelectionId = nextBookmarks[0]?.id
      ? toSavedBookmarkSelectionId(nextBookmarks[0].id)
      : reportBookmarks[0]?.name
        ? toReportBookmarkSelectionId(reportBookmarks[0].name)
        : "";
    setSelectedBookmarkId(fallbackSelectionId);
    if (fallbackSelectionId) {
      window.localStorage.setItem(selectedBookmarkStorageKey, fallbackSelectionId);
    } else {
      window.localStorage.removeItem(selectedBookmarkStorageKey);
    }

    if (selectedBookmark) {
      showBookmarkStatus(`Deleted bookmark "${selectedBookmark.name}"`);
    }
  };

  const handleFiltersApplied = useCallback(async () => {
    try {
      if (!reportRef.current) {
        return;
      }

      const state = await getReportPersonalizationState(reportRef.current);
      setCurrentPage(state.currentPage);
      triggerAutoSaveRevision();
    } catch (error) {
      console.error("Error handling filter changes:", error);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [reportRef]);

  const handlePageChanged = useCallback((event?: any) => {
    try {
      if (event?.detail?.newPage?.name) {
        setCurrentPage(event.detail.newPage.name);
      }

      triggerAutoSaveRevision();
    } catch (error) {
      console.error("Error handling page change:", error);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const defaultEmbedReportEventHandlers: Map<string, EventHandler> = useMemo(
    () =>
      new Map([
        [
          "filtersApplied",
          () => {
            handleFiltersApplied();
          },
        ],
        [
          "pageChanged",
          (event?: any) => {
            handlePageChanged(event);
          },
        ],
        [
          "selectionChanged",
          () => {
            scheduleSessionQuickVisualProfileSync();
            triggerAutoSaveRevision();
          },
        ],
        [
          "visualRendered",
          () => {
            scheduleSessionQuickVisualProfileSync();
          },
        ],
        [
          "bookmarkApplied",
          async (event?: any) => {
            try {
              const appliedBookmarkName = getBookmarkNameFromAppliedEvent(event);

              const latestReportBookmarks = await loadReportBookmarks(true);

              const matchedReportBookmark = appliedBookmarkName
                ? (
                    latestReportBookmarks.length > 0
                      ? latestReportBookmarks
                      : reportBookmarks
                  ).find(
                    (bookmark) =>
                      bookmark.name === appliedBookmarkName ||
                      bookmark.displayName === appliedBookmarkName
                  )
                : null;

              if (matchedReportBookmark) {
                const nextSelectionId = toReportBookmarkSelectionId(
                  matchedReportBookmark.name
                );
                setSelectedBookmarkId(nextSelectionId);
                window.localStorage.setItem(
                  selectedBookmarkStorageKey,
                  nextSelectionId
                );
                return;
              }

              const existingBookmarkByName = appliedBookmarkName
                ? bookmarkProfilesRef.current.find(
                    (bookmark) =>
                      bookmark.name.trim().toLowerCase() ===
                      appliedBookmarkName.trim().toLowerCase()
                  )
                : null;

              if (existingBookmarkByName) {
                const nextSelectionId = toSavedBookmarkSelectionId(
                  existingBookmarkByName.id
                );
                setSelectedBookmarkId(nextSelectionId);
                window.localStorage.setItem(
                  selectedBookmarkStorageKey,
                  nextSelectionId
                );
                showBookmarkStatus(`Synced bookmark "${existingBookmarkByName.name}"`);
                return;
              }

              if (appliedBookmarkName && latestReportBookmarks.length === 0) {
                showBookmarkStatus(`Applied bookmark "${appliedBookmarkName}"`);
                return;
              }

              if (!reportRef.current?.bookmarksManager?.capture) {
                return;
              }

              const captured = await reportRef.current.bookmarksManager.capture({
                personalizeVisuals: true,
              });
              const bookmarkStateJson = captured?.state || "";
              if (!bookmarkStateJson) {
                return;
              }

              const existingBookmarkByState = bookmarkProfilesRef.current.find(
                (bookmark) =>
                  bookmark.state === bookmarkStateJson ||
                  bookmark.bookmarkStateJson === bookmarkStateJson
              );

              if (existingBookmarkByState) {
                const nextSelectionId = toSavedBookmarkSelectionId(
                  existingBookmarkByState.id
                );
                setSelectedBookmarkId(nextSelectionId);
                window.localStorage.setItem(
                  selectedBookmarkStorageKey,
                  nextSelectionId
                );
                showBookmarkStatus(
                  `Synced bookmark "${existingBookmarkByState.name}"`
                );
                return;
              }

              const resolvedBookmarkName = toReadableBookmarkName(
                captured?.displayName,
                appliedBookmarkName
              );

              const syncedBookmark = upsertCapturedBookmark(
                resolvedBookmarkName,
                bookmarkStateJson,
                true,
                "syncApplied"
              );
              if (syncedBookmark) {
                showBookmarkStatus(`Synced bookmark "${syncedBookmark.name}"`);
              }
            } catch (error) {
              console.warn("Unable to sync applied bookmark into dropdown", error);
            }
          },
        ],
        [
          "commandTriggered",
          (event?: any) => {
            const commandName =
              event?.detail?.command || event?.detail?.name || event?.detail?.id;
            const normalizedCommandName = String(commandName || "").toLowerCase();

            if (normalizedCommandName === "createquickvisual") {
              void openQuickVisualModal("create");
              return;
            }

            if (normalizedCommandName !== "changevisual") {
              return;
            }

            const visualName =
              event?.detail?.visual?.name || event?.detail?.data?.visual?.name;
            void openQuickVisualModal("change", visualName);
          },
        ],
        [
          "buttonClicked",
          () => {
            void openQuickVisualModal("create");
          },
        ],
        [
          "loaded",
          () => {},
        ],
        [
          "rendered",
          async () => {
            if (!isReportLoaded) {
              setIsReportLoaded(true);
            }

            await loadReportBookmarks();

            try {
              const activePage = await reportRef.current?.getActivePage?.();
              if (activePage?.name) {
                setCurrentPage(activePage.name);
              }
              await captureOriginalReportStateIfMissing();
              if (onReportReady && reportRef.current && activePage) {
                onReportReady(reportRef.current, activePage);
              }
            } catch (error) {
              console.warn("Unable to read active page on render", error);
            }
          },
        ],
      ]),
    [
      handleFiltersApplied,
      handlePageChanged,
      scheduleSessionQuickVisualProfileSync,
      triggerAutoSaveRevision,
      loadReportBookmarks,
      reportBookmarks,
      selectedBookmarkStorageKey,
      showBookmarkStatus,
      reportRef,
      upsertCapturedBookmark,
      openQuickVisualModal,
      isReportLoaded,
      captureOriginalReportStateIfMissing,
    ]
  );

  const mergedEmbedReportEventHandlers: Map<string, EventHandler> = useMemo(
    () => {
      const merged = new Map(defaultEmbedReportEventHandlers);
      if (embedReportEventHandlers && embedReportEventHandlers.size > 0) {
        embedReportEventHandlers.forEach((handler, eventName) => {
          if (eventName) {
            merged.set(eventName, handler);
          }
        });
      }
      return merged;
    },
    [defaultEmbedReportEventHandlers, embedReportEventHandlers]
  );

  // FIX: On authoring report "loaded", navigate to pages[1] (the blank authoring page)
  // and store it in authoringPageRef. This matches the showcase pattern exactly:
  //   pages[1].setActive() → visualCreatorShowcaseState.page = pages[1]
  const authoringEmbedReportEventHandlers: Map<string, EventHandler> = useMemo(
    () =>
      new Map([
        [
          "loaded",
          async () => {
            setAuthoringEmbedError(null);
            try {
              if (
                authoringReportRef.current &&
                typeof authoringReportRef.current.getPages === "function"
              ) {
                const pages = await authoringReportRef.current.getPages();
                if (Array.isArray(pages) && pages.length > 1) {
                  // pages[1] is the blank page designated for visual authoring
                  await pages[1].setActive();
                  authoringPageRef.current = pages[1];
                } else if (Array.isArray(pages) && pages.length > 0) {
                  authoringPageRef.current = pages[0];
                }
              }
            } catch (err) {
              console.warn("Unable to set authoring report to blank page on load", err);
            }
            setIsAuthoringReportLoaded(true);
          },
        ],
        [
          "rendered",
          () => {
            if (!isAuthoringReportLoaded) {
              setIsAuthoringReportLoaded(true);
            }
            setAuthoringEmbedError(null);
          },
        ],
        [
          "error",
          (event?: any) => {
            const readableMessage = toAuthoringEmbedErrorMessage(event);
            setAuthoringEmbedError(readableMessage);
            setIsAuthoringReportLoaded(false);
            console.warn("Authoring report embed error", readableMessage, event?.detail || event);
          },
        ],
      ]),
    [isAuthoringReportLoaded, toAuthoringEmbedErrorMessage]
  );

  const handleReportLoadAttachment = useCallback(
    (report: any) => {
      if (onReportLoadReportAttachmentFunction) {
        onReportLoadReportAttachmentFunction(report);
      }
    },
    [onReportLoadReportAttachmentFunction]
  );

  const handleAuthoringReportLoadAttachment = useCallback(() => {
    // No-op for authoring embed
  }, []);

  const reportSettings: models.ISettings = {
    bars: {
      actionBar: {
        visible: false,
      },
      statusBar: {
        visible: false,
      },
    },
    panes: {
      bookmarks: {
        visible: false,
      },
      filters: {
        expanded: false,
        visible: false,
      },
      pageNavigation: {
        visible: false,
      },
      fields: {
        expanded: false,
        visible: false,
      },
      visualizations: {
        expanded: false,
        visible: false,
      },
    },
    navContentPaneEnabled: false,
    bookmarksPaneEnabled: false,
    personalBookmarksEnabled: false,
    persistentFiltersEnabled: true,
    background: models.BackgroundType.Transparent,
    visualRenderedEvents: true,
    ...embedReportSettingsOverrides,
  };

  const selectedQuickVisualOption = getQuickVisualOption(quickVisualType);
  const selectedQuickVisualRolesCount = selectedQuickVisualOption.dataRoleNames.filter(
    (roleName) => !!quickVisualRoleSelections[roleName]
  ).length;
  const requiredQuickVisualRoles = Math.min(
    QUICK_VISUAL_MIN_REQUIRED_FIELDS,
    selectedQuickVisualOption.dataRoleNames.length
  );
  const canApplyQuickVisual =
    quickVisualFieldOptions.length > 0 &&
    selectedQuickVisualRolesCount >= requiredQuickVisualRoles &&
    (quickVisualMode !== "change" || !!quickVisualTargetVisualName);
  const showLegendToggle = selectedQuickVisualOption.properties.includes(
    "legend"
  );
  const showXAxisToggle = selectedQuickVisualOption.properties.includes("xAxis");
  const showYAxisToggle = selectedQuickVisualOption.properties.includes("yAxis");
  const selectedSavedBookmarkId = getSavedBookmarkIdFromSelection(
    selectedBookmarkId
  );

  return (
    <div className="personalized-editable-report">
      <div className="report-toolbar">
        <div className="toolbar-row">
          <div className="toolbar-cluster primary-actions">
            <button
              onClick={openSaveBookmarkModal}
              className="btn btn-bookmark-save"
              title="Save current view as a bookmark state"
            >
              Save View
            </button>
          </div>

          <div className="toolbar-cluster bookmark-controls">
            <div className="bookmark-selector-wrapper">
              <select
                className="bookmark-select"
                value={selectedBookmarkId}
                onChange={(e) => {
                  const newValue = e.target.value;
                  setSelectedBookmarkId(newValue);
                  // Auto-load on selection
                  if (newValue) {
                    void loadBookmarkById(newValue);
                  }
                }}
              >
                <option value="">Select bookmark</option>
                <option value={ORIGINAL_REPORT_SELECTION_ID}>Original report view</option>
                {reportBookmarks.length > 0 && (
                  <optgroup label="Power BI report bookmarks">
                    {reportBookmarks.map((bookmark) => (
                      <option
                        key={`report_${bookmark.name}`}
                        value={toReportBookmarkSelectionId(bookmark.name)}
                      >
                        {bookmark.displayName || bookmark.name}
                      </option>
                    ))}
                  </optgroup>
                )}
                {bookmarkProfiles.length > 0 && (
                  <optgroup label="Captured and synced personal views (Power BI state)">
                    {bookmarkProfiles.map((bookmark) => (
                      <option
                        key={`saved_${bookmark.id}`}
                        value={toSavedBookmarkSelectionId(bookmark.id)}
                      >
                        {bookmark.name}
                      </option>
                    ))}
                  </optgroup>
                )}
                {reportBookmarks.length === 0 && bookmarkProfiles.length === 0 && (
                  <option value="" disabled>
                    No bookmarks yet. Use Save View.
                  </option>
                )}
              </select>
              <button
                onClick={handleDeleteSelectedBookmark}
                className="btn-delete-bookmark"
                disabled={!selectedSavedBookmarkId}
                title="Delete selected bookmark"
              >
                ✕
              </button>
              <button
                onClick={() => {
                  if (selectedBookmarkId) {
                    void loadBookmarkById(selectedBookmarkId);
                  }
                }}
                className="btn-refresh-bookmarks"
                title="Refresh selected bookmark"
              >
                ⟳
              </button>
            </div>
          </div>

          <div className="toolbar-cluster edit-actions">
            {toggleButton}
            <ReportEditor
              reportRef={reportRef}
              reportId={reportId}
              userId={userId}
              workspaceId={workspaceId}
              allowEdit={allowEdit}
              onSave={handleFiltersApplied}
              onModeChange={(mode) => {
                if (mode !== "edit") {
                  setIsQuickVisualModalOpen(false);
                }
              }}
            />
          </div>
        </div>

        <div className="toolbar-row toolbar-meta">
          <label className="auto-save-toggle">
            <input
              type="checkbox"
              checked={autoSaveEnabled}
              onChange={(e) => setAutoSaveEnabled(e.target.checked)}
            />
            Auto-save personal view
          </label>

          {saveStatus === "saving" && (
            <span className="save-status saving">💾 Saving personalization...</span>
          )}
          {saveStatus === "saved" && (
            <span className="save-status saved">✓ Personalization saved</span>
          )}
          {saveStatus === "error" && (
            <span className="save-status error">✗ Save failed</span>
          )}

          {lastSaved && (
            <span className="last-saved">
              Last saved: {lastSaved.toLocaleTimeString()}
            </span>
          )}

          {bookmarkStatus && (
            <span className="bookmark-status">{bookmarkStatus}</span>
          )}

          {quickVisualStatus && (
            <span className="quick-visual-status">{quickVisualStatus}</span>
          )}
        </div>
      </div>

      {layoutControls && (
        <div className="toolbar-row toolbar-layout-controls" style={{ display: "flex", gap: 8, alignItems: "center", padding: "8px 0" }}>
          {layoutControls}
        </div>
      )}

      <div className="report-container">
        <EmbedReport
          reportRef={reportRef}
          reportId={reportId}
          embedUrl={embedUrl}
          embedReportEventHandlers={mergedEmbedReportEventHandlers}
          reportSettings={reportSettings}
          accessToken={accessToken}
          tokenType={tokenType}
          onReportLoadReportAttachmentFunction={handleReportLoadAttachment}
        />
      </div>

      {allowEdit && (
        <div className="authoring-report-container" aria-hidden="true">
          <EmbedReport
            reportRef={authoringReportRef}
            reportId={reportId}
            embedUrl={embedUrl}
            embedReportEventHandlers={authoringEmbedReportEventHandlers}
            reportSettings={reportSettings}
            reportCssClassName="authoring-report-embed"
            accessToken={accessToken}
            tokenType={tokenType}
            onReportLoadReportAttachmentFunction={handleAuthoringReportLoadAttachment}
          />
        </div>
      )}

      {isQuickVisualModalOpen && (
        <div
          className="quick-visual-modal-backdrop"
          onClick={closeQuickVisualModal}
        >
          <div
            className="quick-visual-modal"
            onClick={(e) => e.stopPropagation()}
          >
            <h3>
              {quickVisualMode === "change"
                ? "Change Existing Visual"
                : "Create Quick Visual"}
            </h3>
            <p>
              {quickVisualMode === "change"
                ? "Update visual type and formatting for an existing visual."
                : "Create a new visual using selected role fields from this report context."}
            </p>

            <div className="quick-visual-grid">
              {quickVisualMode === "change" && (
                <label className="quick-visual-field">
                  <span>Target visual</span>
                  <select
                    value={quickVisualTargetVisualName}
                    onChange={(e) => {
                      const nextVisualName = e.target.value;
                      const nextVisual = quickVisualTargets.find(
                        (visual) => visual.name === nextVisualName
                      );
                      const nextVisualType =
                        nextVisual?.type &&
                        QUICK_VISUAL_OPTIONS.some(
                          (option) => option.name === nextVisual.type
                        )
                          ? nextVisual.type
                          : quickVisualType;

                      setQuickVisualTargetVisualName(nextVisualName);
                      if (
                        nextVisual?.type &&
                        QUICK_VISUAL_OPTIONS.some((option) => option.name === nextVisual.type)
                      ) {
                        const option = getQuickVisualOption(nextVisualType);
                        setQuickVisualType(nextVisualType);
                        setQuickVisualShowLegend(option.properties.includes("legend"));
                        setQuickVisualShowXAxis(option.properties.includes("xAxis"));
                        setQuickVisualShowYAxis(option.properties.includes("yAxis"));
                      }

                      void syncQuickVisualSelectionsFromVisual(
                        nextVisualName,
                        nextVisualType
                      );
                    }}
                  >
                    {quickVisualTargets.map((visual) => (
                      <option key={visual.name} value={visual.name}>
                        {visual.title}
                      </option>
                    ))}
                  </select>
                </label>
              )}

              <label className="quick-visual-field">
                <span>Visual type</span>
                <select
                  value={quickVisualType}
                  onChange={(e) => {
                    const nextType = e.target.value;
                    const option = getQuickVisualOption(nextType);
                    setQuickVisualType(nextType);
                    setQuickVisualShowLegend(option.properties.includes("legend"));
                    setQuickVisualShowXAxis(option.properties.includes("xAxis"));
                    setQuickVisualShowYAxis(option.properties.includes("yAxis"));
                    setQuickVisualRoleSelections((previous) => {
                      const next: QuickVisualRoleSelections = {};
                      option.dataRoleNames.forEach((roleName) => {
                        if (previous[roleName]) {
                          next[roleName] = previous[roleName];
                        }
                      });
                      return next;
                    });
                  }}
                >
                  {QUICK_VISUAL_OPTIONS.map((option) => (
                    <option key={option.name} value={option.name}>
                      {option.label}
                    </option>
                  ))}
                </select>
              </label>

              {selectedQuickVisualOption.dataRoleNames.map((roleName) => (
                <label className="quick-visual-field" key={roleName}>
                  <span>{QUICK_VISUAL_ROLE_LABELS[roleName] || roleName}</span>
                  <select
                    value={quickVisualRoleSelections[roleName] || ""}
                    onChange={(e) => {
                      const selectedFieldKey = e.target.value;
                      setQuickVisualRoleSelections((previous) => ({
                        ...previous,
                        [roleName]: selectedFieldKey,
                      }));
                    }}
                  >
                    <option value="">Select option</option>
                    {quickVisualFieldOptions.map((fieldOption) => (
                      <option key={`${roleName}_${fieldOption.key}`} value={fieldOption.key}>
                        {fieldOption.label}
                      </option>
                    ))}
                  </select>
                </label>
              ))}

              {quickVisualFieldOptions.length === 0 && (
                <div className="quick-visual-empty-fields full-width">
                  No data fields are available for this report in the current context.
                </div>
              )}

              <label className="quick-visual-field full-width">
                <span>Title</span>
                <input
                  type="text"
                  value={quickVisualTitle}
                  onChange={(e) => setQuickVisualTitle(e.target.value)}
                  placeholder="Type a personalized visual title"
                />
              </label>

              <div className="quick-visual-field">
                <span>Title alignment</span>
                <select
                  value={quickVisualTitleAlign}
                  onChange={(e) => setQuickVisualTitleAlign(e.target.value)}
                >
                  <option value="left">Left</option>
                  <option value="center">Center</option>
                  <option value="right">Right</option>
                </select>
              </div>
            </div>

            <div className="quick-visual-toggles">
              <label>
                <input
                  type="checkbox"
                  checked={quickVisualShowTitle}
                  onChange={(e) => setQuickVisualShowTitle(e.target.checked)}
                />
                Show title
              </label>

              <label className={!showLegendToggle ? "disabled" : ""}>
                <input
                  type="checkbox"
                  checked={quickVisualShowLegend}
                  disabled={!showLegendToggle}
                  onChange={(e) => setQuickVisualShowLegend(e.target.checked)}
                />
                Show legend
              </label>

              <label className={!showXAxisToggle ? "disabled" : ""}>
                <input
                  type="checkbox"
                  checked={quickVisualShowXAxis}
                  disabled={!showXAxisToggle}
                  onChange={(e) => setQuickVisualShowXAxis(e.target.checked)}
                />
                Show X axis
              </label>

              <label className={!showYAxisToggle ? "disabled" : ""}>
                <input
                  type="checkbox"
                  checked={quickVisualShowYAxis}
                  disabled={!showYAxisToggle}
                  onChange={(e) => setQuickVisualShowYAxis(e.target.checked)}
                />
                Show Y axis
              </label>
            </div>

            <div className="quick-visual-modal-actions">
              <button
                className="btn btn-modal-cancel"
                onClick={closeQuickVisualModal}
                disabled={isQuickVisualApplying}
              >
                Cancel
              </button>
              <button
                className="btn btn-quick-visual-apply"
                onClick={() => {
                  void handleApplyQuickVisual();
                }}
                disabled={isQuickVisualApplying || !canApplyQuickVisual}
              >
                {isQuickVisualApplying
                  ? "Applying..."
                  : quickVisualMode === "change"
                    ? "Update Visual"
                    : "Create Visual"}
              </button>
            </div>
          </div>
        </div>
      )}

      {isBookmarkModalOpen && (
        <div className="bookmark-modal-backdrop" onClick={closeSaveBookmarkModal}>
          <div className="bookmark-modal" onClick={(e) => e.stopPropagation()}>
            <h3>Save Personal View</h3>
            <p>
              Save the current filters, page, and visual state as a bookmark.
            </p>

            <input
              type="text"
              className="bookmark-modal-input"
              value={bookmarkNameInput}
              onChange={(e) => setBookmarkNameInput(e.target.value)}
              placeholder="Enter bookmark name"
              autoFocus
              onFocus={(e) => e.target.select()}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  void handleSaveBookmark();
                }
                if (e.key === "Escape") {
                  closeSaveBookmarkModal();
                }
              }}
            />

            <div className="bookmark-modal-actions">
              <button
                className="btn btn-modal-cancel"
                onClick={closeSaveBookmarkModal}
                disabled={isSavingBookmark}
              >
                Cancel
              </button>
              <button
                className="btn btn-modal-save"
                onClick={() => {
                  void handleSaveBookmark();
                }}
                disabled={isSavingBookmark || !bookmarkNameInput.trim()}
              >
                {isSavingBookmark ? "Saving..." : "Save Bookmark"}
              </button>
            </div>
          </div>
        </div>
      )}

      {loading && (
        <div className="loading-overlay">
          <div className="spinner">Loading...</div>
        </div>
      )}
    </div>
  );
};

