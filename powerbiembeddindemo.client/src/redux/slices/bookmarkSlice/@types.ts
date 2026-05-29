export interface VisualSnapshot {
  visualType: string;
  dataRoles: Record<string, { table: string; name: string; schema: string; column?: string; measure?: string }>;
  properties: {
    legend: boolean;
    xAxis: boolean;
    yAxis: boolean;
    title: boolean;
    titleText: string | null;
    titleAlign: string | null;
  };
}

export interface BookmarkedVisual {
  id: string;
  name: string;
  visuals: VisualSnapshot[]; // all visuals captured at save time
  createdAt: number;
  layoutState?: {
    selectedVisuals: string[];
    layoutType: string;
    columns: number;
    spanType: number;
    isCustomLayoutActive: boolean;
  };
}

export interface BookmarkSliceState {
  bookmarks: BookmarkedVisual[];
}
