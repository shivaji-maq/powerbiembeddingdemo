// Visual types with their data roles configuration
export const visualTypeToDataRoles = [
  { name: "columnChart", displayName: "Column chart", dataRoles: ["Axis", "Values", "Tooltips"], dataRoleNames: ["Category", "Y", "Tooltips"] },
  { name: "areaChart", displayName: "Area chart", dataRoles: ["Axis", "Legend", "Values"], dataRoleNames: ["Category", "Series", "Y"] },
  { name: "barChart", displayName: "Bar chart", dataRoles: ["Axis", "Values", "Tooltips"], dataRoleNames: ["Category", "Y", "Tooltips"] },
  { name: "pieChart", displayName: "Pie chart", dataRoles: ["Legend", "Values", "Tooltips"], dataRoleNames: ["Category", "Y", "Tooltips"] },
  { name: "lineChart", displayName: "Line chart", dataRoles: ["Axis", "Legend", "Values"], dataRoleNames: ["Category", "Series", "Y"] },
];

// Schemas for Power BI visuals API
export const schemas = {
  column: "http://powerbi.com/product/schema#column",
  measure: "http://powerbi.com/product/schema#measure",
  property: "http://powerbi.com/product/schema#property",
  default: "http://powerbi.com/product/schema#default",
};

// Available visual properties
export const showcaseProperties = ["legend", "xAxis", "yAxis"];

// Title-related properties
export const titleProperties = ["title", "titleText", "titleAlign"];

// Properties available for each visual type
export const visualTypeProperties: Record<string, string[]> = {
  columnChart: ["xAxis", "yAxis"],
  areaChart: ["legend", "xAxis", "yAxis"],
  barChart: ["xAxis", "yAxis"],
  pieChart: ["legend"],
  lineChart: ["legend", "xAxis", "yAxis"],
};

// Convert property name to Power BI property selector
export function propertyToSelector(propertyName: string): { objectName: string; propertyName: string } {
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
      return { objectName: propertyName, propertyName: "visible" };
  }
}

// Layout constants
export const VISUAL_CREATOR_SHOWCASE = {
  MARGIN: 16,
  COLUMNS: 3,
  VISUAL_ASPECT_RATIO: 0.65,
};
