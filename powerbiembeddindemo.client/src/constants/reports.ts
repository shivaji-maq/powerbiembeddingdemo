export interface ReportToEmbed {
  reportId: string;
  workspaceId: string;
  datasetId: string;
  name: string;
  isSecureEmbedded: boolean;
  embeddingIframe?: string;
  globalDateFilter?: {
    table: string;
    column: string;
  };
}

export const reportsToEmbed: ReportToEmbed[] = [
  {
    reportId: "909229f0-1c69-4aa8-bbb0-b60ed2e0c3ba",
    workspaceId: "a0f458be-e2a3-47c8-85f7-75d7b7f6034f",
    datasetId: "",
    name: "Contoso Sample Report",
    isSecureEmbedded: false,
    globalDateFilter: {
      table: "vw_DimOrder",
      column: "Order Date",
    },
  },
  {
    reportId: "ddfd8090-0b57-4635-959a-3362aa2e444b",
    workspaceId: "a0f458be-e2a3-47c8-85f7-75d7b7f6034f",
    datasetId: "",
    name: "Competitive Marketing Analysis",
    isSecureEmbedded: false,
    globalDateFilter: {
      table: "Date",
      column: "Date",
    },
  },
  {
    reportId: "b1189c3b-2f90-4ac1-a08c-8cccc9168bc6",
    workspaceId: "a0f458be-e2a3-47c8-85f7-75d7b7f6034f",
    datasetId: "",
    name: "Artificial Intelligence Sample",
    isSecureEmbedded: false,
    globalDateFilter: {
      table: "Opportunity Calendar",
      column: "Date",
    },
  },
  {
    reportId: "f2496a3d-b253-4ab8-94d4-2f69816551cd",
    workspaceId: "a0f458be-e2a3-47c8-85f7-75d7b7f6034f",
    datasetId: "",
    name: "Writeback Flow",
    isSecureEmbedded: true,
    embeddingIframe: `<iframe title="Translytical flow for writeback" width="1140" height="541.25" src="https://app.powerbi.com/reportEmbed?reportId=f2496a3d-b253-4ab8-94d4-2f69816551cd&autoAuth=true&ctid=e4d98dd2-9199-42e5-ba8b-da3e763ede2e&language=selectLanguage" frameborder="0" allowFullScreen="true"></iframe>`,
  },
  // {
  //   reportId: "a4a308db-00fc-482a-9d48-27b787755580",
  //   workspaceId: "a0f458be-e2a3-47c8-85f7-75d7b7f6034f",
  //   datasetId: "",
  //   name: "User Context Aware Store Sales",
  //   isSecureEmbedded: true,
  //   embeddingIframe: `<iframe title="User Context Aware Column Demo" width="1140" height="541.25" src="https://app.powerbi.com/reportEmbed?reportId=4031228f-f9dc-4aa9-8905-f70414f59698&autoAuth=true&ctid=e4d98dd2-9199-42e5-ba8b-da3e763ede2e&language=selectLanguage" frameborder="0" allowFullScreen="true"></iframe>`,
  // },

  {
    reportId: "a4a308db-00fc-482a-9d48-27b787755580",
    workspaceId: "a0f458be-e2a3-47c8-85f7-75d7b7f6034f",
    datasetId: "",
    name: "User Context Aware Store Sales",
    isSecureEmbedded: false,
    embeddingIframe: `<iframe title="User Context Aware Column Demo" width="1140" height="541.25" src="https://app.powerbi.com/reportEmbed?reportId=4031228f-f9dc-4aa9-8905-f70414f59698&autoAuth=true&ctid=e4d98dd2-9199-42e5-ba8b-da3e763ede2e&language=selectLanguage" frameborder="0" allowFullScreen="true"></iframe>`,
    globalDateFilter: {
      table: "Time",
      column: "Date",
    },
  },
];
