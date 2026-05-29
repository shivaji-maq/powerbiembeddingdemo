export interface ReportToEmbed {
  reportId: string;
  workspaceId: string;
  datasetId: string;
  name: string;
  isSecureEmbedded: boolean;
  embeddingIframe?: string;
}

export const reportsToEmbed: ReportToEmbed[] = [
  {
    reportId: "e479fe39-9d8e-4c72-82c2-eff7d065ce99",
    workspaceId: "a0f458be-e2a3-47c8-85f7-75d7b7f6034f",
    datasetId: "",
    name: "Contoso Sales Report",
    isSecureEmbedded: false,
  },
  {
    reportId: "f2496a3d-b253-4ab8-94d4-2f69816551cd",
    workspaceId: "a0f458be-e2a3-47c8-85f7-75d7b7f6034f",
    datasetId: "",
    name: "Translytical flow Demo",
    isSecureEmbedded: true,
    embeddingIframe: `<iframe title="Translytical flow for writeback" width="1140" height="541.25" src="https://app.powerbi.com/reportEmbed?reportId=f2496a3d-b253-4ab8-94d4-2f69816551cd&autoAuth=true&ctid=e4d98dd2-9199-42e5-ba8b-da3e763ede2e" frameborder="0" allowFullScreen="true"></iframe>`,
  },
  {
    reportId: "a4a308db-00fc-482a-9d48-27b787755580",
    workspaceId: "a0f458be-e2a3-47c8-85f7-75d7b7f6034f",
    datasetId: "",
    name: "User Context Aware Store Sales",
    isSecureEmbedded: false,
  },
];
