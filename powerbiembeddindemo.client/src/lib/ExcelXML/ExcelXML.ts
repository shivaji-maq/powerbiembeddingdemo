/* eslint-disable @typescript-eslint/no-explicit-any */
export class ExcelXML {
    private workbookStart: string;
    private workbookEnd: string = "</ss:Workbook>";
    private sheetName: string = "SHEET 1";
    private styleID: number = 1;
    private columnWidth: number = 80;
    private fileName: string = "Employee_List";
    private jsonData: Record<string, any>[] = [];

    constructor(data: string | object[]) {
        const respArray = Array.isArray(data) ? data : JSON.parse(data);
        this.jsonData = respArray.map((item: any) => ExcelXML.flatten(item));

        this.workbookStart =
            `<?xml version="1.0"?>` +
            `<ss:Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" ` +
            `xmlns:o="urn:schemas-microsoft-com:office:office" ` +
            `xmlns:x="urn:schemas-microsoft-com:office:excel" ` +
            `xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" ` +
            `xmlns:html="http://www.w3.org/TR/REC-html40">`;
    }

    public download(): void {
        const worksheet = this.buildWorksheet(this.sheetName, this.jsonData);
        const styles = this.buildStyles(this.styleID);

        const workbook = this.workbookStart + styles + worksheet + this.workbookEnd;

        const uri = "data:text/xls;charset=utf-8," + encodeURIComponent(workbook);
        const link = document.createElement("a");
        link.href = uri;
        link.style.visibility = "hidden";
        link.download = `${this.fileName}.xls`;

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    // -------------------- Getters/Setters --------------------

    public get FileName(): string {
        return this.fileName;
    }
    public set FileName(n: string) {
        this.fileName = n;
    }

    public get SheetName(): string {
        return this.sheetName;
    }
    public set SheetName(n: string) {
        this.sheetName = n;
    }

    public get StyleID(): number {
        return this.styleID;
    }
    public set StyleID(n: number) {
        this.styleID = n;
    }

    // -------------------- Private Builders --------------------

    private buildStyles(id: number): string {
        return `<ss:Styles><ss:Style ss:ID="${id}"><ss:Font ss:Bold="1"/></ss:Style></ss:Styles>`;
    }

    private buildWorksheet(name: string, data: Record<string, any>[]): string {
        const table = this.buildTable(data);
        return `<ss:Worksheet ss:Name="${name}">${table}</ss:Worksheet>`;
    }

    private buildTable(data: Record<string, any>[]): string {
        let table = "<ss:Table>";

        if (data.length > 0) {
            const columnHeader = Object.keys(data[0]);

            // Columns
            for (let i = 0; i < columnHeader.length; i++) {
                table += this.buildColumn(this.columnWidth);
            }

            // Header
            table += this.buildHead(this.styleID, columnHeader);

            // Rows
            let rowData = "";
            for (let j = 0; j < data.length; j++) {
                rowData += this.buildRow(data[j], columnHeader);
            }
            table += rowData;
        }

        table += "</ss:Table>";
        return table;
    }

    private buildColumn(width: number): string {
        return `<ss:Column ss:AutoFitWidth="0" ss:Width="${width}"/>`;
    }

    private buildHead(id: number, headers: string[]): string {
        let head = `<ss:Row ss:StyleID="${id}">`;
        for (const h of headers) {
            head += this.buildCell(h.toUpperCase());
        }
        head += "</ss:Row>";
        return head;
    }

    private buildRow(row: Record<string, any>, headers: string[]): string {
        let rowXML = "<ss:Row>";
        for (const h of headers) {
            rowXML += this.buildCell(row[h] ?? "");
        }
        rowXML += "</ss:Row>";
        return rowXML;
    }

    private buildCell(value: string | number | boolean): string {
        return `<ss:Cell>${this.buildData(value)}</ss:Cell>`;
    }

    private buildData(value: string | number | boolean): string {
        const safeValue = String(value).replace(/&/g, "&amp;");
        return `<ss:Data ss:Type="String">${safeValue}</ss:Data>`;
    }

    // -------------------- Utility --------------------

    private static flatten(obj: Record<string, any>): Record<string, any> {
        const result: Record<string, any> = {};

        function recurse(curr: Record<string, any>, parentKey = "") {
            for (const key in curr) {
                if (!Object.prototype.hasOwnProperty.call(curr, key)) continue;
                const newKey = parentKey ? `${parentKey}-${key}` : key;

                if (typeof curr[key] === "object" && curr[key] !== null) {
                    recurse(curr[key], newKey);
                } else {
                    result[newKey] = curr[key];
                }
            }
        }

        recurse(obj);
        return result;
    }
}
