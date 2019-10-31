declare module "google-spreadsheet" {
    export default class GoogleSpreadsheet {
        constructor(dockey: string);

        public updateGoogleSheets(sheetName: string, array2D: any[][]): Promise<{url: string, gid: number}>;
        public useServiceAccountAuth(options: any, callback: (err: Error) => void): void;
        public getInfo(callback: (err: Error, docInfo: DocInfo) => void): void;
        public addWorksheet(options: any, callback: (err: Error, sheet: SpreadsheetWorksheet) => void): void;
    }

    export type DocInfo = {
        worksheets: SpreadsheetWorksheet[],
    };

    export type SpreadsheetWorksheet = {
        id: string,
        title: string,
        rowCount: number,
        colCount: number,
        resize: (options: any, callback: (err?: Error) => void) => void;
        getCells: (options: any, callback: (err: Error, cells: any[]) => void) => void;
        bulkUpdateCells: (data: any[], callback: (err: Error) => void) => void;
    };
}
