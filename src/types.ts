declare module "google-spreadsheet" {
    export default class GoogleSpreadsheet {
        constructor(dockey: string);

        public updateGoogleSheets(sheetName: string, array2D: any[][]): Promise<{url: string, gid: number}>;
        public useServiceAccountAuth(options: any, callback: (err: Error) => void): void;
        public getInfo(callback: (err: Error, docInfo: DocInfo) => void): void;
        public addWorksheet(options: any, callback: (err: Error, sheet: WorkSheet) => void): void;
    }

    export type DocInfo = {
        id: string,
        title: string,
        updated: string,
        author: {name: string, email: string},
        worksheets: WorkSheet[],
    };

    export type WorkSheet = {
        url: string,
        id: string,
        title: string,
        rowCount: number,
        colCount: number,
        resize: (options: any, callback: (err?: Error) => void) => void;
        setTitle: (options: any, callback: (err?: Error) => void) => void;
        clear: (options: any, callback: (err?: Error) => void) => void;
        getRows: (options: any, callback: (err?: Error) => void) => void;
        getCells: (options: any, callback: (err: Error, cells: any[]) => void) => void;
        addRow: (options: any, callback: (err?: Error) => void) => void;
        bulkUpdateCells: (data: any[], callback: (err: Error) => void) => void;
        del: (options: any, callback: (err?: Error) => void) => void;
        setHeaderRow: (options: any, callback: (err?: Error) => void) => void;
    };
}
