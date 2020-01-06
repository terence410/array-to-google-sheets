import fs from "fs";
import "./types";
// tslint:disable-next-line
import GoogleSpreadsheet, {DocInfo, WorkSheet} from "google-spreadsheet";

export type IOptions = {
    minRow: number,
    minCol: number,
    margin: number,
    resize: boolean,
    clear: boolean,
};

type IFormulaCell = {
    row: number | "this",
    col: number | "this",
};

type IWorkSheetCell = {
    row: number,
    col: number,
    value: any,
};

export class ArrayToGoogleSheets {
    public doc: GoogleSpreadsheet;
    public docInfo: DocInfo | undefined;
    public credentialJson: object;

    constructor(public readonly docKey: string, credential: any) {
        if (typeof credential === "string") {
            // this is file, convert to json
            this.credentialJson = this.importJson(credential);
        } else {
            this.credentialJson = credential;
        }

        this.doc = new GoogleSpreadsheet(this.docKey);
    }

    // region public methods: utils

    public convertColToAlphabet(num: number): string {
        const char = (10 + (num - 1) % 26).toString(36).toUpperCase();
        return num > 0 ? this.convertColToAlphabet((num - 1) / 26 | 0) + char : "";
    }

    public convertWid2gid(wid: string): number {
        const widVal = wid.length > 3 ? wid.substr(1) : wid;
        const xorVal = wid.length > 3 ? 474 : 31578;
        return parseInt(String(widVal), 36) ^ xorVal;
    }

    public importJson(file: string) {
        const data = fs.readFileSync(file).toString();
        return JSON.parse(data);
    }

    public getUrlObject(workSheet: WorkSheet) {
        const gid = this.convertWid2gid(workSheet.id);
        const url = `https://docs.google.com/spreadsheets/d/${this.docKey}/edit#gid=${gid}`;
        return {url, gid};
    }

    public formatFormula(formulaFormat: string, cells: IFormulaCell[], currentRow: number, currentCol: number) {
        if (!Array.isArray(cells)) {
            return formulaFormat;
        }

        const arr = cells.map(cell => {
            const col = cell.col === "this" ? currentCol : cell.col;
            const row = cell.row === "this" ? currentRow : cell.row;
            return (col > 0 ? this.convertColToAlphabet(col) : "") + (row > 0 ? row : "");
        });

        return formulaFormat.replace(/%(\d+)/g, (_, m) => {
            return arr[--m];
        });
    }

    // endregion

    // region public methods

    public async getDocInfo(): Promise<DocInfo> {
        if (!this.docInfo) {
            // connect to the service account
            await new Promise((resolve, reject) => {
                this.doc.useServiceAccountAuth(this.credentialJson, (err) => {
                    if (err) {
                        return reject(err);
                    }
                    resolve();
                });
            });

            // open the document
            this.docInfo = await new Promise((resolve, reject) => {
                this.doc.getInfo((err, docInfo) => {
                    if (err) {
                        return reject(err);
                    }
                    resolve(docInfo);
                });
            });
        }

        return this.docInfo as DocInfo;
    }

    public async updateWorkSheet(sheetName: string, values: any[], partialOptions: Partial<IOptions> = {}) {
        const options: IOptions = Object.assign({
            minCol: 10,
            minRow: 20,
            margin: 2,
            resize: true,
            clear: true,
        }, partialOptions);

        // the values array must be 2 dimensional
        if (values.length > 0 && !Array.isArray(values[0])) {
            throw new Error("Values must be a 2 dimensional array.");
        }

        // get docInfo
        const docInfo = await this.getDocInfo();

        // search for existing worksheet
        let workSheet = docInfo.worksheets.find(x => x.title.toLowerCase() === sheetName.toLowerCase());

        // create a new work sheet if not exist
        if (!workSheet) {
            workSheet = await this._addWorkSheet(sheetName);
        }

        // make sure sheet exist
        if (!workSheet) {
            throw new Error(`Cannot create new sheet: ${sheetName}`);
        }

        // update array to the sheet
        await this._updateWorkSheetInternal(workSheet, values, options);
        return this.getUrlObject(workSheet);
    }

    public async getWorkSheetDataAsArray(sheetName: string): Promise<any[] | undefined>;
    public async getWorkSheetDataAsArray(sheetNames: string[]): Promise<any[]>;
    public async getWorkSheetDataAsArray(sheetName: string | string[]) {
        const sheetNames = Array.isArray(sheetName) ? sheetName : [sheetName];

        // get doc
        const docInfo = await this.getDocInfo();

        const promises: any[] = [];
        for (const currentSheetName of sheetNames) {
            const workSheet = docInfo.worksheets.find(x => currentSheetName === x.title);
            if (workSheet) {
                promises.push(this._getCells(workSheet, workSheet.rowCount));
            } else {
                promises.push(Promise.resolve(undefined));
            }
        }
        const promiseResults = await Promise.all(promises);

        // prepare result
        const result: any[] = [];
        for (const cells of promiseResults) {
            if (Array.isArray(cells)) {
                // turn into array2D
                let array2d: any[] = [];
                cells.forEach(cell => {
                    if (!array2d[cell.row - 1]) {
                        array2d[cell.row - 1] = [];
                    }
                    array2d[cell.row - 1].push(cell.value);
                });

                // filter empty rows
                array2d = array2d.filter(row => {
                    return row.filter((cell: any[]) => cell.length > 0).length > 0;
                });

                result.push(array2d);
            } else {
                result.push(undefined);
            }
        }

        if (Array.isArray(sheetName)) {
            return result;
        } else {
            return result[0];
        }
    }

    public async getWorkSheetDataAsCsv(sheetName: string): Promise<string | undefined>;
    public async getWorkSheetDataAsCsv(sheetNames: string[]): Promise<string[]>;
    public async getWorkSheetDataAsCsv(sheetName: string | string[]) {
        const sheetNames = Array.isArray(sheetName) ? sheetName : [sheetName];
        const array2dList = await this.getWorkSheetDataAsArray(sheetNames);

        // escape strings
        const result: Array<string | undefined> = [];
        for (const array2d of array2dList) {

            if (Array.isArray(array2d)) {
                const escapedArray2d = array2d.map(row => {
                    return row.map((cell: string) => {
                        // we use double quote to wrap the strings if it has [,"\r\n]
                        if (cell.match(/[,"\r\n]/)) {
                            return `"${cell.replace(/"/g, "\"\"")}"`;
                        }
                        return cell;
                    });
                });

                // format into csv
                const csv = escapedArray2d.map(row => row.join(",")).join("\r\n");
                result.push(csv);
            } else {
                result.push(undefined);
            }
        }

        if (Array.isArray(sheetName)) {
            return result;
        } else {
            return result[0];
        }
    }

    // endregion

    // region private methods

    private async _updateWorkSheetInternal(workSheet: WorkSheet, values: any[][], options: IOptions) {
        let rowCount = Math.max(options.minRow, values.length) + options.margin;
        let colCount = Math.max(options.minCol, values.reduce((a, b) => Math.max(a, b.length), 0)) + options.margin;

        // apply options
        if (!options.resize) {
            rowCount = workSheet.rowCount;
            colCount = workSheet.colCount;
        }

        await this._resizeWorkSheet(workSheet, rowCount, colCount);
        const cells = await this._getCells(workSheet, rowCount);

        // store the cells in array[row][column] format for ease of access
        const myCells: any[] = [];
        const updatedCells: any[] = [];

        // process the cells
        cells.forEach(cell => {
            if (!myCells[cell.row - 1]) {
                myCells[cell.row - 1] = [];
            }
            myCells[cell.row - 1].push(cell);

            // clear all the cells's value
            if (options.clear) {
                cell.value = "";
            }
        });

        // update the cell
        values.forEach((list, i) => {
            list.forEach((value, j) => {
                if (myCells[i] && myCells[i][j]) {
                    if (!isNaN(parseFloat(value)) && isFinite(value)) {
                        myCells[i][j].numericValue = value;
                    } else if (typeof value === "object") {
                        myCells[i][j].formula = this.formatFormula(value.formula, value.cells, i + 1, j + 1);
                    } else {
                        myCells[i][j].value = value;
                    }
                    updatedCells.push(myCells[i][j]);
                }
            });
        });

        // update the cells
        await this._updateCells(workSheet, options.clear ? cells : updatedCells);
    }

    private async _addWorkSheet(sheetName: string): Promise<WorkSheet> {
        return await new Promise((resolve, reject) => {
            this.doc.addWorksheet({title: sheetName}, (err, newSheet2) => {
                if (err) {
                    return reject(err);
                }
                resolve(newSheet2);
            });
        });
    }

    private async _resizeWorkSheet(workSheet: WorkSheet, rowCount: number, colCount: number) {
        return new Promise((resolve, reject) => {
            workSheet.resize({rowCount, colCount}, (err) => {
                if (err) {
                    return reject(err);
                }

                resolve();
            });
        });
    }

    private async _getCells(workSheet: WorkSheet, maxRow: number): Promise<IWorkSheetCell[]> {
        return new Promise((resolve, reject) => {
            workSheet.getCells({
                "min-row": 1,
                "max-row": maxRow,
                "return-empty": true,
            }, (err, cells) => {
                if (err) {
                    return reject(err);
                }
                resolve(cells);
            });
        });
    }

    private async _updateCells(workSheet: WorkSheet, cells: any[][]) {
        await new Promise((resolve, reject) => {
            workSheet.bulkUpdateCells(cells, (err) => {
                if (err) {
                    return reject(err);
                }
                resolve();
            });
        });
    }

    // endregion
}
