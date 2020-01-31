import fs from "fs";
import {GoogleAuth, JWT} from "google-auth-library";
import {IFormulaCells, IPlainValues, IUpdateOptions, IUpdateResponse, IValues} from "./types";

const GOOGLE_SPREADSHEETS_URL = "https://sheets.googleapis.com/v4/spreadsheets";
const clientSymbol = Symbol("auth");

export class Sheet {
    public [clientSymbol]: JWT;
    public isDeleted = false;
    public spreadsheetId!: string;
    public sheetId!: number;
    public title!: string;
    public index!: 0;
    public sheetType!: string;
    public gridProperties!: {
        rowCount: number;
        columnCount: number;
    };

    constructor(jwt: JWT, spreadsheetId: string, properties: object) {
        this.spreadsheetId = spreadsheetId;
        Object.assign(this, properties);
        this[clientSymbol] = jwt;
    }

    /** @internal */
    public getUrl() {
        return `https://docs.google.com/spreadsheets/d/${this.spreadsheetId}/edit#gid=${this.sheetId}`;
    }

    public async exportAsCsv(filename: string = ""): Promise<string> {
        const values = await this.getValues();

        const escapedValues = values.map(row => {
            return row.map((cell: string | number) => {
                // we use double quote to wrap the strings if it has [,"\r\n]
                if (cell.toString().match(/[,"\r\n]/)) {
                    return `"${cell.toString().replace(/"/g, "\"\"")}"`;
                }
                return cell;
            });
        });

        // format into csv
        const csv = escapedValues.map(row => row.join(",")).join("\r\n");

        // if we wanted to save as file
        if (filename) {
            await new Promise((resolve, reject) =>
                fs.writeFile(filename, csv, (err) => err ? reject(err) : resolve()));
        }

        return csv;
    }

    // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/get
    public async getValues(): Promise<IPlainValues> {
        const client = this[clientSymbol];
        const range = this._getRange();
        const url = `/${this.spreadsheetId}/values/${range}`;
        const params = {
            valueRenderOption: "UNFORMATTED_VALUE",
        };
        const res = await client.request({ baseURL: GOOGLE_SPREADSHEETS_URL, url, params});
        return (res.data as any).values as IPlainValues;
    }

    // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/clear
    /** @internal */
    public async clear() {
        const client = this[clientSymbol];
        const range = this._getRange();
        const url = `/${this.spreadsheetId}/values/${range}:clear`;
        const params = {};
        const res = await client.request({ baseURL: GOOGLE_SPREADSHEETS_URL, url, params, method: "POST"});
    }

    // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/sheets#SheetProperties
    /** @internal */
    public async fitToSize(rowCount: number, columnCount: number) {
        const client = this[clientSymbol];
        const url = `/${this.spreadsheetId}:batchUpdate`;
        const params = {};
        const body = {
            requests: [
                {
                    updateSheetProperties: {
                        fields: "*",
                        properties: {sheetId: this.sheetId, title: this.title, gridProperties: {rowCount, columnCount}},
                    },
                },
            ],
            includeSpreadsheetInResponse: true,
        };
        const res = await client.request({ baseURL: GOOGLE_SPREADSHEETS_URL, url, params, data: body, method: "POST"});
        this.gridProperties.rowCount = rowCount;
        this.gridProperties.columnCount = columnCount;
    }

    public async delete() {
        // mark as deleted
        this.isDeleted = true;

        const client = this[clientSymbol];
        const url = `/${this.spreadsheetId}:batchUpdate`;
        const params = {};
        const body = {
            requests: [
                {
                    deleteSheet: {
                        sheetId: this.sheetId,
                    },
                },
            ],
            includeSpreadsheetInResponse: true,
        };
        const res = await client.request({ baseURL: GOOGLE_SPREADSHEETS_URL, url, params, data: body, method: "POST"});
        return res.data;
    }

    // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/update
    public async update(values: IValues, updateOptions: Partial<IUpdateOptions> = {}): Promise<IUpdateResponse> {
        const options = Object.assign({minRow: 0, minColumn: 0, margin: 2, fitToSize: false, clearAllValues: false}, updateOptions);

        const totalRows = values.length;
        const totalColumns = Math.max(...values.map(column => column.length));
        const extraRows = Math.max(Math.max(0, options.minRow - totalRows), options.margin);
        const extraColumns = Math.max(Math.max(0, options.minColumn - totalColumns), options.margin);
        const finalTotalRows = totalRows + extraRows;
        const finalTotalColumns = totalColumns + extraColumns;

        // fit size
        if (options.fitToSize || (finalTotalRows > this.gridProperties.rowCount && finalTotalColumns > this.gridProperties.columnCount)) {
            await this.fitToSize(finalTotalRows, finalTotalColumns);

        } else if (finalTotalRows > this.gridProperties.rowCount) {
            await this.fitToSize(finalTotalRows, this.gridProperties.columnCount);

        } else if (finalTotalColumns > this.gridProperties.columnCount) {

            await this.fitToSize(this.gridProperties.rowCount, finalTotalColumns);
        }

        // update gird properties
        this.gridProperties.rowCount = Math.max(this.gridProperties.rowCount, finalTotalRows);
        this.gridProperties.columnCount = Math.max(this.gridProperties.columnCount, finalTotalColumns);

        // clear the entire sheet first
        if (options.clearAllValues) {
            await this.clear();
        }

        const batch = 1000;
        const client = this[clientSymbol];
        const params = {
            valueInputOption: "USER_ENTERED",
        };
        
        const updateResponse: IUpdateResponse = {updatedCells: 0, updatedColumns: 0, updatedRows: 0};
        for (let i = 0; i < Math.ceil(totalRows / batch); i++) {
            const offsetY = i * batch;
            const offsetX = 0;
            
            const range = this._getRange(totalColumns, batch, offsetX, offsetY);
            
            const url = `/${this.spreadsheetId}/values/${range}`;
            const normalizedValues = this._normalizeValues(values.slice(i * batch, (i + 1) * batch));
            const body = {values: normalizedValues};
            const res = await client.request({
                baseURL: GOOGLE_SPREADSHEETS_URL,
                url,
                params,
                data: body,
                method: "PUT",
            });

            const currentUpdateResponse = res.data as IUpdateResponse;
            updateResponse.updatedRows += currentUpdateResponse.updatedRows;
            updateResponse.updatedColumns = Math.max(updateResponse.updatedColumns, currentUpdateResponse.updatedColumns);
            updateResponse.updatedCells += currentUpdateResponse.updatedCells;
        }

        return updateResponse;
    }

    // region private methods

    private _convertColToAlphabet(num: number): string {
        const char = (10 + (num - 1) % 26).toString(36).toUpperCase();
        return num > 0 ? this._convertColToAlphabet((num - 1) / 26 | 0) + char : "";
    }

    private _getRange(width: number = 0, height: number = 0, offsetX: number = 0, offsetY: number = 0): string {
        if (!width) {
            width = this.gridProperties.columnCount;
        }

        if (!height) {
            height = this.gridProperties.rowCount;
        }

        const rangeA = `${this._convertColToAlphabet(offsetX + 1)}${offsetY + 1}`;
        const rangeB = `${this._convertColToAlphabet(width)}${offsetY + height + 1}`;
        return `${this.title}!${rangeA}:${rangeB}`;
    }

    private _normalizeValues(values: IValues): Array<Array<number | string>> {
        return values.map((row, i) => {
            return row.map((value, j) => {
                if (typeof value === "object") {
                    return this._formatFormula(value.formula, value.cells, i + 1, j + 1);
                } 
                
                return value;
            });
        });
    }

    private _formatFormula(formulaFormat: string, formulaCells: IFormulaCells, currentRow: number, currentCol: number) {
        if (!Array.isArray(formulaCells)) {
            return formulaFormat;
        }

        const cells = formulaCells.map(cell => {
            const col = typeof cell.col === "string" ? currentCol : cell.col;
            const row = typeof cell.row === "string" ? currentRow : cell.row;
            return (col && col > 0 ? this._convertColToAlphabet(col) : "") + (row && row > 0 ? row : "");
        });

        return formulaFormat.replace(/%(\d+)/g, (_, m) => cells[--m]);
    }

    // endregion
}