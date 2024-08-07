import fs from "fs";
import {AuthClient} from "google-auth-library/build/src/auth/authclient";
import {ObjectSheet} from "./ObjectSheet";
import {
    GOOGLE_SPREADSHEETS_URL,
    ICell,
    IFormulaCells,
    INormalizedCell,
    INormalizedRow,
    INormalizedValues,
    IRow,
    ISheetProperties,
    IUpdateBaseOptions,
    IUpdateCells,
    IUpdateOptions,
    IUpdateResponse,
    IValues,
} from "./types";

export class Sheet {
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

    constructor(client: AuthClient, spreadsheetId: string, properties: ISheetProperties) {
        this.spreadsheetId = spreadsheetId;
        Object.assign(this, properties);

        // hide the property from console.log
        Object.defineProperty(this, "client", {
            enumerable: false,
            value: client,
        });
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
                fs.writeFile(filename, csv, (err) => err ? reject(err) : resolve(0)));
        }

        return csv;
    }

    public async exportAsObjectSheet<T extends object = object>(): Promise<ObjectSheet<T>> {
        const values = await this.getValues();
        return new ObjectSheet<T>(this, values);
    }

    // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/get
    public async getValues(): Promise<INormalizedValues> {
        const client = this._getClient();
        const range = this._getRange();
        const url = `/${this.spreadsheetId}/values/${range}`;
        const params = {
            valueRenderOption: "UNFORMATTED_VALUE",
        };
        const res = await client.request({ baseURL: GOOGLE_SPREADSHEETS_URL, url, params});
        const values = (res.data as any).values;
        return (values || []) as INormalizedValues;
    }

    // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/clear
    /** @internal */
    public async clear() {
        const client = this._getClient();
        const range = this._getRange();
        const url = `/${this.spreadsheetId}/values/${range}:clear`;
        const params = {};
        const res = await client.request({ baseURL: GOOGLE_SPREADSHEETS_URL, url, params, method: "POST"});
    }

    // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/sheets#SheetProperties
    /** @internal */
    public async resize(rowCount: number, columnCount: number) {
        // we update it first to avoid async problem
        this.gridProperties.rowCount = rowCount;
        this.gridProperties.columnCount = columnCount;

        const client = this._getClient();
        const url = `/${this.spreadsheetId}:batchUpdate`;
        const params = {};
        const body = {
            requests: [
                {
                    updateSheetProperties: {
                        fields: "*",
                        properties: {index: this.index, sheetId: this.sheetId, title: this.title, gridProperties: this.gridProperties},
                    },
                },
            ],
            includeSpreadsheetInResponse: false,
        };
        const res = await client.request({baseURL: GOOGLE_SPREADSHEETS_URL, url, params, data: body, method: "POST"});
    }

    // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/batchUpdate
    public async delete() {
        // mark as deleted
        this.isDeleted = true;

        const client = this._getClient();
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
            includeSpreadsheetInResponse: false,
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

        // sync grid size first
        await this._syncGridSize(finalTotalRows, finalTotalColumns, options.fitToSize);

        // clear the entire sheet first
        if (options.clearAllValues) {
            await this.clear();
        }

        const batch = 1000;
        const client = this._getClient();
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

    public async updateRow(rowIndex: number, row: IRow, options: IUpdateBaseOptions = {}): Promise<IUpdateResponse> {
        const client = this._getClient();
        const range = this._getRange(row.length, 1, 0, rowIndex);
        const params = {
            valueInputOption: options.valueInputOption || "USER_ENTERED",
        };
        const url = `/${this.spreadsheetId}/values/${range}`;
        const normalizedRow = this._normalizeRow(rowIndex, row);
        const body = {values: [normalizedRow]};
        const res = await client.request({
            baseURL: GOOGLE_SPREADSHEETS_URL,
            url,
            params,
            data: body,
            method: "PUT",
        });

        const updateResponse: IUpdateResponse = {updatedCells: 0, updatedColumns: 0, updatedRows: 0};
        return this._addUpdateResponse(updateResponse, res.data as IUpdateResponse);
    }

    public async updateCell(rowIndex: number, columnIndex: number, cell: ICell, options: IUpdateBaseOptions = {}): Promise<IUpdateResponse> {
        const client = this._getClient();
        const range = this._getRange(1, 1, columnIndex, rowIndex);
        const params = {
            valueInputOption: options.valueInputOption || "USER_ENTERED",
        };
        const url = `/${this.spreadsheetId}/values/${range}`;
        const normalizedCell = this._normalizeCell(rowIndex, columnIndex, cell);
        const body = {values: [[normalizedCell]]};
        const res = await client.request({
            baseURL: GOOGLE_SPREADSHEETS_URL,
            url,
            params,
            data: body,
            method: "PUT",
        });

        const updateResponse: IUpdateResponse = {updatedCells: 0, updatedColumns: 0, updatedRows: 0};
        return this._addUpdateResponse(updateResponse, res.data as IUpdateResponse);
    }

    public async updateCells(cells: IUpdateCells, options: IUpdateBaseOptions = {}): Promise<IUpdateResponse> {
        const client = this._getClient();
        const url = `/${this.spreadsheetId}/values:batchUpdate`;
        const params = {};
        const body = {
            data: cells.map(x => {
                return {values: [[this._normalizeCell(x.rowIndex, x.columnIndex, x.cell)]], range: this._getRange(1, 1, x.columnIndex, x.rowIndex)};
            }),
            valueInputOption: options.valueInputOption || "USER_ENTERED",
        };

        const res = await client.request({
            baseURL: GOOGLE_SPREADSHEETS_URL,
            url,
            params,
            data: body,
            method: "POST",
        });

        const updateResponse: IUpdateResponse = {updatedCells: 0, updatedColumns: 0, updatedRows: 0};
        return this._addUpdateResponse(updateResponse, res.data as IUpdateResponse);
    }

    // region private methods

    private _getClient(): AuthClient {
        return (this as any).client as AuthClient;
    }

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

        if (!width && !height) {
            return this.title;
        }

        const rangeA = `${this._convertColToAlphabet(offsetX + 1)}${offsetY + 1}`;
        const rangeB = `${this._convertColToAlphabet(offsetX + width)}${offsetY + height}`;
        return `${this.title}!${rangeA}:${rangeB}`;
    }

    private _normalizeValues(values: IValues): INormalizedValues {
        return values.map((row, i) => {
            return row.map((cell, j) => {
                if (typeof cell === "object") {
                    return this._formatFormula(cell.formula, cell.cells, i + 1, j + 1);
                }

                return cell;
            });
        });
    }

    private _normalizeRow(rowIndex: number, row: IRow): INormalizedRow {
        return row.map((cell, j) => {
            if (typeof cell === "object") {
                return this._formatFormula(cell.formula, cell.cells, rowIndex + 1, j + 1);
            }

            return cell;
        });
    }

    private _normalizeCell(rowIndex: number, columnIndex: number, cell: ICell): INormalizedCell {
        if (typeof cell === "object") {
            return this._formatFormula(cell.formula, cell.cells, rowIndex + 1, columnIndex + 1);
        }

        return cell;
    }

    private _formatFormula(formulaFormat: string, formulaCells: IFormulaCells | undefined, currentRow: number, currentCol: number) {
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

    private async _syncGridSize(finalTotalRows: number, finalTotalColumns: number, forceUpdate: boolean) {
        if (forceUpdate) {
            await this.resize(finalTotalRows, finalTotalColumns);

        } else if (finalTotalRows > this.gridProperties.rowCount && finalTotalColumns > this.gridProperties.columnCount) {
            await this.resize(finalTotalRows, finalTotalColumns);

        } else if (finalTotalRows > this.gridProperties.rowCount) {
            await this.resize(finalTotalRows, this.gridProperties.columnCount);

        } else if (finalTotalColumns > this.gridProperties.columnCount) {
            await this.resize(this.gridProperties.rowCount, finalTotalColumns);

        }
    }

    private _addUpdateResponse(updateResponse: IUpdateResponse, newUpdateResponse: IUpdateResponse) {
        updateResponse.updatedRows += newUpdateResponse.updatedRows;
        updateResponse.updatedColumns = Math.max(updateResponse.updatedColumns, newUpdateResponse.updatedColumns);
        updateResponse.updatedCells += newUpdateResponse.updatedCells;
        return updateResponse;
    }

    // endregion
}
