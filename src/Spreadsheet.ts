import {AuthClient} from "google-auth-library/build/src/auth/authclient";
import {Sheet} from "./Sheet";
import {GOOGLE_SPREADSHEETS_URL, ISpreadsheetProperties, IUpdateOptions, IValues} from "./types";

export class Spreadsheet {
    public spreadsheetId!: string;
    public spreadsheetUrl!: string;
    public properties!: ISpreadsheetProperties;
    public sheets!: Sheet[];

    constructor(spreadsheetId: string, client: AuthClient) {
        this.spreadsheetId = spreadsheetId;

        // hide the property from console.log
        Object.defineProperty(this, "client", {
            enumerable: false,
            value: client,
        });
    }

    // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/get
    public async init() {
        const client = this._getClient();
        const url = `/${this.spreadsheetId}/`;
        const params = {includeGridData: false, ranges: []};
        const res = await client.request({ baseURL: GOOGLE_SPREADSHEETS_URL, url, params});
        this._syncProperties(res.data);
    }

    public findSheet(name: string): Sheet | undefined {
        return this.sheets.find(x => x.title.toLowerCase() === name.toLowerCase() && !x.isDeleted);
    }

    public findSheets(names: string[]): Sheet[] {
        names = names.map(x => x.toLowerCase());
        return this.sheets.filter(x => names.includes(x.title.toLowerCase()) && !x.isDeleted);
    }

    public async createSheet(name: string): Promise<Sheet> {
        const client = this._getClient();
        const url = `/${this.spreadsheetId}:batchUpdate`;
        const params = {};
        const body = {
            requests: [
                {addSheet: {properties: {title: name, gridProperties: {rowCount: 10, columnCount: 10}}}},
            ],
            includeSpreadsheetInResponse: true,
        };
        const res = await client.request({ baseURL: GOOGLE_SPREADSHEETS_URL, url, params, data: body, method: "POST"});
        this._syncProperties((res.data as any).updatedSpreadsheet);
        return this.findSheet(name) as Sheet;
    }

    public async findOrCreateSheet(name: string): Promise<Sheet> {
        let sheet = await this.findSheet(name);
        if (!sheet) {
            sheet =  await this.createSheet(name);
        }

        return sheet;
    }

    public async updateSheet(name: string, values: IValues, options: Partial<IUpdateOptions> = {}) {
        const sheet = await this.findOrCreateSheet(name);
        return await sheet.update(values, options);
    }

    private _getClient(): AuthClient {
        return (this as any).client as any;
    }

    private _syncProperties(data: any) {
        const client = this._getClient();
        this.spreadsheetUrl = data.spreadsheetUrl;
        this.properties = data.properties;
        this.sheets = (data.sheets as any[]).map(x => new Sheet(client, this.spreadsheetId, x.properties));
    }
}
