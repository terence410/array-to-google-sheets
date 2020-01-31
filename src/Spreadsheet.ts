import {GoogleAuth, JWT} from "google-auth-library";
import {Sheet} from "./Sheet";
import {IUpdateOptions, IValues} from "./types";

const GOOGLE_SPREADSHEETS_URL = "https://sheets.googleapis.com/v4/spreadsheets";
const clientSymbol = Symbol("client");

type ISpreadsheetProperties = {
    title: string
    locale: string,
    autoRecalc: string,
    timeZone: string,
    defaultFormat: {
        background: {red: number, green: number, blue: number},
        padding: {top: number, right: number, bottom: number, left: number},
        verticalAlignment: string,
        wrapStrategy: string,
        textFormat: {
            foregroundColor: any,
            fontFamily: string,
            fontSize: number,
            bold: boolean,
            italic: boolean,
            strikethrough: boolean,
            underline: boolean,
        },
    },
    spreadsheetTheme: {
        primaryFontFamily: string,
        themeColors: Array<{
            colorType: string,
            color: {
                rgbColor: {
                    red: number,
                    green: number,
                    blue: number,
                },
            },
        }>,
    },
};

export class Spreadsheet {
    public spreadsheetId!: string;
    public spreadsheetUrl!: string;
    public properties!: ISpreadsheetProperties;
    public sheets!: Sheet[];
    private [clientSymbol]: JWT;

    constructor(spreadsheetId: string, jwt: JWT) {
        this.spreadsheetId = spreadsheetId;
        this[clientSymbol] = jwt;
    }

    // https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/get
    public async init() {
        const client = this[clientSymbol];
        const url = `/${this.spreadsheetId}/`;
        const params = {includeGridData: false, ranges: []};
        const res = await client.request({ baseURL: GOOGLE_SPREADSHEETS_URL, url, params});
        this._syncProperties(res.data);
    }

    public findSheet(name: string): Sheet | undefined {
        return this.sheets.find(x => x.title.toLowerCase() === name.toLowerCase() && !x.isDeleted);
    }

    public async createSheet(name: string): Promise<Sheet> {
        const client = this[clientSymbol];
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

    private _syncProperties(data: any) {
        this.spreadsheetUrl = data.spreadsheetUrl;
        this.properties = data.properties;
        this.sheets = (data.sheets as any[]).map(x => new Sheet(this[clientSymbol], this.spreadsheetId, x.properties));
    }
}
