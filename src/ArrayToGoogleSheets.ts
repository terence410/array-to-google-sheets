import {GoogleAuth, GoogleAuthOptions, JWT} from "google-auth-library";
import {Spreadsheet} from "./Spreadsheet";

const authSymbol = Symbol("auth");
const clientSymbol = Symbol("client");

export class ArrayToGoogleSheets {
    private [authSymbol]: GoogleAuth;
    private [clientSymbol]: JWT;

    constructor(options: GoogleAuthOptions) {
        options.scopes = "https://www.googleapis.com/auth/spreadsheets";
        this[authSymbol] = new GoogleAuth(options);
    }

    public async getSpreadsheet(spreadsheetId: string): Promise<Spreadsheet> {
        const client = await this._getClient();
        const spreadsheet = new Spreadsheet(spreadsheetId, client);
        await spreadsheet.init();
        return spreadsheet;
    }

    private async _getClient() {
        if (!this[clientSymbol]) {
            this[clientSymbol] = await this[authSymbol].getClient() as JWT;
        }

        return this[clientSymbol];
    }
}
