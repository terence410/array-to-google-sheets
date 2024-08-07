import {Credentials, GoogleAuth, GoogleAuthOptions, OAuth2Client, OAuth2ClientOptions} from "google-auth-library";
import {AuthClient} from "google-auth-library/build/src/auth/authclient";
import {Spreadsheet} from "./Spreadsheet";

const scope = "https://www.googleapis.com/auth/spreadsheets";

export class ArrayToGoogleSheets {
    private _client?: AuthClient; // hide from the object

    constructor(public options: GoogleAuthOptions | {oAuthCredentials: Credentials, oauthClientOptions?: OAuth2ClientOptions}) {
        Object.defineProperty(this, "_client", {
            enumerable: false,
            writable: true,
            value: undefined,
        });
    }

    public async getSpreadsheet(spreadsheetId: string): Promise<Spreadsheet> {
        const client = await this._getClient();
        const spreadsheet = new Spreadsheet(spreadsheetId, client);
        await spreadsheet.init();
        return spreadsheet;
    }

    public static async getClient(options: GoogleAuthOptions | {
        oAuthCredentials: Credentials,
        oauthClientOptions?: OAuth2ClientOptions
    }): Promise<AuthClient> {
        if ("oAuthCredentials" in options) {
            const oauth = new OAuth2Client(options.oauthClientOptions);
            oauth.setCredentials(options.oAuthCredentials);
            return oauth;
        } else {
            const googleAuth = new GoogleAuth({...options, scopes: [scope]});
            return await googleAuth.getClient();
        }
    }

    // caching for client
    private async _getClient(): Promise<AuthClient> {
        if (!this._client) {
            if ("oAuthCredentials" in this.options) {
                const oauth = new OAuth2Client(this.options.oauthClientOptions);
                oauth.setCredentials(this.options.oAuthCredentials);
                this._client = oauth;

            } else {
                const googleAuth = new GoogleAuth({...this.options, scopes: [scope]});
                this._client = await googleAuth.getClient();
            }
        }

        return this._client!;
    }
}
