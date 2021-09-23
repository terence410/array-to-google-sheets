import {ArrayToGoogleSheets, IUpdateOptions} from "./src/";

async function simple() {
    const googleSheets = new ArrayToGoogleSheets({keyFilename: "serviceAccount.json"});
    // https://docs.google.com/spreadsheets/d/[ID]
    const spreadsheet = await googleSheets.getSpreadsheet("ID");
    await spreadsheet.updateSheet("sheetName", [[1, 2, 3]]);
}

async function auth() {
    // https://www.npmjs.com/package/google-auth-library
    const googleSheets1 = new ArrayToGoogleSheets({keyFilename: "serviceAccount.json"});
    const googleSheets2 = new ArrayToGoogleSheets({credentials: {client_email: "", private_key: ""}});

    // oauth
    const googleSheets3 = new ArrayToGoogleSheets({oAuthCredentials: {access_token: ""}});
    const googleSheets4 = new ArrayToGoogleSheets({oAuthCredentials: {refresh_token: ""}, oauthClientOptions: {clientId: "", clientSecret: ""}});
}

async function advance() {
    const spreadsheetId = "";
    const googleSheets = new ArrayToGoogleSheets({keyFilename: "serviceAccount.json"});
    const spreadsheet = await googleSheets.getSpreadsheet(spreadsheetId);
    const {spreadsheetUrl, properties} = spreadsheet;
    const {title, locale, timeZone, defaultFormat} = properties;

    // find and delete
    const sheetName = "sheetName";
    const sheet = await spreadsheet.findSheet(sheetName);
    if (sheet) {
        const result = await sheet.delete();
    }

    // get sheet again
    const newSheet = await spreadsheet.findOrCreateSheet(sheetName);
    const url = newSheet.getUrl();

    //  update
    const values1 = [
        [1, 2, 3],
        [1.1, 2.2, -3.33],
        ["abc", "cde", "xyz"],
    ];
    const updateOptions: IUpdateOptions = {
        minRow: 3, // styling
        minColumn: 3, // styling
        margin: 2,  // styling
        fitToSize: true,  // remove empty cells
        clearAllValues: true, // clear all existing values
    };
    const updateResult1 = await newSheet.update(values1, updateOptions);
    const resultValues1 = await newSheet.getValues();

    // export into csv
    await newSheet.exportAsCsv("data.csv");
}

async function updateRowsAndCells() {
    // expand the sheet size first if u have many rows
    const googleSheets = new ArrayToGoogleSheets({keyFilename: "serviceAccount.json"});
    const spreadsheet = await googleSheets.getSpreadsheet("spreadsheetId");
    const sheet = await spreadsheet.findOrCreateSheet("sheetName");

    // we have to make sure we have enough grids
    await sheet.resize(10, 10);

    for (let i = 0; i < 10; i++) {
        await sheet.updateRow(i, [1, 2, 3]);
        await sheet.updateCell(i, i, 1);
    }
}

async function experimentalObjectSheet() {
    const googleSheets = new ArrayToGoogleSheets({keyFilename: "serviceAccount.json"});
    const spreadsheet = await googleSheets.getSpreadsheet("spreadsheetId");
    const sheet = await spreadsheet.findOrCreateSheet("sheetName");

    const values = [
        ["value1", "value2/string", "value3/number", "value4/boolean", "value5/date", "value6/number[]", "value7/string[]", "value8/ignore"],
        ["1", "2", "3", "4", "5", "6", "7", "8", "9"],
        [1, 2, 3, 4, 5, 6, 7, 8, 9],
        ["a", "b", "c", "d", "e", "f", "g", "h", "i"],
    ];
    await sheet.update(values, {clearAllValues: true, margin: 2});

    interface IObject {
        value1: string;
        value2: string;
        value3: number;
        value4: boolean;
        value5: Date;
        value6: number[];
        value7: string[];
    }

    const objectSheet = await sheet.exportAsObjectSheet<IObject>();
    const objectInterface = objectSheet.getInterface();
    const {headers, size, rawValues, rawHeaders} = objectSheet;

    // this can print the interface IObject
    console.log("typescript", objectInterface);

    // get data as object
    const firstItem = objectSheet.get(0);
    for (const item of objectSheet) {
        item.value1 = "key";
        item.value2 = "value";
        item.value3 = Math.random();
        item.value4 = true;
        item.value5 = new Date();
        item.value6 = [1, 2, Math.random()];
        item.value7 = ["a", "b", "c"];
        await item.save();
    }

    const findItem = objectSheet.toArray().find(x => x.value1 === "key");
    const objects = objectSheet.toObjects();

    // add new item
    const newItem = await objectSheet.append({} as any);
    // you have to manage the sheet size yourself
    await sheet.resize(100, 100);
}
