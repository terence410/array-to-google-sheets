import {ArrayToGoogleSheets, IUpdateOptions} from "./src/";

async function simple() {
    const googleSheets = new ArrayToGoogleSheets({keyFilename: "serviceAccount.json"});
    const spreadsheet = await googleSheets.getSpreadsheet("19Ef3falKKiHAOo3Ps13tVC9M0BaG-NoYngAZggt8Jzk-NodnEAz5gt3iak");
    await spreadsheet.updateSheet("sheetName", [[1, 2, 3]]);
}

async function advance() {
    // https://www.npmjs.com/package/google-auth-library

    // service account json (either one)
    const keyFilename = "serviceAccount.json";
    // credentials for service account (either one)
    const credentials = {client_email: "", private_key: ""};

    const spreadsheetId = "";
    const googleSheets = new ArrayToGoogleSheets({keyFilename, credentials});
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
