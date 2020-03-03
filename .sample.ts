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

    type IObject = {value1: string; value2: string; value3: number; value4: boolean; value5: Date; value6: number[]; value7: string[]};
    const objectSheet = await sheet.exportAsObjectSheet<IObject>();
    const type = objectSheet.getType();

    // this can print the type IObject
    console.log("typescript", type);

    // get data as object
    for (let i = 0; i < objectSheet.length; i++) {
        const item = objectSheet.get(i);
        item.value1 = "key" + i;
        item.value2 = "value" + i;
        item.value3 = Math.random();
        item.value4 = true;
        item.value5 = new Date();
        item.value6 = [i, 1, 2, Math.random()];
        item.value7 = [i.toString(), "a", "b", "c"];
        await item.save();
    }
}
