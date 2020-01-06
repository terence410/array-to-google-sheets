import {ArrayToGoogleSheets} from "./src/index";

async function main() {
    // init the library
    const docKey = "Google Sheets Key";
    const keyFilename = "./google-creds.json"; // file or json object both ok
    const sheetName = "Sheet Name";
    const a2gs = new ArrayToGoogleSheets(docKey, keyFilename);

    // save array
    const array2d = [
        [1, 2, 3.5555],
        [4, 5, 6],
        ["a", "b", "c"],
        [{formula: "=sum(%1:%2)", cells: [{row: 1, col: 1}, {row: 1, col: 3}]}], // =sum(A1:C1)
    ];

    // Options:
    // margin: Extra blank cells, for better styling ^_^
    // minRow: Min. Rows
    // minCol: Min. Cols
    // resize: Resize the worksheet according to the array size
    // clear: Clear all cell values before updating the cells
    const {url, gid} = await a2gs.updateWorkSheet(sheetName, array2d,
        {margin: 2, minRow: 10, minCol: 10, resize: true, clear: true});

    // get doc info object (from google-spreadsheet)
    const docInfo = await a2gs.getDocInfo();
    const {id, title, updated, worksheets, author} = docInfo;

    // get all sheet names
    const sheetNames = docInfo.worksheets.map(x => x.title);

    // get workSheet
    const workSheet = docInfo.worksheets[0];

    // query data from spreadsheet
    const array = await a2gs.getWorkSheetDataAsArray("sheetName");
    const csv = await a2gs.getWorkSheetDataAsCsv("sheetName");

    // you can also pass an array as parameters, it will return array
    const arrayList = await a2gs.getWorkSheetDataAsArray(["sheetName"]);
    const csvList = await a2gs.getWorkSheetDataAsCsv(["sheetName"]);
}
