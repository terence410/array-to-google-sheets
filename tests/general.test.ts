import { assert, expect } from "chai";
import {config} from "dotenv";
config();
import csvParse from "csv-parse";
import "mocha";
import {ArrayToGoogleSheets} from "../src/ArrayToGoogleSheets";

const spreadsheetId = process.env.SPREADSHEET_ID || "";
const keyFilename = process.env.KEY_FILENAME || "";
const clientEmail = process.env.CLIENT_EMAIL;
const privateKey = process.env.PRIVATE_KEY;
const credentials = clientEmail && privateKey ? {client_email: clientEmail, private_key: privateKey} : undefined;

describe.only("general", () => {
    let memory = 0;
    beforeEach(() => {
        memory = process.memoryUsage().heapUsed;
    });

    afterEach(() => {
        const memoryDiff = process.memoryUsage().heapUsed - memory;
        console.log(`Used memory: ${memoryDiff / 1024 / 1024 | 0}MB`);
    });

    it("basic operation", async () => {
        const googleSheets = new ArrayToGoogleSheets({keyFilename, credentials});
        const spreadsheet = await googleSheets.getSpreadsheet(spreadsheetId);
        const {spreadsheetUrl, properties} = spreadsheet;
        const {title, locale, timeZone, defaultFormat} = properties;

        const sheetName = "basic";
        const sheet = await spreadsheet.findSheet(sheetName);
        if (sheet) {
            const result = await sheet.delete();
        }

        // get sheet again
        const newSheet = await spreadsheet.findOrCreateSheet(sheetName);

        //  update
        const values1 = [
            [1, 2, 3],
            [1.1, 2.2, -3.33],
            ["abc", "cde", "xyz"],
        ];
        const updateResult1 = await newSheet.update(values1, {minRow: 3, minColumn: 3, margin: 2, fitToSize: true});
        const resultValues1 = await newSheet.getValues();
        assert.deepEqual(values1, resultValues1);

        // update in simple way
        const values2 = [
            [1, 2, 3],
            [{formula: "=sum(%1:%2)", cells: [{row: 1}, {row: 1}]}],
        ];
        const updateResult2 = await spreadsheet.updateSheet(sheetName, values2, {clearAllValues: true});
        const resultValues2 = await newSheet.getValues();
        assert.deepEqual(resultValues2, [[1, 2, 3], [6]]);
    });

    it("formula", async () => {
        const googleSheets = new ArrayToGoogleSheets({keyFilename, credentials});
        const spreadsheet = await googleSheets.getSpreadsheet(spreadsheetId);
        const sheetName = "formula";
        
        const values = [
            [1, 2, 3, {formula: "=sum(%1:%2)", cells: [{row: 1, col: 1}, {row: 1, col: 3}]}], // =sum(A1:C1)
            [4, 5, 6, {formula: "=%1/50", cells: [{row: 1, col: 3}]}], // =C1/50
            [7, 8, 9, {formula: "=sum(%1:%2)", cells: [{row: "this", col: 1}, {row: "this", col: 3}]}], // =sum(A3:C3)
            [{formula: "=sum(%1:%2)", cells: [{row: 1, col: "this"}, {row: 3, col: "this"}]}], // =sum(A1:A3);
            [{formula: "=sum(%1:%2)", cells: [{row: 1}, {row: 1}]}], // =sum(1:1);
            [{formula: "=sum(%1:%2)", cells: [{row: 2}, {row: 2}]}], // =sum(1:1);
            [{formula: "=sum(%1:%2)", cells: [{col: 3}, {col: 3}]}], // =sum(B:B);
        ];

        const sheet = await spreadsheet.findOrCreateSheet(sheetName);
        await sheet.update(values);
        const resultValues = await sheet.getValues();
        assert.deepEqual(resultValues, [
                [ 1, 2, 3, 6 ],
                [ 4, 5, 6, 0.06 ],
                [ 7, 8, 9, 24 ],
                [ 12 ],
                [ 12 ],
                [ 15.06 ],
                [ 18 ],
            ],
        );
    });

    it("csv", async () => {
        const googleSheets = new ArrayToGoogleSheets({keyFilename, credentials});
        const spreadsheet = await googleSheets.getSpreadsheet(spreadsheetId);

        const sheetName = "csv";
        const sheet = await spreadsheet.findOrCreateSheet(sheetName);

        const values1 = [
            [1, 234, 567],
            [1.1111, 2.2222, -3.3333],
            [",,,,,,,,,", "----------", "``````````"],
            ["\t\t\t", "     ", "~!@#$%^&*()-="],
            ["comma,", "\"quote\"", ",mixed\",", "line1\rline2"],
            ["line1\r\nline2", "line1\nline2", "line1\rline2\r\nline3\nline4\"\"\n\nline6"],
        ];
        await sheet.update(values1, {clearAllValues: true});

        // export into csv
        const csvString = await sheet.exportAsCsv("./index.csv");

        const resultValues = await new Promise((resolve, reject) => {
            csvParse(csvString, {delimiter: ",", relax_column_count: true, cast: true}, (err, csvJson) => {
                if (err) {
                    return resolve(err);
                }

                resolve(csvJson);
            });
        });

        assert.deepEqual(resultValues, values1);
    });

    // completed in 23s for 10000 * 100 records, Need around 20MB allocations
    it.skip("massive operation", async () => {
        function checkMemory() {
            const {heapUsed, heapTotal} = process.memoryUsage();
            console.log(`heap: ${heapUsed / 1024 / 1024 | 0}MB, healTotal: ${heapTotal / 1024 / 1024 | 0}MB`);
        }
        setInterval(checkMemory, 1000);
        checkMemory();

        const googleSheets = new ArrayToGoogleSheets({keyFilename, credentials});
        const spreadsheet = await googleSheets.getSpreadsheet(spreadsheetId);

        const sheetName = "massive";
        const totalRow = 10000;
        const totalColumn = 100;
        const values = Array(totalRow).fill(0).map((x, i) => Array(totalColumn).fill(i));

        const sheet = await spreadsheet.findOrCreateSheet(sheetName);

        const updateResult = await sheet.update(values);
        assert.equal(updateResult.updatedRows, totalRow);
        assert.equal(updateResult.updatedColumns, totalColumn);
        assert.equal(updateResult.updatedCells, totalRow * totalColumn);

        const resultValues = await sheet.getValues();
        assert.deepEqual(resultValues, values);
    }).timeout(60 * 1000);
});
