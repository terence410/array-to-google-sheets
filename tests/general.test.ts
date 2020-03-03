import { assert, expect } from "chai";
import {config} from "dotenv";
config();
import crypto from "crypto";
import csvParse from "csv-parse";
import "mocha";
import {ArrayToGoogleSheets} from "../src/ArrayToGoogleSheets";
import {IRow} from "../src/types";

export function generateRandomString(length: number = 16) {
    const value = crypto.randomBytes(Math.ceil(length / 2)).toString("hex");
    return value.substr(0, length);
}

function getMemory() {
    const {heapUsed, heapTotal} = process.memoryUsage();
    return {heapUsed, heapTotal};
    // console.log(`heap: ${heapUsed / 1024 / 1024 | 0}MB, healTotal: ${heapTotal / 1024 / 1024 | 0}MB`);
}

const spreadsheetId = process.env.SPREADSHEET_ID || "";
const keyFilename = process.env.KEY_FILENAME || "";
const clientEmail = process.env.CLIENT_EMAIL;
const privateKey = process.env.PRIVATE_KEY;
const credentials = clientEmail && privateKey ? {client_email: clientEmail, private_key: privateKey} : undefined;
const sheetName = generateRandomString();
const emptySheetName = generateRandomString();
const googleSheets = new ArrayToGoogleSheets({keyFilename, credentials});

describe.only("general", () => {
    afterEach(async () => {
        const spreadsheet = await googleSheets.getSpreadsheet(spreadsheetId);
        const sheet = await spreadsheet.findSheet(sheetName);
        if (sheet) {
            await sheet.delete();
        }
    });

    it("basic operation", async () => {
        const spreadsheet = await googleSheets.getSpreadsheet(spreadsheetId);
        const {spreadsheetUrl, properties} = spreadsheet;
        const {title, locale, timeZone, defaultFormat} = properties;

        const sheet = await spreadsheet.findSheet(sheetName);
        if (sheet) {
            const result = await sheet.delete();
        }

        // get sheet again
        const newSheet = await spreadsheet.findOrCreateSheet(sheetName);

        // find sheet
        const sheets = spreadsheet.findSheets([sheetName]);
        assert.equal(sheets.length, 1);

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

        // clean up
        await newSheet.delete();

        // try empty sheet
        const emptySheet = await spreadsheet.findOrCreateSheet(emptySheetName);
        const emptyValues = await emptySheet.getValues();
        assert.deepEqual(emptyValues, []);
        await emptySheet.delete();
    });

    it("formula", async () => {
        const spreadsheet = await googleSheets.getSpreadsheet(spreadsheetId);

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

        // clean up
        await sheet.delete();
    });

    it("csv", async () => {
        const spreadsheet = await googleSheets.getSpreadsheet(spreadsheetId);
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

        // clean up
        await sheet.delete();
    });

    it("updateRow operations", async () => {
        const spreadsheet = await googleSheets.getSpreadsheet(spreadsheetId);
        const sheet = await spreadsheet.findOrCreateSheet(sheetName);

        // resize it first
        await sheet.clear();
        await sheet.resize(10, 10);

        // add rows
        const total = 5;
        const promises: any[] = [];
        for (let i = 0; i < total; i++) {
            const row: IRow = Array(total).fill(0).map((x, j) => j + i);
            row.push({formula: "=sum(%1:%2)", cells: [{row: "this", col: 1}, {row: "this", col: total}]});
            promises.push(sheet.updateRow(i, row));
        }
        await Promise.all(promises);

        const finalValues = await sheet.getValues();
        assert.equal(finalValues.length, total);

        for (let i = 0; i < total; i++) {
            const row = Array(total).fill(0).map((x, j) => j + i);
            const sum = row.reduce((a, b) => a + b, 0);
            assert.equal(sum, finalValues[i][total]);
        }
    });

    it("updateCell operations", async () => {
        const spreadsheet = await googleSheets.getSpreadsheet(spreadsheetId);
        const sheet = await spreadsheet.findOrCreateSheet(sheetName);

        // resize it first
        await sheet.clear();
        await sheet.resize(10, 10);

        // add rows
        const total = 4;
        const promises: any[] = [];
        for (let i = 0; i < total; i++) {
            const formula = {formula: "=sum(%1:%2)", cells: [{row: "this", col: i + 1}, {row: "this", col: i + 1}]};
            promises.push(sheet.updateCell(i, i, i));
            promises.push(sheet.updateCell(i, i + 1, formula));
        }
        await Promise.all(promises);

        const sumFormula = {formula: "=sum(%1:%2)", cells: [{row: 1, col: 1}, {row: total, col: total}]};
        await sheet.updateCell(total, total, sumFormula);

        const finalValues = await sheet.getValues();
        assert.equal(finalValues.length, total + 1);
        assert.equal(finalValues[total][total], Math.pow(total - 1, 2));
    });

    // completed in 23s for 10000 * 100 records, Need around 20MB allocations
    it("massive operation with limited ram usage", async () => {
        // max allow used memory is 30MB
        const maxMemoryDiff = 20 * 1024 * 1024;
        const initialMemory = getMemory();

        function checkMemory() {
            const currentMemory = getMemory();
            const diffHeapUsed = currentMemory.heapUsed - initialMemory.heapUsed;
            const diffHeapTotal = currentMemory.heapTotal - initialMemory.heapTotal;
            if (diffHeapUsed > maxMemoryDiff) {
                throw Error(`Memory error. Difference in Heap Used: ${diffHeapUsed / 1024 / 1024 | 0}`);
            } else if (diffHeapTotal > maxMemoryDiff) {
                throw Error(`Memory error. Difference in Heap Total: ${diffHeapTotal / 1024 / 1024 | 0}`);
            }
        }
        const id = setInterval(checkMemory, 1000);

        const spreadsheet = await googleSheets.getSpreadsheet(spreadsheetId);
        const totalRow = 2000;
        const totalColumn = 50;
        const values = Array(totalRow).fill(0).map((x, i) => Array(totalColumn).fill(i));

        const sheet = await spreadsheet.findOrCreateSheet(sheetName);

        const updateResult = await sheet.update(values, {clearAllValues: true});
        assert.equal(updateResult.updatedRows, totalRow);
        assert.equal(updateResult.updatedColumns, totalColumn);
        assert.equal(updateResult.updatedCells, totalRow * totalColumn);

        const resultValues = await sheet.getValues();
        assert.equal(resultValues.length, totalRow);
        assert.deepEqual(resultValues[0], values[0]);
        assert.deepEqual(resultValues.slice(-1), values.slice(-1));

        // clean up
        await sheet.delete();

        clearInterval(id);
    }).timeout(20 * 1000);
});
