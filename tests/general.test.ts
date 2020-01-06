import { assert, expect } from "chai";
import {config} from "dotenv";
config();
import "mocha";
import {ArrayToGoogleSheets} from "../src";

const docKey = process.env.DOC_KEY || "";
const creds = process.env.CRED_FILE || "";

describe("general", () => {
    const a2gs = new ArrayToGoogleSheets(docKey, creds);

    it("Array with Number/String", async () => {
        const sheetName = "Testing1";
        const array2d = [
            [1, 2, 3.5555],
            [4, 5, 6],
            ["a", "b", "c"],
            ["comma,", "\"quote\"", ",mixed\",", "line1\rline2", "line1\r\nline2", "line1\nline2", "line1\rline2\r\nline3\nline4\"\"\n\nline6"],
            [{formula: "=sum(%1:%2)", cells: [{row: 1, col: 1}, {row: 1, col: 3}]}], // =sum(A1:C1)
        ];

        try {
            const result = await a2gs.updateWorkSheet(sheetName, array2d,
                {margin: 2, resize: true, clear: true});
            assert.hasAllKeys(result, ["url", "gid"]);
        } catch (err) {
            console.log(err);
            assert.isFalse(true);
        }
    });

    it("Array With Formula", async () => {
        // pls refer to the README how to set this value
        const sheetName = "Testing2";
        const array2d = [
            [1, 2, 3, {formula: "=sum(%1:%2)", cells: [{row: 1, col: 1}, {row: 1, col: 3}]}], // =sum(A1:C1)
            [4, 5, 6, {formula: "=%1/50", cells: [{row: 1, col: 3}]}], // =C1/50
            [7, 8, 9, {formula: "=sum(%1:%2)", cells: [{row: "this", col: 1}, {row: "this", col: 3}]}], // =sum(A3:C3)
            [{formula: "=sum(%1:%2)", cells: [{row: 1, col: "this"}, {row: 3, col: "this"}]}], // =sum(A1:A3);
            [{formula: "=sum(%1:%2)", cells: [{row: 1, col: 0}, {row: 1, col: 0}]}], // =sum(1:1);
            [{formula: "=sum(%1:%2)", cells: [{row: 1}, {row: 1}]}], // =sum(1:1);
            [{formula: "=sum(%1:%2)", cells: [{row: 0, col: 2}, {row: 0, col: 2}]}], // =sum(B:B);
        ];

        try {
            const result = await a2gs.updateWorkSheet(sheetName, array2d, {
                margin: 5,
                minRow: 10,
                minCol: 10,
                clear:  false,
                resize: false,
            });
            assert.hasAllKeys(result, ["url", "gid"]);
        } catch (err) {
            console.log(err);
            assert.isFalse(true);
        }
    });

    it("Get WorkSheets", async () => {
        const docInfo = await a2gs.getDocInfo();
        const sheetNames = docInfo.worksheets.map(x => x.title);
        assert.isAtLeast(sheetNames.length, 1);
    });

    it("Get WorkSheets As Array", async () => {
        const sheetNames = ["Testing1", "Testing2", "Unknown"];
        const array2dList = await a2gs.getWorkSheetDataAsArray(sheetNames);
        assert.equal(array2dList.length, sheetNames.length);

        const array2d1 = await a2gs.getWorkSheetDataAsArray("Testing1");
        assert.isTrue(Array.isArray(array2d1));
        assert.deepEqual(array2dList[0], array2d1);

        const array2d2 = await a2gs.getWorkSheetDataAsArray("Testing2");
        assert.isTrue(Array.isArray(array2d2));
        assert.deepEqual(array2dList[1], array2d2);

        const array2d3 = await a2gs.getWorkSheetDataAsArray("Unknown");
        assert.isUndefined(array2d3);
        assert.isUndefined(array2dList[2]);
    });

    it("Get WorkSheets as csv", async () => {
        const sheetNames = ["Testing1", "Testing2", "Unknown"];
        const csvList = await a2gs.getWorkSheetDataAsCsv(sheetNames);
        assert.equal(csvList.length, sheetNames.length);
        // fs.writeFileSync("./tests/test1.csv", csvList[0]);

        const csv1 = await a2gs.getWorkSheetDataAsCsv("Testing1");
        assert.isTrue(typeof csv1 === "string");
        assert.equal(csv1, csvList[0]);

        const csv2 = await a2gs.getWorkSheetDataAsCsv("Testing2");
        assert.isTrue(typeof csv2 === "string");
        assert.equal(csv2, csvList[1]);

        const csv3 = await a2gs.getWorkSheetDataAsCsv("Unknown");
        assert.isUndefined(csv3);
        assert.isUndefined(csvList[2]);
    });

    it("Test Errors", async () => {
        try {
            const a2gs1 = new ArrayToGoogleSheets("", {});
            await a2gs1.updateWorkSheet("sheet1", []);
        } catch (err) {
            assert.equal(err.message, "Spreadsheet key not provided.");
        }

        try {
            const a2gs2 = new ArrayToGoogleSheets(docKey, {});
            await a2gs2.updateWorkSheet("sheet1", []);
        } catch (err) {
            assert.equal(err.message, "No key or keyFile set.");
        }

        try {
            const a2gs3 = new ArrayToGoogleSheets(docKey, creds);
            await a2gs3.updateWorkSheet("sheet1", [1, 2, 3]);
        } catch (err) {
            assert.equal(err.message, "Values must be a 2 dimensional array.");
        }
    });
});
