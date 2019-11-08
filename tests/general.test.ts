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
        const values = [
            [1, 2, 3.5555],
            [4, 5, 6],
            ["a", "b", "c"],
            ["comma,", "\"quote\"", ",mixed\","],
            [{formula: "=sum(%1:%2)", cells: [{row: 1, col: 1}, {row: 1, col: 3}]}], // =sum(A1:C1)
        ];

        try {
            const result = await a2gs.updateGoogleSheets(sheetName, values,
                {margin: 2, resize: true, clear: true});
            assert.hasAllKeys(result, ["url", "gid"]);
        } catch (err) {
            console.log(err);
            assert.isFalse(true);
        }
    });

    it("Get the worksheet", async () => {
        const sheetNames = ["Testing1"];
        const result = await a2gs.getGoogleSheets(sheetNames);
        assert.hasAllKeys(result, sheetNames);
    });

    it("Array With Formula", async () => {
        // pls refer to the README how to set this value
        const sheetName = "Testing2";
        const values = [
            [1, 2, 3, {formula: "=sum(%1:%2)", cells: [{row: 1, col: 1}, {row: 1, col: 3}]}], // =sum(A1:C1)
            [4, 5, 6, {formula: "=%1/50", cells: [{row: 1, col: 3}]}], // =C1/50
            [7, 8, 9, {formula: "=sum(%1:%2)", cells: [{row: "this", col: 1}, {row: "this", col: 3}]}], // =sum(A3:C3)
            [{formula: "=sum(%1:%2)", cells: [{row: 1, col: "this"}, {row: 3, col: "this"}]}], // =sum(A1:A3);
            [{formula: "=sum(%1:%2)", cells: [{row: 1, col: 0}, {row: 1, col: 0}]}], // =sum(1:1);
            [{formula: "=sum(%1:%2)", cells: [{row: 1}, {row: 1}]}], // =sum(1:1);
            [{formula: "=sum(%1:%2)", cells: [{row: 0, col: 2}, {row: 0, col: 2}]}], // =sum(B:B);
        ];

        try {
            const result = await a2gs.updateGoogleSheets(sheetName, values, {
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

    it("Test Errors", async () => {
        try {
            const a2gs1 = new ArrayToGoogleSheets("", {});
            await a2gs1.updateGoogleSheets("sheet1", []);
        } catch (err) {
            assert.equal(err.message, "Spreadsheet key not provided.");
        }

        try {
            const a2gs2 = new ArrayToGoogleSheets(docKey, {});
            await a2gs2.updateGoogleSheets("sheet1", []);
        } catch (err) {
            assert.equal(err.message, "No key or keyFile set.");
        }

        try {
            const a2gs3 = new ArrayToGoogleSheets(docKey, creds);
            await a2gs3.updateGoogleSheets("sheet1", [1, 2, 3]);
        } catch (err) {
            assert.equal(err.message, "Values must be 2 dimensional array.");
        }
    });
});
