let ArrayToGoogleSheets = require('../index');
let assert = require('chai').assert;

let timeout = t => new Promise(cb => setInterval(cb, t));

// pls refer to the README how to get the docKey and creds
let docKey = "1Kaf_UF_rL_LoluJ4kG2aNZP8oU7-SxRLdUMcx85VkhY";
let creds = require('../creds.json');
let a2gs = new ArrayToGoogleSheets(docKey, creds);

describe('Update Google Sheets', () => {
    it('Array with Number/String', async () => {
        let sheetName = "Testing1";
        let values = [
            [0],
            [1, 2, 3, 4.5],
            ['a', 'b', 'c'],
            ['=sum(A2:D2)', {formula: '=SUM(A2:D2)'}]
        ];

        try {
            await a2gs.updateGoogleSheets(sheetName, values, {margin: 1, minRow: 0, minCol: 1});
        }catch (err){
            console.log('caught error');
            console.log(err);
        }

    }).timeout(30 * 1000); // we have to wait longer for large data set

    it('Array With Formula', async () => {
        // pls refer to the README how to set this value
        let sheetName = "testing2";
        let values = [
            [1, 2, 3, {formula: '=sum(%1:%2)', cells: [{row: 1, col: 1}, {row: 1, col: 3}]}], // =sum(A1:C1)
            [4, 5, 6, {formula: '=%1/50', cells: [{row: 1, col: 3}]}], // =C1/50
            [7, 8, 9, {formula: '=sum(%1:%2)', cells: [{row: 'this', col: 1}, {row: 'this', col: 3}]}], // =sum(A3:C3)
            [{formula: '=sum(%1:%2)', cells: [{row: 1, col: 'this'}, {row: 3, col: 'this'}]}], // =sum(A1:A3);
            [{formula: '=sum(%1:%2)', cells: [{row: 1, col: 0}, {row: 1, col: 0}]}], // =sum(1:1);
            [{formula: '=sum(%1:%2)', cells: [{row: 1}, {row: 1}]}], // =sum(1:1);
            [{formula: '=sum(%1:%2)', cells: [{row: 0, col: 2}, {row: 0, col: 2}]}], // =sum(B:B);
        ];

        try {
            await a2gs.updateGoogleSheets(sheetName, values, {margin: 5, minRow: 10, minCol: 10, clear: false, resize: false});
        }catch (err){
            console.log('caught error');
            console.log(err);
        }

    }).timeout(30 * 1000); // we have to wait longer for large data set

    it('Test Errors', async () => {
        try {
            let a2gs = new ArrayToGoogleSheets("", {});
            await a2gs.updateGoogleSheets("sheet1", []);
        }catch (err){
            assert.equal(err.message, 'Spreadsheet key not provided.');
        }

        try {
            let a2gs = new ArrayToGoogleSheets(docKey, {});
            await a2gs.updateGoogleSheets("sheet1", []);
        }catch (err){
            assert.equal(err.message, 'No key or keyFile set.');
        }

        try {
            let a2gs = new ArrayToGoogleSheets(docKey, creds);
            await a2gs.updateGoogleSheets("sheet1", [1, 2, 3]);
        }catch (err){
            assert.equal(err.message, 'values must be 2 dimensional array.');
        }

    }).timeout(5000);
});

