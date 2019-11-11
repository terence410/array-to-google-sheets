# Array To Google Sheets #

A simple Node.js (Written in typescript) module for updating a 2 dimensional array into Google Sheets (Spreadsheets).
You can asl get back the Spreadsheets in array or csv format.
This module is built on top of another package [google-spreadsheet](https://www.npmjs.com/package/google-spreadsheet). 

[![NPM version](https://badge.fury.io/js/array-to-google-sheets.png)](http://badge.fury.io/js/array-to-google-sheets)

# Features

- Support Number/String/Formula
- Auto resize the worksheet according to the array size
- Support generating formula
  - {formula: '=sum(%1:%2)', cells: [{row: 1, col: 1}, {row: 1, col: 3}]}
  - equivalent to =sum(A1:C1)
- Return the url and gid of the sheet upon edit
- Get Spreadsheets data in array or csv format
- Written in Typescript
                
# Installation

[![NPM Info](https://nodei.co/npm/array-to-google-sheets.png?downloads=true&downloadRank=true&stars=true)](https://www.npmjs.org/package/array-to-google-sheets)

# Basic Usage

```javascript
const ArrayToGoogleSheets = require("array-to-google-sheets");
import {ArrayToGoogleSheets} from "array-to-google-sheets"; // typescript

async function main() {
    const array2d = [
        [1, 2, 3.5555],
        [4, 5, 6],
        ["a", "b", "c"],
        [{formula: "=sum(%1:%2)", cells: [{row: 1, col: 1}, {row: 1, col: 3}]}], // =sum(A1:C1)
    ];
    const docKey = "Google Sheets Key";
    const creds = "./google-creds.json"; // file or json object both ok
    const a2gs = new ArrayToGoogleSheets(docKey, creds);
    const sheetName = "Sheet Name";
    const {url, gid} = await a2gs.updateGoogleSheets(sheetName, array2d, 
        {margin: 2, minRow: 10, minCol: 10, resize: true, clear: true});
    
    // get data 
    const array2dObject = await a2gs.getGoogleSheets(["sheetname"]);
    const array2d = await a2gs.getGoogleSheet("sheetname");
    const csvObject = await a2gs.getGoogleSheetsAsCsv(["sheetname"]);
    const csv = await a2gs.getGoogleSheetAsCsv("sheetname");
}
```

__Options__
- margin: Extra blank cells, for better styling ^_^
- minRow: Min. Rows
- minCol: Min. Cols
- resize: Resize the worksheet according to the array size
- clear: Clear all cell values before updating the cells

# docKey 

- Every Google Sheets has a unique key in the URL
- https://docs.google.com/spreadsheets/d/{docKey}/

# Creds

1. Go to the Google [Developers Console](https://console.developers.google.com/cloud-resource-manager)
2. Select or Create Project
3. Dashboard > [APIs and Services](https://console.cloud.google.com/apis/dashboard) > + Enable APIS AND SERVICES (Title Bar) > Search Drive > Enable the Drive API for your project
4. Dashboard > APIs & Services > Credentials > Create Service Account Key > You do not need to select any role 
5. For Key type, select JSON > Save the downloaded json file to your project
6. Once you have created the services account, you will found an email xxx@xxx.iam.gserviceaccount.com. 
7. Go to your Google Sheets file and shared the edit permission to the email address.

# Formula 

```javascript
let values = [
    [{formula: '=sum(%1:%2)', cells: [{row: 1, col: 1}, {row: 1, col: 3}]}], 
    // =sum(A1:C1)
    [{formula: '=%1/50', cells: [{row: 1, col: 3}]}], 
    // =C1/50
    [{formula: '=sum(%1:%2)', cells: [{row: 'this', col: 1}, {row: 'this', col: 3}]}], 
    // =sum(A3:C3)
    [{formula: '=sum(%1:%2)', cells: [{row: 1, col: 'this'}, {row: 3, col: 'this'}]}],
     // =sum(A1:A3);
    [{formula: '=sum(%1:%2)', cells: [{row: 1, col: 0}, {row: 1, col: 0}]}], 
    // =sum(1:1);
    [{formula: '=sum(%1:%2)', cells: [{row: 1}, {row: 1}]}], 
    // =sum(1:1);
    [{formula: '=sum(%1:%2)', cells: [{row: 0, col: 2}, {row: 0, col: 2}]}] 
    // =sum(B:B);
];
```

# Links
- https://developers.google.com/google-apps/spreadsheets/
- https://www.npmjs.com/package/google-spreadsheet
- https://www.npmjs.com/package/array-to-google-sheets
