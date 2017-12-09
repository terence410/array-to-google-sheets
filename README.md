# Array To Google Sheets #

A simple Node.js module for updating a 2 dimensional array into Google Sheets (Spreadsheets).
This module is built on top of another package [google-spreadsheet](https://www.npmjs.com/package/google-spreadsheet). 

# Features

- Support Number/String/Formula
- Auto resize the worksheet according to the array size
- Support generating formula
  - {formula: '=sum(%1:%2)', cells: [{row: 1, col: 1}, {row: 1, col: 3}]}
  - equivalent to =sum(A1:C1)
- Support Promise
                
# Installation

[![NPM Info](https://nodei.co/npm/array-to-google-sheets.png?downloads=true&downloadRank=true&stars=true)](https://www.npmjs.org/package/array-to-google-sheets)

# Basic Usage

```javascript
let ArrayToGoogleSheets = require('array-to-google-sheets');
async function start()
{
    let values = [
        [1, 2, 3.5555],
        [4, 5, 6],
        ['a', 'b', 'c'],
        [{formula: '=sum(%1:%2)', cells: [{row: 1, col: 1}, {row: 1, col: 3}]}], // =sum(A1:C1)
    ];
    let docKey = '{Google Sheets Key}';
    let creds = require('./creds.json');
    let a2gs = new ArrayToGoogleSheets(docKey, creds);
    try{
        await a2gs.updateGoogleSheets(sheetName, values, 
            {margin: 2, minRow: 10, minCol: 10, resize: true, clear: true});
    }catch(err){
        
    }
}
```

__Options__
- margin: Extra blank cells, for better styling ^_^
- minRow: Min. Rows
- minCol: Min Cols
- resize: Resize the worksheet according to the array size
- clear: Clear all cell values

# docKey 

- Every Google Sheets has a unique key in the URL
- https://docs.google.com/spreadsheets/d/{docKey}/

# Creds

1. Go to the Google [Developers Console](https://console.developers.google.com/cloud-resource-manager)
2. Select or Create Project
3. Dashboard > Enable APIs and Services > Enable the Drive API for your project
4. Credentials > Create Service Account Key
5. Select Json Key type and save the downloaded json file to your project
6. Once you have created the services account, you will have an email xxx@xxx.iam.gserviceaccount.com. **Go to your Goggle Sheets file and shared the edit permission to the email address.**
2. For more details, please refer to https://www.npmjs.com/package/google-spreadsheet

# Formula 

```javascript
let values = [
    [{formula: '=sum(%1:%2)', cells: [{row: 1, col: 1}, {row: 1, col: 3}]}], // =sum(A1:C1)
    [{formula: '=%1/50', cells: [{row: 1, col: 3}]}], // =C1/50
    [{formula: '=sum(%1:%2)', cells: [{row: 'this', col: 1}, {row: 'this', col: 3}]}], // =sum(A3:C3)
    [{formula: '=sum(%1:%2)', cells: [{row: 1, col: 'this'}, {row: 3, col: 'this'}]}], // =sum(A1:A3);
    [{formula: '=sum(%1:%2)', cells: [{row: 1, col: 0}, {row: 1, col: 0}]}], // =sum(1:1);
    [{formula: '=sum(%1:%2)', cells: [{row: 1}, {row: 1}]}], // =sum(1:1);
    [{formula: '=sum(%1:%2)', cells: [{row: 0, col: 2}, {row: 0, col: 2}]}] // =sum(B:B);
];
```

# Links
- https://developers.google.com/google-apps/spreadsheets/
- https://www.npmjs.com/package/google-spreadsheet

# License

array-to-google-sheets is free and unencumbered public domain software. For more information, see the accompanying UNLICENSE file.
