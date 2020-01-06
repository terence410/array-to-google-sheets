# Array To Google Sheets #

[![NPM version][npm-image]][npm-url]
[![build status][travis-image]][travis-url]
[![Test coverage][codecov-image]][codecov-url]
[![David deps][david-image]][david-url]

[npm-image]: https://img.shields.io/npm/v/array-to-google-sheets.svg
[npm-url]: https://npmjs.org/package/ts-datastoreorm
[travis-image]: https://img.shields.io/travis/eggjs/egg.svg?style=flat-square
[travis-url]: https://travis-ci.org/eggjs/egg
[codecov-image]: https://img.shields.io/codecov/c/github/terence410/array-to-google-sheets.svg?style=flat-square
[codecov-url]: https://codecov.io/gh/terence410/array-to-google-sheets
[david-image]: https://img.shields.io/david/terence410/array-to-google-sheets.svg?style=flat-square
[david-url]: https://david-dm.org/terence410/array-to-google-sheets

Update a 2 dimensional array into Google Sheets (Spreadsheets).
You can also get back the Spreadsheets in array or csv format.
This module is built on top of another package [google-spreadsheet](https://www.npmjs.com/package/google-spreadsheet). 

# Features

- Support Number/String/Formula
- Auto resize the worksheet according to the array size
- Support generating formula
  - {formula: '=sum(%1:%2)', cells: [{row: 1, col: 1}, {row: 1, col: 3}]}
  - equivalent to =sum(A1:C1)
- Return the url and gid of the sheet upon edit
- Get Spreadsheets data in array or csv format
- Written in Typescript

# Basic Usage

```typescript
const ArrayToGoogleSheets = require("array-to-google-sheets");
import {ArrayToGoogleSheets} from "array-to-google-sheets"; // typescript

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

```

# docKey 

Every Google Sheets has a unique key in the URL
https://docs.google.com/spreadsheets/d/{docKey}/

# KeyFilename / Service Account

- Create a Google Cloud Project
- [Create Service Account](https://console.cloud.google.com/iam-admin/serviceaccounts/create)
  - Service account details > Choose any service account name > CREATE
  - Grant this service account access to project > CONTINUE
  - Grant users access to this service account ( > CREATE KEY
- In the JSON key file, you will find an email xxx@xxx.iam.gserviceaccount.com. 
- Go to your Google Sheets file and shared the edit permission to the email address.
- If things doesnt work, try to enable the Drive API first.
  -  [APIs and Services](https://console.cloud.google.com/apis/dashboard) > + Enable APIS AND SERVICES > Search Google Drive API > Enable

# Formula Example

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
