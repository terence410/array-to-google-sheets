# Array To Google Sheets #

[![NPM version][npm-image]][npm-url]
[![build status][travis-image]][travis-url]
[![Test coverage][codecov-image]][codecov-url]
[![David deps][david-image]][david-url]

[npm-image]: https://img.shields.io/npm/v/array-to-google-sheets.svg
[npm-url]: https://npmjs.org/package/array-to-google-sheets
[travis-image]: https://img.shields.io/travis/terence410/array-to-google-sheets.svg?style=flat-square
[travis-url]: https://travis-ci.org/terence410/array-to-google-sheets
[codecov-image]: https://img.shields.io/codecov/c/github/terence410/array-to-google-sheets.svg?style=flat-square
[codecov-url]: https://codecov.io/gh/terence410/array-to-google-sheets
[david-image]: https://img.shields.io/david/terence410/array-to-google-sheets.svg?style=flat-square
[david-url]: https://david-dm.org/terence410/array-to-google-sheets

Update a 2 dimensional array into Google Sheets (Spreadsheets). You can also get back data in array or csv.
The library is build with [Google Sheets API v4](https://developers.google.com/sheets).

# Features

- Tested with huge amount of data (10000 rows x 100 columns) and optimized memory usage.
- Support generating formula
  - {formula: '=sum(%1:%2)', cells: [{row: 1, col: 1}, {row: 1, col: 3}]}
  - equivalent to =sum(A1:C1)
- Get Spreadsheets data in array or csv format
- Can handle multiple sheet
- Can update single row and single
- An experiment feature of converting the sheet as array object

# Usage

```typescript
import {ArrayToGoogleSheets, IUpdateOptions} from "array-to-google-sheets"; // typescript

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

```

# spreadsheetId 

Every Google Sheets has a unique key in the URL
https://docs.google.com/spreadsheets/d/{spreadsheetId}/

# KeyFilename / Service Account

- Create a Google Cloud Project
- [Create Service Account](https://console.cloud.google.com/iam-admin/serviceaccounts/create)
  - Service account details > Choose any service account name > CREATE
  - Grant this service account access to project > CONTINUE
  - Grant users access to this service account ( > CREATE KEY
  - Save the key file into your project
- Enable Drive API & Google Sheets API
  -  [APIs and Services](https://console.cloud.google.com/apis/dashboard) > Enable APIS AND SERVICES 
  - Search Google Drive API > Enable
  - Search Google Sheets API > Enable
- Enable Google Sheets API
- Open the JSON key file, you will find an email xxx@xxx.iam.gserviceaccount.com. 
- Go to your Google Spreadsheets and shared the edit permission to the email address.

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

# Experiment Object Sheet Feature
```typescript
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
```

| value1 | value2/string | value3/number | value4/boolean | value5/date | value6/number[] | value7/string[] | value8/ignore |
| ------ | ------------- | ------------- | -------------- | ----------- | --------------- | --------------- | ------------- |
| key0 | value0 | 0.7238840059 | TRUE | 2020-03-03T05:41:02.926Z | 0, 1, 2, 0.865| 0, a, b, c | 8 |
| key1 | value1 | 0.2963643265 | FALSE | 2020-03-03T05:41:03.149Z | 1, 1, 2, 0.995| 1, a, b, c | 8 |

The above table will be converted as:
```typescript

interface IObject {value1: string; value2: string; value3: number; value4: boolean; value5: Date; value6: number[]; value7: string[];}
const objectSheet = await sheet.exportAsObjectSheet<IObject>();
const objectInterface = objectSheet.getInterface();
const {headers, size, rawValues, rawHeaders} = objectSheet;
const firstItem = objectSheet.get(0);

for (const item of objectSheet) {
    console.log(item.toObject());
    item.value1 = "new Value";
    // this will only update the changed cell values to minimize modifing the original values as much as possible
    await item.save();
}
/* 
[
  {
    value1: 'key0',
    value2: 'value0',
    value3: 0.7238840059,
    value4: true,
    value5: 2020-03-03T05:41:02.926Z
    value6: [ 0, 1, 2, 0.865 ],
    value7: [ '0', 'a', 'b', 'c' ]
  }
  {
    value1: 'key1',
    value2: 'value1',
    value3: 0.2963643265,
    value4: false,
    value5: 2020-03-03T05:41:03.149Z
    value6: [ 1, 1, 2, 0.995 ],
    value7: [ '1', 'a', 'b', 'c' ]
  }
]
*/
```


# Links
- https://www.npmjs.com/package/array-to-google-sheets
- https://www.npmjs.com/package/google-auth-library
