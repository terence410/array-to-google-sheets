import fs from "fs";
import "./types";
// tslint:disable-next-line
import GoogleSpreadsheet, {DocInfo, SpreadsheetWorksheet} from "google-spreadsheet";

export type IOption = {
  minRow: number,
  minCol: number,
  margin: number,
  resize: boolean,
  clear: boolean,
};

type IFormulaCell = {
  row: number | "this",
  col: number | "this",
};

type ISpredsheetCell = {
  row: number,
  col: number,
  value: any,
};

export class ArrayToGoogleSheets {
  public doc: GoogleSpreadsheet;
  public credsJson: object;

  constructor(public readonly docKey: string, public readonly creds: any) {
    if (typeof creds === "string") {
      // this is file, convert to json
      this.credsJson = this.importJson(creds);
    } else {
      this.credsJson = creds;
    }

    this.doc = new GoogleSpreadsheet(this.docKey);
  }

  // region public methods: utils

  public convertColToAlphabet(num: number): string {
    const char = (10 + (num - 1) % 26).toString(36).toUpperCase();
    return num > 0 ? this.convertColToAlphabet((num - 1) / 26 | 0) + char : "";
  }

  public convertWid2gid(wid: string): number {
    const widVal = wid.length > 3 ? wid.substr(1) : wid;
    const xorVal = wid.length > 3 ? 474 : 31578;
    return parseInt(String(widVal), 36) ^ xorVal;
  }

  public importJson(file: string) {
    const data = fs.readFileSync(file).toString();
    return JSON.parse(data);
  }

  public getUrlObject(wid: string) {
    const gid = this.convertWid2gid(wid);
    const url = `https://docs.google.com/spreadsheets/d/${this.docKey}/edit#gid=${gid}`;
    return {url, gid};
  }

  public formatFormula(formulaFormat: string, cells: IFormulaCell[], currentRow: number, currentCol: number) {
    if (!Array.isArray(cells)) {
      return formulaFormat;
    }

    const arr = cells.map(cell => {
      const col = cell.col === "this" ? currentCol : cell.col;
      const row = cell.row === "this" ? currentRow : cell.row;
      return (col > 0 ? this.convertColToAlphabet(col) : "") + (row > 0 ? row : "");
    });

    return formulaFormat.replace(/%(\d+)/g, (_, m) => {
      return arr[--m];
    });
  }

  // endregion

  // region public methods
  public async updateGoogleSheetsData(sheet: SpreadsheetWorksheet, values: any[][], options: IOption) {
    let rowCount = Math.max(options.minRow, values.length) + options.margin;
    let colCount = Math.max(options.minCol, values.reduce((a, b) => Math.max(a, b.length), 0)) + options.margin;

    // apply options
    if (!options.resize) {
      rowCount = sheet.rowCount;
      colCount = sheet.colCount;
    }

    await this._resizeSheet(sheet, rowCount, colCount);
    const cells = await this._getCells(sheet, rowCount);

    // store the cells in array[row][column] format for ease of access
    const myCells: any[] = [];
    const updatedCells: any[] = [];

    // process the cells
    cells.forEach(cell => {
      if (!myCells[cell.row - 1]) {
        myCells[cell.row - 1] = [];
      }
      myCells[cell.row - 1].push(cell);

      // clear all the cells's value
      if (options.clear) {
        cell.value = "";
      }
    });

    // update the cell
    values.forEach((list, i) => {
      list.forEach((value, j) => {
        if (myCells[i] && myCells[i][j]) {
          if (!isNaN(parseFloat(value)) && isFinite(value)) {
            myCells[i][j].numericValue = value;
          } else if (typeof value === "object") {
            myCells[i][j].formula = this.formatFormula(value.formula, value.cells, i + 1, j + 1);
          } else {
            myCells[i][j].value = value;
          }
          updatedCells.push(myCells[i][j]);
        }
      });
    });

    // update the cells
    await this._updateCells(sheet, options.clear ? cells : updatedCells);
  }

  public async updateGoogleSheets(sheetName: string, values: any[], partialOption: Partial<IOption> = {}) {
    const option: IOption = Object.assign({
      minCol: 10,
      minRow: 20,
      margin: 2,
      resize: true,
      clear: true,
    }, partialOption);

    // the values array must be 2 dimensional
    if (values.length > 0 && !Array.isArray(values[0])) {
      throw new Error("Values must be 2 dimensional array.");
    }

    // get docInfo
    const docInfo = await this._getDocInfo();

    // search for existing worksheet
    let sheet = docInfo.worksheets.find(x => x.title.toLowerCase() === sheetName.toLowerCase());

    // create a new work sheet if not exist
    if (!sheet) {
        sheet = await this._addWorkSheet(sheetName);
    }

    // make sure sheet exist
    if (!sheet) {
      throw new Error(`Cannot create new sheet: ${sheetName}`);
    }

    // update array to the sheet
    await this.updateGoogleSheetsData(sheet, values, option);
    return this.getUrlObject(sheet.id);
  }

  public async getGoogleSheet(sheetName: string): Promise<any[] | undefined> {
    const data = await this.getGoogleSheets([sheetName]);
    if (sheetName in data) {
      return data[sheetName];
    }
  }

  public async getGoogleSheetAsCsv(sheetName: string): Promise<string | undefined> {
    const data = await this.getGoogleSheetsAsCsv([sheetName]);
    if (sheetName in data) {
      return data[sheetName];
    }
  }

  public async getGoogleSheets(sheetNames: string[] = []): Promise<{[key: string]: any[]}> {
    const docInfo = await this._getDocInfo();

    // get all sheet by default
    if (sheetNames.length === 0) {
      sheetNames = docInfo.worksheets.map(x => x.title);
    }

    const existingWorkSheets = docInfo.worksheets
        .filter(x => sheetNames.includes(x.title));
    const existingSheetNames = existingWorkSheets.map(x => x.title);

    const promises = existingWorkSheets.map(x => this._getCells(x, x.rowCount));
    const cellsList = await Promise.all(promises);

    // prepare result
    const result: {[key: string]: any[]} = {};
    for (let i = 0; i < cellsList.length; i++) {
      const cells = cellsList[i];
      const sheetName  = sheetNames[i];

      // turn into array2D
      let array2d: any[] = [];
      cells.forEach(cell => {
        if (!array2d[cell.row - 1]) {
          array2d[cell.row - 1] = [];
        }
        array2d[cell.row - 1].push(cell.value);
      });

      // filter empty rows
      array2d = array2d.filter(row => {
        return row.filter((cell: any[]) => cell.length > 0).length > 0;
      });

      // format and assign
      result[sheetName] = array2d;
    }

    return result;
  }

  public async getGoogleSheetsAsCsv(sheetNames: string[] = []): Promise<{[key: string]: string}> {
    const array2dObject = await this.getGoogleSheets(sheetNames);

    // escape strings
    const csvObject: {[key: string]: string} = {};
    for (const [key, array2d] of Object.entries(array2dObject)) {

      const escapedArray2d = array2d.map(row => {
        return row.map((cell: string) => {
          // we use double quote to wrap the strings if it has [,"\r\n]
          if (cell.match(/[,"\r\n]/)) {
            return `"${cell.replace(/"/g, "\"\"")}"`;
          }
          return cell;
        });
      });

      // format into csv
      csvObject[key] = escapedArray2d.map(row => row.join(",")).join("\r\n");
    }

    return csvObject;
  }

  // endregion

  // region private methods

  private async _getDocInfo(): Promise<DocInfo> {
    // connect to the service account
    await new Promise((resolve, reject) => {
      this.doc.useServiceAccountAuth(this.credsJson, (err) => {
        if (err) {
          return reject(err);
        }
        resolve();
      });
    });

    // open the document
    return await new Promise((resolve, reject) => {
      this.doc.getInfo((err, docInfo1) => {
        if (err) {
          return reject(err);
        }
        resolve(docInfo1);
      });
    });
  }

  private async _addWorkSheet(sheetName: string): Promise<SpreadsheetWorksheet> {
    return await new Promise((resolve, reject) => {
      this.doc.addWorksheet({title: sheetName}, (err, newSheet2) => {
        if (err) {
          return reject(err);
        }
        resolve(newSheet2);
      });
    });
  }

  private async _resizeSheet(sheet: SpreadsheetWorksheet, rowCount: number, colCount: number) {
    return new Promise((resolve, reject) => {
      sheet.resize({rowCount, colCount}, (err) => {
        if (err) {
          return reject(err);
        }

        resolve();
      });
    });
  }

  private async _getCells(sheet: SpreadsheetWorksheet, maxRow: number): Promise<ISpredsheetCell[]> {
    return new Promise((resolve, reject) => {
      sheet.getCells({
        "min-row": 1,
        "max-row": maxRow,
        "return-empty": true,
      }, (err, cells) => {
        if (err) {
          return reject(err);
        }
        resolve(cells);
      });
    });
  }

  private async _updateCells(sheet: SpreadsheetWorksheet, cells: any[][]) {
    await new Promise((resolve, reject) => {
      sheet.bulkUpdateCells(cells, (err) => {
        if (err) {
          return reject(err);
        }
        resolve();
      });
    });
  }

  // endregion
}
