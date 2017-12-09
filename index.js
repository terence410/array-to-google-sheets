let GoogleSpreadsheet = require('google-spreadsheet');

class ArrayToGoogleSheets
{
    /**
     *
     * @constructor
     * @param {String} docKey
     * @param {Object} creds
     */
    constructor(docKey, creds)
    {
        this.docKey = docKey;
        this.creds = creds;
        this.doc = new GoogleSpreadsheet(this.docKey);
    }

    /**
     *
     * @param {Number} num
     * @return {string}
     */
    colToAlphabet(num)
    {
        let char = (10 + (num - 1) % 26).toString(36).toUpperCase()
        return num > 0 ? this.colToAlphabet(parseInt((num - 1) / 26)) + char: '';
    }

    /**
     *
     * @param {String} formulaFormat
     * @param {Array} cells
     * @param {Number} currentRow
     * @param {Number} currentCol
     * @return {string}
     */
    formatFormula(formulaFormat, cells, currentRow, currentCol)
    {
        let arr = cells.map(function (cell){
            let col = cell.col === 'this' ? currentCol: cell.col;
            let row = cell.row === 'this' ? currentCol: cell.row;
            return (col > 0 ? this.colToAlphabet(col): '') + (row > 0 ? row: '');
        }.bind(this));

        return formulaFormat.replace(/%(\d+)/g, function(_,m) {
            return arr[--m];
        });
    }

    /**
     *
     * @param {SpreadsheetWorksheet} sheet
     * @param {Array} values
     * @param {Object} options
     * @return {Promise}
     */
    updateGoogleSheetsData(sheet, values, options) {
        let that = this;
        return new Promise((resolve, reject) => {
            let that = this;
            let rowCount = Math.max(options.minRow, values.length) + options.margin;
            let colCount = Math.max(options.minCol, values.reduce((a, b) => Math.max(a, b.length))) + options.margin;

            // apply options
            if(!options.resize){
                rowCount = sheet.rowCount;
                colCount = sheet.colCount;
            }

            // make sure we have enough cells for all data
            sheet.resize({rowCount: rowCount, colCount: colCount}, (err) => {
                if(err) return reject(err);

                sheet.getCells({
                    'min-row':      1,
                    'max-row':      rowCount,
                    'return-empty': true
                }, function (err, cells) {
                    if(err) return reject(err);

                    // store the cells in array[row][column] format for ease of access
                    let myCells = [];
                    let updatedCells = [];

                    cells.forEach(cell => {
                        if (!myCells[cell.row - 1]) myCells[cell.row - 1] = [];
                        myCells[cell.row - 1].push(cell);

                        // clear all the cells's value
                        if(options.clear) cell.value = '';
                    });

                    // update the cell
                    values.forEach((list, i) => {
                        list.forEach((value, j) => {
                            if(myCells[i] && myCells[i][j]) {
                                if (Number.isInteger(value)) {
                                    myCells[i][j].numericValue = value;
                                } else if (typeof value === 'object') {
                                    let formula = that.formatFormula(value.formula, value.cells, i, j);
                                    myCells[i][j].formula = formula;
                                } else {
                                    myCells[i][j].value = value;
                                }

                                updatedCells.push(myCells[i][j]);
                            }
                        });
                    });

                    // update the cells
                    sheet.bulkUpdateCells(options.clear ? cells: updatedCells, (err) => {
                        resolve();
                    });
                });
            });
        });
    }

    /**
     *
     * @param {String} sheetName
     * @param {Array} values
     * @param {Object} options
     * @return {Promise}
     */
    updateGoogleSheets(sheetName, values, options) {
        options = Object.assign({
            minRow: 20,
            minCol: 10,
            margin: 2,
            resize: true,
            clear: true
        }, options);

        return new Promise((resolve, reject) => {
            // the values array must be 2 dimensional
            if(values.length > 0 && !Array.isArray(values[0])) return reject(new Error('values must be 2 dimensional array.'))

            // connect to the service account
            this.doc.useServiceAccountAuth(this.creds, function (err) {
                if(err) return reject(err);

                // open the document
                this.doc.getInfo(function (err, info) {
                    if(err) return reject(err);

                    // search for existing worksheet
                    let sheet = info.worksheets.filter(x => {
                        return x.title === sheetName;
                    }).shift();

                    if (!sheet) { // create worksheet if not exist
                        this.doc.addWorksheet({
                            title: sheetName,
                        }, function (err, sheet) {
                            if(err) return reject(err);

                            // update array to the sheet
                            this.updateGoogleSheetsData(sheet, values, options).then(() => {
                                resolve();

                            }).catch((err) => {
                                reject(err);

                            });
                        }.bind(this));
                    } else {
                        // update array to the sheet
                        this.updateGoogleSheetsData(sheet, values, options).then(() => {
                            resolve();

                        }).catch((err) => {
                            reject(err);

                        });
                    }
                }.bind(this));
            }.bind(this));
        });
    }
}

module.exports = ArrayToGoogleSheets;

