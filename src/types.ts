export const GOOGLE_SPREADSHEETS_URL = "https://sheets.googleapis.com/v4/spreadsheets";

export type IUpdateBaseOptions = {
    valueInputOption?: "RAW" | "USER_ENTERED",
};

export type IUpdateOptions = IUpdateBaseOptions & {
    minRow: number,
    minColumn: number,
    margin: number,
    fitToSize: boolean,
    clearAllValues: boolean,
};

/** @internal */
export type IUpdateResponse = {
    updatedRows: number;
    updatedColumns: number;
    updatedCells: number;
};

/** @internal */
export type IFormulaCells = Array<{row?: number | string, col?: number | string}>;
/** @internal */
export type IFormula = {formula: string, cells: IFormulaCells};
/** @internal */
export type ICell = number | string | IFormula;
/** @internal */
export type IRow = ICell[];
/** @internal */
export type IValues = IRow[];
/** @internal */
export type INormalizedCell = number | string;
/** @internal */
export type INormalizedRow = INormalizedCell[];
/** @internal */
export type INormalizedValues = INormalizedRow[];

/** @internal */
export type ISpreadsheetProperties = {
    title: string
    locale: string,
    autoRecalc: string,
    timeZone: string,
    defaultFormat: {
        background: {red: number, green: number, blue: number},
        padding: {top: number, right: number, bottom: number, left: number},
        verticalAlignment: string,
        wrapStrategy: string,
        textFormat: {
            foregroundColor: any,
            fontFamily: string,
            fontSize: number,
            bold: boolean,
            italic: boolean,
            strikethrough: boolean,
            underline: boolean,
        },
    },
    spreadsheetTheme: {
        primaryFontFamily: string,
        themeColors: Array<{
            colorType: string,
            color: {
                rgbColor: {
                    red: number,
                    green: number,
                    blue: number,
                },
            },
        }>,
    },
};
