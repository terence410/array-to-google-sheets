export type IUpdateOptions = {
    minRow: number,
    minColumn: number,
    margin: number,
    fitToSize: boolean,
    clearAllValues: boolean,
};

/** @internal */
export type IFormulaCells = Array<{row?: number | string, col?: number | string}>;
/** @internal */
export type IFormula = {formula: string, cells: IFormulaCells};
/** @internal */
export type IValues = Array<Array<number | string | IFormula>>;
/** @internal */
export type IPlainValues = Array<Array<number | string>>;
/** @internal */
export type IUpdateResponse = {
    updatedRows: number;
    updatedColumns: number;
    updatedCells: number;
};
