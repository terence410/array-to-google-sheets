import {Sheet} from "./Sheet";
import {INormalizedRow, INormalizedValues, IUpdateCells} from "./types";

const sheetSymbol = Symbol("sheet");

export class ObjectSheetRow<T extends object = object> {
  private _cachedValues: any = {};
  private _changes: Set<string> = new Set();

  constructor(sheet: Sheet,
              private _rawRowValues: INormalizedRow,
              private _index: number,
              private _headerTypes: Array<{ name: string, type: string }>,
  ) {

    // add values into object
    for (let i = 0; i < this._headerTypes.length; i++) {
      const {name, type} = this._headerTypes[i];
      const castedValue = this._castValue(type, _rawRowValues[i]);

      if (type !== "ignore") {
        this._cachedValues[name] = castedValue;

        Object.defineProperty(this, name, {
          get() {
            return this._cachedValues[name];
          },
          set(value: any) {
            this._changes.add(name);
            this._cachedValues[name] = value;
          },
        });
      }
    }

    Object.defineProperty(this, sheetSymbol, {
      enumerable: false,
      value: sheet,
    });

  }

  public toObject(): T {
    const object: any = {};
    for (const header of this._headerTypes) {
      if (header.type !== "ignore") {
        object[header.name] = this._cachedValues[header.name];
      }
    }
    return object;
  }

  public async save() {
    const sheet = this._getSheet();
    const values: INormalizedRow = [];

    const cells: IUpdateCells = [];
    for (let i = 0; i < this._headerTypes.length; i++) {
      const {name, type} = this._headerTypes[i];
      if (this._changes.has(name)) {
        const normalizedValue = this._normalizeValue(type, this._cachedValues[name]);
        cells.push({rowIndex: this._index + 1, columnIndex: i, cell: normalizedValue});
      }
    }

    this._changes.clear();
    await sheet.updateCells(cells, {valueInputOption: "RAW"});
  }

  private _getSheet(): Sheet {
    return (this as any)[sheetSymbol] as Sheet;
  }

  private _castValue(type: string, value: any) {
    switch (type) {
      case "string":
        return String(value || "");

      case "number":
        return Number(value);

      case "boolean":
        return Boolean(value);

      case "Date":
        return new Date(value);

      case "number[]":
        return value ? String(value).split(",").map(x => x.trim()).map(Number) : [];

      case "string[]":
        return value ? String(value).split(",").map(x => x.trim()) : [];

      default:
        return "";
    }
  }

  private _normalizeValue(type: string, value: any) {
    switch (type) {
      case "number":
        return Number.isNaN(value) ? "NaN" : value;

      case "string":
      case "boolean":
        return value;

      case "Date":
        try {
          return (value as Date).toISOString();
        } catch (err) {
          return "Invalid Date";
        }

      case "number[]":
      case "string[]":
        return (value as number[]).join(", ");

      default:
        return "";
    }
  }

}
