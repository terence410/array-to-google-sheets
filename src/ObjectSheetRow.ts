import {Sheet} from "./Sheet";
import {INormalizedRow, INormalizedValues} from "./types";

const sheetSymbol = Symbol("sheet");

export class ObjectSheetRow<T extends object> {
  constructor(sheet: Sheet,
              public index: number,
              public headerTypes: Array<{ name: string, type: string }>,
              public rawRowValues: INormalizedRow,
  ) {

    // add values into object
    for (let i = 0; i < this.headerTypes.length; i++) {
      const {name, type} = this.headerTypes[i];
      const castedValue = this._castValue(type, rawRowValues[i]);

      if (type !== "ignore") {
        (this as any)[name] = castedValue;
      }
    }

    Object.defineProperty(this, sheetSymbol, {
      enumerable: false,
      value: sheet,
    });
  }

  public toObject(): T {
    const object: any = {};
    for (const header of this.headerTypes) {
      if (header.type !== "ignore") {
        object[header.name] = (this as any)[header.name];
      }
    }
    return object;
  }

  public async save() {
    const sheet = this._getSheet();
    const values: INormalizedRow = [];

    for (let i = 0; i < this.headerTypes.length; i++ ) {
      const {name, type} = this.headerTypes[i];
      const normalizedValue = this._normalizeValue(type, (this as any)[name]);

      if (type === "ignore") {
        values.push(this.rawRowValues[i]);
      } else {
        values.push(normalizedValue);
      }
    }

    await sheet.updateRow(this.index + 1, values, {valueInputOption: "RAW"});
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
