import {ObjectSheetRow} from "./ObjectSheetRow";
import {Sheet} from "./Sheet";
import {INormalizedValues, IValues} from "./types";

const sheetSymbol = Symbol("sheet");

export class ObjectSheet<T extends object = object> {
  public readonly headers: string[];
  public readonly rawHeaders: string[];
  public readonly rawValues: INormalizedValues;

  private _headerTypes: Array<{name: string, type: string}> = [];
  private _cachedObjects: any[] = [];

  constructor(sheet: Sheet, rawValues: INormalizedValues) {
    this.rawValues = rawValues;

    // validate it
    this._validateValues();

    // assign other values
    this.rawHeaders = rawValues[0].map(x => x.toString());
    this.headers = this._headerTypes.filter(x => x.type !== "ignore").map(x => x.name);

    // assign an internal attribute
    Object.defineProperty(this, sheetSymbol, {
      enumerable: false,
      value: sheet,
    });
  }

  public [Symbol.iterator](): Iterator<T & ObjectSheetRow> {
    let step = 0;
    const next = () => {
      if (step < this.size) {
        const value = this.get(step);
        step++;
        return {value, done: false} ;
      }
      return {value: null, done: true} ;
    };

    return {next} as any;
  }

  public get size() {
    return this.rawValues.length - 1;
  }

  public getInterface() {
    const types: any = [];
    for (const header of this._headerTypes) {
      const {name, type} = header;
      if (type !== "ignore") {
        types.push(`${name}: ${type};`);
      }
    }

    return `interface IObject {${types.join(" ")}}`;
  }

  public entries(): T[] {
    const objects: any[] = [];
    for (let i = 1; i < this.rawValues.length; i++) {
      const objectSheetRow = this.get(i - 1);
      objects.push(objectSheetRow.toObject());
    }

    return objects;
  }

  public get(index: number): T & ObjectSheetRow<T>  {
    if (index < this.rawValues.length - 1) {
      if (!this._cachedObjects[index]) {
        const rawRowValues = this.rawValues[index + 1];
        this._cachedObjects[index] = new ObjectSheetRow<T>(this._getSheet(), rawRowValues, index, this._headerTypes);
      }

      return this._cachedObjects[index] as T & ObjectSheetRow<T>;
    }
    
    return undefined as any;
  }

  private _validateValues() {
    if (this.rawValues.length <= 1) {
      throw new Error("ObjectSheet Error. Please make sure the sheet has two rows or above.");
    }

    // validate and update header map
    const rawHeaders = this.rawValues[0].map(x => x.toString());
    const headerRegex = /^([a-zA-Z0-9_]*)(\/(string|number|boolean|date|number\[\]|string\[\]|ignore))?$/;
    for (const rawHeader of rawHeaders) {
      const matches = rawHeader.match(headerRegex);
      if (matches) {
        const name = matches[1];
        const type = (matches[3] === "date" ? "Date" : matches[3]) || "string";

        // check for errors
        if (this._headerTypes.find(x => x.name === name) && type !== "ignore") {
          throw new Error(`Object Sheet Error. Header with the name "${name}" already exist`);
        }

        // check for protected values
        if (["save", "toObject"].includes(name)) {
          throw new Error(`Object Sheet Error. Header with the name "${name}" is not allowed`);
        }

        this._headerTypes.push({name, type});

      } else {
        throw new Error(`Object Sheet Error. Header "${rawHeader}" does not match with the regex: ${headerRegex}`);
      }
    }

    // check header format
  }

  private _getSheet(): Sheet {
    return (this as any)[sheetSymbol] as Sheet;
  }

}
