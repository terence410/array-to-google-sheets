import {ObjectSheetRow} from "./ObjectSheetRow";
import {Sheet} from "./Sheet";
import {INormalizedValues, IValues} from "./types";

const sheetSymbol = Symbol("sheet");

export class ObjectSheet<T extends any> {
  public rawHeaders: string[];
  public headers: Array<{name: string, type: string}> = [];
  public objects: any[] = [];

  constructor(sheet: Sheet, public rawValues: INormalizedValues) {
    this._validateValues();

    // assign values for this object
    this.rawHeaders = rawValues[0].map(x => x.toString());

    Object.defineProperty(this, sheetSymbol, {
      enumerable: false,
      value: sheet,
    });
  }

  public get length() {
    return this.rawValues.length - 1;
  }

  public toObjects(): T[] {
    const objects: any[] = [];
    for (let i = 1; i < this.rawValues.length; i++) {
      const objectSheetRow = this.get(i - 1);
      objects.push(objectSheetRow.toObject());
    }

    return objects;
  }

  public getType() {
    const types: any = [];
    for (const header of this.headers) {
      const {name, type} = header;
      if (type !== "ignore") {
        types.push(`${name}: ${type}`);
      }
    }

    return `type IObject = {${types.join("; ")}};`;
  }

  public get(index: number): T & ObjectSheetRow<T>  {
    if (index < this.rawValues.length - 1) {
      if (!this.objects[index]) {
        const rawRowValues = this.rawValues[index + 1];
        this.objects[index] = new ObjectSheetRow<T>(this._getSheet(), index, this.headers, rawRowValues);
      }

      return this.objects[index] as T & ObjectSheetRow<T>;
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

        if (this.headers.find(x => x.name === name)) {
          throw new Error(`Object Sheet Error. Header with the name "${name}" already exist`);
        }
        this.headers.push({name, type});

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
