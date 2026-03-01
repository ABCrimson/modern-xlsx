import type {
  CellData,
  DateSystem,
  RowData,
  SheetData,
  StylesData,
  WorkbookData,
} from './types.js';
import { ensureInitialized, wasmWrite } from './wasm-loader.js';

export class Workbook {
  private data: WorkbookData;

  constructor(data?: Partial<WorkbookData>) {
    this.data = {
      sheets: data?.sheets ?? [],
      dateSystem: data?.dateSystem ?? 'date1900',
      styles: data?.styles ?? defaultStyles(),
    };
  }

  get sheetNames(): string[] {
    return this.data.sheets.map((s) => s.name);
  }

  get dateSystem(): DateSystem {
    return this.data.dateSystem;
  }

  getSheet(name: string): Worksheet | undefined {
    const sheet = this.data.sheets.find((s) => s.name === name);
    return sheet ? new Worksheet(sheet) : undefined;
  }

  getSheetByIndex(index: number): Worksheet | undefined {
    const sheet = this.data.sheets[index];
    return sheet ? new Worksheet(sheet) : undefined;
  }

  addSheet(name: string): Worksheet {
    const sheetData: SheetData = {
      name,
      worksheet: {
        dimension: null,
        rows: [],
        mergeCells: [],
        autoFilter: null,
        frozenPane: null,
        columns: [],
      },
    };
    this.data.sheets.push(sheetData);
    return new Worksheet(sheetData);
  }

  async toBuffer(): Promise<Uint8Array> {
    ensureInitialized();
    return wasmWrite(this.data);
  }

  toJSON(): WorkbookData {
    return this.data;
  }
}

export class Worksheet {
  private data: SheetData;

  constructor(data: SheetData) {
    this.data = data;
  }

  get name(): string {
    return this.data.name;
  }

  get rows(): RowData[] {
    return this.data.worksheet.rows;
  }

  get mergeCells(): string[] {
    return this.data.worksheet.mergeCells;
  }

  cell(ref: string): Cell {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (!match || match[2] == null) throw new Error(`Invalid cell reference: ${ref}`);
    const rowIndex = parseInt(match[2], 10);

    let row = this.data.worksheet.rows.find((r) => r.index === rowIndex);
    if (!row) {
      row = { index: rowIndex, cells: [], hidden: false };
      this.data.worksheet.rows.push(row);
      this.data.worksheet.rows.sort((a, b) => a.index - b.index);
    }

    let cellData = row.cells.find((c) => c.reference === ref);
    if (!cellData) {
      cellData = { reference: ref, cellType: 'number', value: null, formula: null };
      row.cells.push(cellData);
    }

    return new Cell(cellData);
  }
}

export class Cell {
  private data: CellData;

  constructor(data: CellData) {
    this.data = data;
  }

  get reference(): string {
    return this.data.reference;
  }

  get type(): string {
    return this.data.cellType;
  }

  get value(): string | number | boolean | null {
    if (this.data.value == null) return null;

    switch (this.data.cellType) {
      case 'number':
        return parseFloat(this.data.value);
      case 'boolean':
        return this.data.value === '1';
      case 'sharedString':
        return this.data.value;
      default:
        return this.data.value;
    }
  }

  set value(val: string | number | boolean | null) {
    if (val === null) {
      this.data.value = null;
      return;
    }

    if (typeof val === 'string') {
      this.data.cellType = 'sharedString';
      this.data.value = val;
    } else if (typeof val === 'number') {
      this.data.cellType = 'number';
      this.data.value = val.toString();
    } else if (typeof val === 'boolean') {
      this.data.cellType = 'boolean';
      this.data.value = val ? '1' : '0';
    }
  }

  get formula(): string | null {
    return this.data.formula ?? null;
  }

  set formula(f: string | null) {
    this.data.formula = f;
    if (f !== null) {
      this.data.cellType = 'formulaStr';
    }
  }
}

function defaultStyles(): StylesData {
  return {
    numFmts: [],
    fonts: [
      {
        name: 'Aptos',
        size: 11,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
      },
    ],
    fills: [{ patternType: 'none' }, { patternType: 'gray125' }],
    borders: [{}],
    cellXfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }],
  };
}
