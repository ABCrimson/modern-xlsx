import { columnToLetter, decodeCellRef } from './cell-ref.js';
import { StyleBuilder } from './style-builder.js';
import type {
  AutoFilterData,
  CalcChainEntryData,
  CellData,
  CellType,
  ColumnInfo,
  CommentData,
  DataValidationData,
  DateSystem,
  DefinedNameData,
  DocPropertiesData,
  FrozenPane,
  HyperlinkData,
  PageMarginsData,
  PageSetupData,
  RowData,
  SheetData,
  SheetProtectionData,
  StylesData,
  ThemeColorsData,
  WorkbookData,
  WorkbookViewData,
} from './types.js';
import { ensureInitialized, wasmWrite } from './wasm-loader.js';

// ---------------------------------------------------------------------------
// Binary search helpers for sorted rows
// ---------------------------------------------------------------------------

function findRowBinary(rows: readonly RowData[], index: number): RowData | undefined {
  let lo = 0;
  let hi = rows.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    const row = rows[mid];
    if (!row) return undefined;
    if (row.index === index) return row;
    if (row.index < index) lo = mid + 1;
    else hi = mid - 1;
  }
  return undefined;
}

function binaryInsertIndex(rows: readonly RowData[], index: number): number {
  let lo = 0;
  let hi = rows.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if ((rows[mid]?.index ?? 0) <= index) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

// ---------------------------------------------------------------------------
// Workbook
// ---------------------------------------------------------------------------

export class Workbook {
  private readonly data: WorkbookData;

  constructor(data?: Partial<WorkbookData>) {
    this.data = {
      sheets: data?.sheets ?? [],
      dateSystem: data?.dateSystem ?? 'date1900',
      styles: data?.styles ?? defaultStyles(),
    };
    if (data?.definedNames) this.data.definedNames = data.definedNames;
    if (data?.sharedStrings) this.data.sharedStrings = data.sharedStrings;
    if (data?.docProperties) this.data.docProperties = data.docProperties;
    if (data?.themeColors) this.data.themeColors = data.themeColors;
    if (data?.calcChain) this.data.calcChain = data.calcChain;
    if (data?.workbookViews) this.data.workbookViews = data.workbookViews;
    if (data?.preservedEntries) this.data.preservedEntries = data.preservedEntries;
  }

  get sheetNames(): string[] {
    return this.data.sheets.map((s) => s.name);
  }

  get sheetCount(): number {
    return this.data.sheets.length;
  }

  get dateSystem(): DateSystem {
    return this.data.dateSystem;
  }

  get styles(): StylesData {
    return this.data.styles;
  }

  createStyle(): StyleBuilder {
    return new StyleBuilder();
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

  removeSheet(nameOrIndex: string | number): boolean {
    const idx =
      typeof nameOrIndex === 'string'
        ? this.data.sheets.findIndex((s) => s.name === nameOrIndex)
        : nameOrIndex;
    if (idx < 0 || idx >= this.data.sheets.length) return false;
    this.data.sheets.splice(idx, 1);
    return true;
  }

  // --- Named ranges ---

  get namedRanges(): readonly DefinedNameData[] {
    return this.data.definedNames ?? [];
  }

  addNamedRange(name: string, value: string, sheetId?: number): void {
    if (!this.data.definedNames) {
      this.data.definedNames = [];
    }
    this.data.definedNames.push({
      name,
      value,
      sheetId: sheetId ?? null,
    });
  }

  getNamedRange(name: string): DefinedNameData | undefined {
    return this.data.definedNames?.find((d) => d.name === name);
  }

  removeNamedRange(name: string): boolean {
    if (!this.data.definedNames) return false;
    const idx = this.data.definedNames.findIndex((d) => d.name === name);
    if (idx === -1) return false;
    this.data.definedNames.splice(idx, 1);
    return true;
  }

  // --- Document properties ---

  get docProperties(): DocPropertiesData | null {
    return this.data.docProperties ?? null;
  }

  set docProperties(props: DocPropertiesData | null) {
    this.data.docProperties = props ?? undefined;
  }

  // --- Theme colors ---

  get themeColors(): ThemeColorsData | null {
    return this.data.themeColors ?? null;
  }

  // --- Calc chain ---

  get calcChain(): readonly CalcChainEntryData[] {
    return this.data.calcChain ?? [];
  }

  // --- Workbook views ---

  get workbookViews(): readonly WorkbookViewData[] {
    return this.data.workbookViews ?? [];
  }

  set workbookViews(views: WorkbookViewData[]) {
    this.data.workbookViews = views;
  }

  // --- Serialization ---

  async toBuffer(): Promise<Uint8Array> {
    ensureInitialized();
    return wasmWrite(this.data);
  }

  /**
   * Write the workbook to a file on disk.
   * Only available in Node.js, Bun, and Deno environments.
   */
  async toFile(path: string): Promise<void> {
    const buffer = await this.toBuffer();
    const { writeFile } = await import('node:fs/promises');
    await writeFile(path, buffer);
  }

  toJSON(): WorkbookData {
    return this.data;
  }
}

// ---------------------------------------------------------------------------
// Worksheet
// ---------------------------------------------------------------------------

export class Worksheet {
  private readonly data: SheetData;

  constructor(data: SheetData) {
    this.data = data;
  }

  get name(): string {
    return this.data.name;
  }

  set name(value: string) {
    this.data.name = value;
  }

  get rows(): readonly RowData[] {
    return this.data.worksheet.rows;
  }

  // --- Columns ---

  get columns(): readonly ColumnInfo[] {
    return this.data.worksheet.columns;
  }

  set columns(cols: ColumnInfo[]) {
    this.data.worksheet.columns = cols;
  }

  setColumnWidth(col: number, width: number): void {
    const existing = this.data.worksheet.columns.find((c) => c.min === col && c.max === col);
    if (existing) {
      existing.width = width;
      existing.customWidth = true;
    } else {
      this.data.worksheet.columns.push({
        min: col,
        max: col,
        width,
        hidden: false,
        customWidth: true,
      });
    }
  }

  // --- Merge cells ---

  get mergeCells(): readonly string[] {
    return this.data.worksheet.mergeCells;
  }

  addMergeCell(range: string): void {
    if (!this.data.worksheet.mergeCells.includes(range)) {
      this.data.worksheet.mergeCells.push(range);
    }
  }

  removeMergeCell(range: string): boolean {
    const idx = this.data.worksheet.mergeCells.indexOf(range);
    if (idx === -1) return false;
    this.data.worksheet.mergeCells.splice(idx, 1);
    return true;
  }

  // --- Auto filter ---

  get autoFilter(): AutoFilterData | null {
    return this.data.worksheet.autoFilter;
  }

  set autoFilter(filter: AutoFilterData | string | null) {
    if (filter === null) {
      this.data.worksheet.autoFilter = null;
    } else if (typeof filter === 'string') {
      this.data.worksheet.autoFilter = { range: filter };
    } else {
      this.data.worksheet.autoFilter = filter;
    }
  }

  // --- Hyperlinks ---

  get hyperlinks(): readonly HyperlinkData[] {
    return this.data.worksheet.hyperlinks ?? [];
  }

  addHyperlink(
    cellRef: string,
    location: string,
    opts?: { display?: string; tooltip?: string },
  ): void {
    if (!this.data.worksheet.hyperlinks) {
      this.data.worksheet.hyperlinks = [];
    }
    this.data.worksheet.hyperlinks.push({
      cellRef,
      location,
      display: opts?.display ?? null,
      tooltip: opts?.tooltip ?? null,
    });
  }

  removeHyperlink(cellRef: string): boolean {
    if (!this.data.worksheet.hyperlinks) return false;
    const idx = this.data.worksheet.hyperlinks.findIndex((h) => h.cellRef === cellRef);
    if (idx === -1) return false;
    this.data.worksheet.hyperlinks.splice(idx, 1);
    return true;
  }

  // --- Page setup ---

  get pageSetup(): PageSetupData | null {
    return this.data.worksheet.pageSetup ?? null;
  }

  set pageSetup(setup: PageSetupData | null) {
    this.data.worksheet.pageSetup = setup ?? undefined;
  }

  // --- Sheet protection ---

  get sheetProtection(): SheetProtectionData | null {
    return this.data.worksheet.sheetProtection ?? null;
  }

  set sheetProtection(protection: SheetProtectionData | null) {
    this.data.worksheet.sheetProtection = protection ?? undefined;
  }

  // --- Frozen pane ---

  get frozenPane(): FrozenPane | null {
    return this.data.worksheet.frozenPane;
  }

  set frozenPane(pane: FrozenPane | null) {
    this.data.worksheet.frozenPane = pane;
  }

  // --- Data validations ---

  get validations(): readonly DataValidationData[] {
    return this.data.worksheet.dataValidations ?? [];
  }

  addValidation(ref: string, rule: Omit<DataValidationData, 'sqref'>): void {
    if (!this.data.worksheet.dataValidations) {
      this.data.worksheet.dataValidations = [];
    }
    this.data.worksheet.dataValidations.push({ sqref: ref, ...rule });
  }

  removeValidation(ref: string): boolean {
    if (!this.data.worksheet.dataValidations) return false;
    const idx = this.data.worksheet.dataValidations.findIndex((v) => v.sqref === ref);
    if (idx === -1) return false;
    this.data.worksheet.dataValidations.splice(idx, 1);
    return true;
  }

  // --- Comments ---

  get comments(): readonly CommentData[] {
    return this.data.worksheet.comments ?? [];
  }

  addComment(cellRef: string, author: string, text: string): void {
    if (!this.data.worksheet.comments) {
      this.data.worksheet.comments = [];
    }
    this.data.worksheet.comments.push({ cellRef, author, text });
  }

  removeComment(cellRef: string): boolean {
    if (!this.data.worksheet.comments) return false;
    const idx = this.data.worksheet.comments.findIndex((c) => c.cellRef === cellRef);
    if (idx === -1) return false;
    this.data.worksheet.comments.splice(idx, 1);
    return true;
  }

  // --- Page margins ---

  get pageMargins(): PageMarginsData | null {
    return this.data.worksheet.pageSetup?.margins ?? null;
  }

  set pageMargins(margins: PageMarginsData | null) {
    if (!this.data.worksheet.pageSetup) {
      this.data.worksheet.pageSetup = {};
    }
    this.data.worksheet.pageSetup.margins = margins;
  }

  // --- Cell access ---

  cell(ref: string): Cell {
    const decoded = decodeCellRef(ref);
    const rowIndex = decoded.row + 1; // decodeCellRef returns 0-based row
    const normalRef = `${columnToLetter(decoded.col)}${rowIndex}`;

    let row = findRowBinary(this.data.worksheet.rows, rowIndex);
    if (!row) {
      row = { index: rowIndex, cells: [], height: null, hidden: false };
      const insertAt = binaryInsertIndex(this.data.worksheet.rows, rowIndex);
      this.data.worksheet.rows.splice(insertAt, 0, row);
    }

    let cellData = row.cells.find((c) => c.reference === normalRef);
    if (!cellData) {
      cellData = {
        reference: normalRef,
        cellType: 'number',
        value: null,
        formula: null,
        styleIndex: null,
      };
      row.cells.push(cellData);
    }

    return new Cell(cellData);
  }

  // --- Row utilities ---

  setRowHeight(rowIndex: number, height: number): void {
    const row = this.ensureRow(rowIndex);
    row.height = height;
  }

  setRowHidden(rowIndex: number, hidden: boolean): void {
    const row = this.ensureRow(rowIndex);
    row.hidden = hidden;
  }

  /** Find or create a row at the given 1-based index. */
  private ensureRow(rowIndex: number): RowData {
    const existing = findRowBinary(this.data.worksheet.rows, rowIndex);
    if (existing) return existing;
    const newRow: RowData = { index: rowIndex, cells: [], height: null, hidden: false };
    const insertAt = binaryInsertIndex(this.data.worksheet.rows, rowIndex);
    this.data.worksheet.rows.splice(insertAt, 0, newRow);
    return newRow;
  }
}

// ---------------------------------------------------------------------------
// Cell
// ---------------------------------------------------------------------------

export class Cell {
  private readonly data: CellData;

  constructor(data: CellData) {
    this.data = data;
  }

  get reference(): string {
    return this.data.reference;
  }

  get type(): CellType {
    return this.data.cellType;
  }

  get styleIndex(): number | null {
    return this.data.styleIndex ?? null;
  }

  set styleIndex(value: number | null) {
    this.data.styleIndex = value ?? null;
  }

  get value(): string | number | boolean | null {
    if (this.data.value == null) return null;
    switch (this.data.cellType) {
      case 'number':
        return Number.parseFloat(this.data.value);
      case 'boolean':
        return this.data.value === '1';
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
      // Don't change cellType if a formula is set (value is the cached result).
      if (!this.data.formula) {
        this.data.cellType = 'sharedString';
      }
      this.data.value = val;
    } else if (typeof val === 'number') {
      if (!this.data.formula) {
        this.data.cellType = 'number';
      }
      this.data.value = val.toString();
    } else if (typeof val === 'boolean') {
      if (!this.data.formula) {
        this.data.cellType = 'boolean';
      }
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

// ---------------------------------------------------------------------------
// Default styles (Excel minimum)
// ---------------------------------------------------------------------------

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
        color: null,
      },
    ],
    fills: [
      { patternType: 'none', fgColor: null, bgColor: null },
      { patternType: 'gray125', fgColor: null, bgColor: null },
    ],
    borders: [{ left: null, right: null, top: null, bottom: null }],
    cellXfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }],
  } satisfies StylesData;
}
