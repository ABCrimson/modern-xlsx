import type { DrawBarcodeOptions, ImageAnchor } from './barcode/index.js';
import { generateBarcode, generateDrawingRels, generateDrawingXml } from './barcode/index.js';
import { columnToLetter, decodeCellRef } from './cell-ref.js';
import { ChartBuilder } from './chart-builder.js';
import { isDateFormatCode, serialToDate } from './dates.js';
import { COMMENT_NOT_FOUND, INVALID_ARGUMENT, ModernXlsxError, SHEET_NOT_FOUND } from './errors.js';
import { getBuiltinFormat } from './format-cell.js';
import { PivotTableBuilder } from './pivot-builder.js';
import { StyleBuilder } from './style-builder.js';
import type {
  AutoFilterData,
  CalcChainEntryData,
  CellData,
  CellType,
  ChartType,
  ColumnInfo,
  CommentData,
  DataValidationData,
  DateSystem,
  DefinedNameData,
  DocPropertiesData,
  FrozenPane,
  HeaderFooterData,
  HyperlinkData,
  OutlinePropertiesData,
  PageBreakData,
  PageBreaksData,
  PageMarginsData,
  PageSetupData,
  PaneSelectionData,
  PivotCacheDefinitionData,
  PivotCacheRecordsData,
  PivotTableData,
  RepairResult,
  RichTextRun,
  RowData,
  SheetData,
  SheetProtectionData,
  SheetState,
  SheetViewData,
  SlicerCacheData,
  SlicerData,
  SparklineGroupData,
  SplitPaneData,
  StylesData,
  TableDefinitionData,
  ThemeColorsData,
  ThreadedCommentData,
  TimelineCacheData,
  TimelineData,
  ValidationReport,
  ViewMode,
  WorkbookData,
  WorkbookProtectionData,
  WorkbookViewData,
  WorksheetChartData,
  WriteOptions,
} from './types.js';
import { validateChartData } from './validate-chart.js';
import {
  ensureInitialized,
  wasmRepair,
  wasmValidate,
  wasmWrite,
  wasmWriteWithPassword,
} from './wasm-loader.js';

const TEXT_ENCODER = new TextEncoder();

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
// Rich text SST resolution helper
// ---------------------------------------------------------------------------

/**
 * Build a reverse map from plain text to rich text runs using the SST.
 * Returns `null` if the SST has no rich text entries.
 */
function buildRichTextMap(
  sst:
    | { readonly strings: string[]; readonly richRuns?: (RichTextRun[] | null)[] }
    | undefined
    | null,
): Map<string, RichTextRun[]> | null {
  if (!sst?.richRuns?.length) return null;
  const map = new Map<string, RichTextRun[]>();
  for (let i = 0; i < sst.richRuns.length; i++) {
    const runs = sst.richRuns[i];
    const key = sst.strings[i];
    if (runs && key !== undefined) {
      map.set(key, runs);
    }
  }
  return map.size > 0 ? map : null;
}

// ---------------------------------------------------------------------------
// Workbook
// ---------------------------------------------------------------------------

/**
 * Represents an Excel workbook containing sheets, styles, and metadata.
 *
 * This is the root object for reading, writing, and manipulating XLSX files.
 * Obtain an instance via {@link readBuffer}, {@link readFile}, or construct
 * a new empty workbook with `new Workbook()`.
 *
 * @example
 * ```ts
 * const wb = new Workbook();
 * const ws = wb.addSheet('Data');
 * ws.cell('A1').value = 'Hello';
 * const bytes = await wb.toBuffer();
 * ```
 */
export class Workbook {
  private readonly data: WorkbookData;
  private readonly imageAnchors = new Map<string, { anchor: ImageAnchor; imageIndex: number }[]>();
  private readonly imageRels = new Map<string, { rId: string; target: string }[]>();
  private _sheetIndex = new Map<string, number>();

  /**
   * Create a new workbook, optionally seeded with existing data from a read operation.
   *
   * @param data - Partial workbook data to seed the workbook with. Omit for an empty workbook.
   *
   * @example
   * ```ts
   * const wb = new Workbook();
   * wb.addSheet('Sheet1');
   * ```
   */
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
    if (data?.protection) this.data.protection = data.protection;
    if (data?.persons) this.data.persons = data.persons;
    if (data?.pivotCaches) this.data.pivotCaches = data.pivotCaches;
    if (data?.pivotCacheRecords) this.data.pivotCacheRecords = data.pivotCacheRecords;
    if (data?.slicerCaches) this.data.slicerCaches = data.slicerCaches;
    if (data?.timelineCaches) this.data.timelineCaches = data.timelineCaches;
    if (data?.preservedEntries) this.data.preservedEntries = data.preservedEntries;
    this._rebuildSheetIndex();
    this._resolveRichText();
  }

  private _rebuildSheetIndex(): void {
    this._sheetIndex = new Map<string, number>();
    for (let i = 0; i < this.data.sheets.length; i++) {
      const sheet = this.data.sheets[i];
      if (sheet) this._sheetIndex.set(sheet.name, i);
    }
  }

  /**
   * Propagate rich text runs from the shared string table to individual cells.
   *
   * The WASM JSON streaming reader resolves SST indices to plain text but does
   * not populate the `richText` field on cells. This method builds a reverse
   * map (text -> rich runs) from the SST and attaches runs to matching
   * SharedString cells that don't already have rich text set.
   */
  private _resolveRichText(): void {
    const richMap = buildRichTextMap(this.data.sharedStrings);
    if (!richMap) return;

    for (const sheet of this.data.sheets) {
      for (const row of sheet.worksheet.rows) {
        for (const cell of row.cells) {
          if (cell.cellType === 'sharedString' && cell.value != null && !cell.richText) {
            const runs = richMap.get(cell.value);
            if (runs) cell.richText = runs;
          }
        }
      }
    }
  }

  /** Returns an array of all sheet names in order. */
  get sheetNames(): readonly string[] {
    return this.data.sheets.map((s) => s.name);
  }

  /** Returns the number of sheets in the workbook. */
  get sheetCount(): number {
    return this.data.sheets.length;
  }

  /** Returns the date epoch system used by this workbook (`'date1900'` or `'date1904'`). */
  get dateSystem(): DateSystem {
    return this.data.dateSystem;
  }

  /** Returns the workbook's shared styles (fonts, fills, borders, number formats, cell XFs). */
  get styles(): StylesData {
    return this.data.styles;
  }

  /**
   * Creates a new fluent style builder for constructing cell styles.
   *
   * @returns A fresh StyleBuilder instance.
   *
   * @example
   * ```ts
   * const idx = wb.createStyle()
   *   .font({ bold: true, size: 14 })
   *   .fill({ pattern: 'solid', fgColor: 'FFFF00' })
   *   .build(wb.styles);
   * ws.cell('A1').styleIndex = idx;
   * ```
   */
  createStyle(): StyleBuilder {
    return new StyleBuilder();
  }

  /**
   * Returns the worksheet with the given name, or `undefined` if not found.
   *
   * @param name - The sheet tab name to look up.
   * @returns The matching Worksheet, or `undefined`.
   *
   * @example
   * ```ts
   * const ws = wb.getSheet('Sales');
   * if (ws) console.log(ws.cell('A1').value);
   * ```
   */
  getSheet(name: string): Worksheet | undefined {
    const idx = this._sheetIndex.get(name);
    if (idx === undefined) return undefined;
    const sheet = this.data.sheets[idx];
    return sheet ? new Worksheet(sheet, this.data.styles, this.data) : undefined;
  }

  /**
   * Returns the worksheet at the given 0-based index, or `undefined` if out of range.
   *
   * @param index - Zero-based sheet index.
   * @returns The Worksheet at the given position, or `undefined`.
   *
   * @example
   * ```ts
   * const firstSheet = wb.getSheetByIndex(0);
   * ```
   */
  getSheetByIndex(index: number): Worksheet | undefined {
    const sheet = this.data.sheets[index];
    return sheet ? new Worksheet(sheet, this.data.styles, this.data) : undefined;
  }

  /**
   * Adds a new empty sheet to the workbook and returns it.
   * The name is validated per ECMA-376 rules (max 31 chars, no special characters).
   *
   * @param name - The sheet tab name. Must be unique and 1-31 characters.
   * @returns The newly created Worksheet.
   * @throws If the name is invalid or a sheet with that name already exists.
   *
   * @example
   * ```ts
   * const ws = wb.addSheet('Q1 Report');
   * ws.cell('A1').value = 'Revenue';
   * ws.cell('B1').value = 50000;
   * ```
   */
  addSheet(name: string): Worksheet {
    validateSheetName(name);
    if (this.data.sheets.some((s) => s.name === name)) {
      throw new ModernXlsxError(INVALID_ARGUMENT, `Sheet "${name}" already exists`);
    }
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
    this._rebuildSheetIndex();
    return new Worksheet(sheetData, this.data.styles, this.data);
  }

  /**
   * Removes a sheet by name or 0-based index.
   *
   * @param nameOrIndex - The sheet name or zero-based index to remove.
   * @returns `true` if the sheet was removed, `false` if not found.
   *
   * @example
   * ```ts
   * wb.removeSheet('OldSheet');
   * wb.removeSheet(2);
   * ```
   */
  removeSheet(nameOrIndex: string | number): boolean {
    const idx =
      typeof nameOrIndex === 'string'
        ? this.data.sheets.findIndex((s) => s.name === nameOrIndex)
        : nameOrIndex;
    if (idx < 0 || idx >= this.data.sheets.length) return false;
    this.data.sheets.splice(idx, 1);
    this._rebuildSheetIndex();
    return true;
  }

  /**
   * Moves a sheet from one position to another.
   *
   * @param fromIndex - Zero-based source position.
   * @param toIndex - Zero-based destination position.
   * @throws If either index is out of range.
   *
   * @example
   * ```ts
   * wb.moveSheet(2, 0); // move third sheet to first position
   * ```
   */
  moveSheet(fromIndex: number, toIndex: number): void {
    if (fromIndex < 0 || fromIndex >= this.data.sheets.length) {
      throw new ModernXlsxError(INVALID_ARGUMENT, `Invalid source index: ${fromIndex}`);
    }
    if (toIndex < 0 || toIndex >= this.data.sheets.length) {
      throw new ModernXlsxError(INVALID_ARGUMENT, `Invalid destination index: ${toIndex}`);
    }
    const [sheet] = this.data.sheets.splice(fromIndex, 1);
    if (!sheet) return;
    this.data.sheets.splice(toIndex, 0, sheet);
    this._rebuildSheetIndex();
  }

  /**
   * Clones a sheet and inserts the copy at the end (or given position).
   *
   * @param sourceIndex - Zero-based index of the sheet to clone.
   * @param newName - Tab name for the cloned sheet.
   * @param insertIndex - Optional zero-based position to insert the clone.
   * @returns The newly created Worksheet (deep copy of the source).
   * @throws If the source index is invalid or the new name already exists.
   *
   * @example
   * ```ts
   * const copy = wb.cloneSheet(0, 'Sheet1 Copy');
   * ```
   */
  cloneSheet(sourceIndex: number, newName: string, insertIndex?: number): Worksheet {
    if (sourceIndex < 0 || sourceIndex >= this.data.sheets.length) {
      throw new ModernXlsxError(INVALID_ARGUMENT, `Invalid source index: ${sourceIndex}`);
    }
    validateSheetName(newName);
    if (this.data.sheets.some((s) => s.name === newName)) {
      throw new ModernXlsxError(INVALID_ARGUMENT, `Sheet "${newName}" already exists`);
    }
    const source = this.data.sheets[sourceIndex];
    if (!source)
      throw new ModernXlsxError(INVALID_ARGUMENT, `Invalid source index: ${sourceIndex}`);
    const clone: SheetData = structuredClone(source);
    clone.name = newName;
    const idx = insertIndex ?? this.data.sheets.length;
    this.data.sheets.splice(idx, 0, clone);
    this._rebuildSheetIndex();
    return new Worksheet(clone, this.data.styles, this.data);
  }

  /**
   * Renames a sheet.
   *
   * @param nameOrIndex - The current sheet name or zero-based index.
   * @param newName - The new tab name (validated per ECMA-376 rules).
   * @throws If the sheet is not found, or the new name is invalid/duplicate.
   *
   * @example
   * ```ts
   * wb.renameSheet('Sheet1', 'January');
   * ```
   */
  renameSheet(nameOrIndex: string | number, newName: string): void {
    const idx =
      typeof nameOrIndex === 'string'
        ? this.data.sheets.findIndex((s) => s.name === nameOrIndex)
        : nameOrIndex;
    if (idx < 0 || idx >= this.data.sheets.length) {
      throw new ModernXlsxError(SHEET_NOT_FOUND, `Sheet not found: ${nameOrIndex}`);
    }
    validateSheetName(newName);
    if (this.data.sheets.some((s, i) => s.name === newName && i !== idx)) {
      throw new ModernXlsxError(INVALID_ARGUMENT, `Sheet "${newName}" already exists`);
    }
    const sheet = this.data.sheets[idx];
    if (!sheet) throw new ModernXlsxError(SHEET_NOT_FOUND, `Sheet not found: ${nameOrIndex}`);
    sheet.name = newName;
    this._rebuildSheetIndex();
  }

  /**
   * Hides a sheet. At least one sheet must remain visible.
   *
   * @param nameOrIndex - The sheet name or zero-based index to hide.
   * @throws If the sheet is not found or it is the last visible sheet.
   *
   * @example
   * ```ts
   * wb.hideSheet('Internal');
   * ```
   */
  hideSheet(nameOrIndex: string | number): void {
    const idx =
      typeof nameOrIndex === 'string'
        ? this.data.sheets.findIndex((s) => s.name === nameOrIndex)
        : nameOrIndex;
    if (idx < 0 || idx >= this.data.sheets.length) {
      throw new ModernXlsxError(SHEET_NOT_FOUND, `Sheet not found: ${nameOrIndex}`);
    }
    const target = this.data.sheets[idx];
    if (!target) throw new ModernXlsxError(SHEET_NOT_FOUND, `Sheet not found: ${nameOrIndex}`);
    if (target.state === 'hidden' || target.state === 'veryHidden') return;
    const visibleCount = this.data.sheets.filter((s) => !s.state || s.state === 'visible').length;
    if (visibleCount <= 1) {
      throw new ModernXlsxError(INVALID_ARGUMENT, 'Cannot hide the last visible sheet');
    }
    target.state = 'hidden';
  }

  /**
   * Unhides a sheet (sets state to visible).
   *
   * @param nameOrIndex - The sheet name or zero-based index to unhide.
   * @throws If the sheet is not found.
   *
   * @example
   * ```ts
   * wb.unhideSheet('Internal');
   * ```
   */
  unhideSheet(nameOrIndex: string | number): void {
    const idx =
      typeof nameOrIndex === 'string'
        ? this.data.sheets.findIndex((s) => s.name === nameOrIndex)
        : nameOrIndex;
    if (idx < 0 || idx >= this.data.sheets.length) {
      throw new ModernXlsxError(SHEET_NOT_FOUND, `Sheet not found: ${nameOrIndex}`);
    }
    const sheet = this.data.sheets[idx];
    if (!sheet) throw new ModernXlsxError(SHEET_NOT_FOUND, `Sheet not found: ${nameOrIndex}`);
    sheet.state = null;
  }

  // --- Named ranges ---

  /** Returns all defined names (named ranges) in the workbook. */
  get namedRanges(): readonly DefinedNameData[] {
    return this.data.definedNames ?? [];
  }

  /**
   * Adds a named range. If `sheetId` is provided, the name is scoped to that sheet.
   *
   * @param name - The defined name (e.g., `'TotalSales'`).
   * @param value - The range reference (e.g., `'Sheet1!$A$1:$A$10'`).
   * @param sheetId - Optional zero-based sheet index to scope the name to.
   *
   * @example
   * ```ts
   * wb.addNamedRange('Revenue', 'Sheet1!$B$2:$B$100');
   * ```
   */
  addNamedRange(name: string, value: string, sheetId?: number): void {
    this.data.definedNames ??= [];
    this.data.definedNames.push({
      name,
      value,
      sheetId: sheetId ?? null,
    });
  }

  /**
   * Returns the named range with the given name, or `undefined` if not found.
   *
   * @param name - The defined name to look up.
   * @returns The matching DefinedNameData, or `undefined`.
   *
   * @example
   * ```ts
   * const range = wb.getNamedRange('Revenue');
   * console.log(range?.value); // 'Sheet1!$B$2:$B$100'
   * ```
   */
  getNamedRange(name: string): DefinedNameData | undefined {
    return this.data.definedNames?.find((d) => d.name === name);
  }

  /**
   * Removes a named range by name.
   *
   * @param name - The defined name to remove.
   * @returns `true` if the named range was removed, `false` if not found.
   *
   * @example
   * ```ts
   * wb.removeNamedRange('OldRange');
   * ```
   */
  removeNamedRange(name: string): boolean {
    if (!this.data.definedNames) return false;
    const idx = this.data.definedNames.findIndex((d) => d.name === name);
    if (idx === -1) return false;
    this.data.definedNames.splice(idx, 1);
    return true;
  }

  // --- Print titles & print area ---

  /**
   * Gets the print titles (repeat rows/columns) for the given sheet.
   * @param sheetIndex 0-based sheet index.
   * @returns The defined name value (e.g., `"Sheet1!$1:$2,Sheet1!$A:$B"`) or `null`.
   */
  getPrintTitles(sheetIndex: number): string | null {
    const dn = this.data.definedNames?.find(
      (d) => d.name === '_xlnm.Print_Titles' && d.sheetId === sheetIndex,
    );
    return dn?.value ?? null;
  }

  /**
   * Sets print titles (repeat rows/columns) for the given sheet.
   * @param sheetIndex 0-based sheet index.
   * @param value The range reference (e.g., `"Sheet1!$1:$2"` for first 2 rows). Pass `null` to clear.
   */
  setPrintTitles(sheetIndex: number, value: string | null): void {
    this.data.definedNames ??= [];
    const idx = this.data.definedNames.findIndex(
      (d) => d.name === '_xlnm.Print_Titles' && d.sheetId === sheetIndex,
    );
    if (value === null) {
      if (idx !== -1) this.data.definedNames.splice(idx, 1);
    } else if (idx !== -1) {
      const entry = this.data.definedNames[idx];
      if (entry) entry.value = value;
    } else {
      this.data.definedNames.push({
        name: '_xlnm.Print_Titles',
        value,
        sheetId: sheetIndex,
      });
    }
  }

  /**
   * Gets the print area for the given sheet.
   * @param sheetIndex 0-based sheet index.
   * @returns The defined name value (e.g., `"Sheet1!$A$1:$D$50"`) or `null`.
   */
  getPrintArea(sheetIndex: number): string | null {
    const dn = this.data.definedNames?.find(
      (d) => d.name === '_xlnm.Print_Area' && d.sheetId === sheetIndex,
    );
    return dn?.value ?? null;
  }

  /**
   * Sets the print area for the given sheet.
   * @param sheetIndex 0-based sheet index.
   * @param value The range reference (e.g., `"Sheet1!$A$1:$D$50"`). Pass `null` to clear.
   */
  setPrintArea(sheetIndex: number, value: string | null): void {
    this.data.definedNames ??= [];
    const idx = this.data.definedNames.findIndex(
      (d) => d.name === '_xlnm.Print_Area' && d.sheetId === sheetIndex,
    );
    if (value === null) {
      if (idx !== -1) this.data.definedNames.splice(idx, 1);
    } else if (idx !== -1) {
      const entry = this.data.definedNames[idx];
      if (entry) entry.value = value;
    } else {
      this.data.definedNames.push({
        name: '_xlnm.Print_Area',
        value,
        sheetId: sheetIndex,
      });
    }
  }

  // --- Document properties ---

  /** Returns the document properties (title, author, etc.), or `null` if unset. */
  get docProperties(): DocPropertiesData | null {
    return this.data.docProperties ?? null;
  }

  /** Sets or clears the document properties (title, author, description, etc.). */
  set docProperties(props: DocPropertiesData | null) {
    this.data.docProperties = props;
  }

  // --- Theme colors ---

  /** Returns the parsed theme color palette, or `null` if no theme was present. */
  get themeColors(): ThemeColorsData | null {
    return this.data.themeColors ?? null;
  }

  // --- Calc chain ---

  /** Returns the formula calculation chain entries, if present. */
  get calcChain(): readonly CalcChainEntryData[] {
    return this.data.calcChain ?? [];
  }

  // --- Workbook views ---

  /** Returns the workbook view settings (window position, active tab, etc.). */
  get workbookViews(): readonly WorkbookViewData[] {
    return this.data.workbookViews ?? [];
  }

  /** Sets the workbook view settings (window position, active tab, etc.). */
  set workbookViews(views: WorkbookViewData[]) {
    this.data.workbookViews = views;
  }

  // --- Workbook protection ---

  /** Returns the workbook protection settings, or `null` if unprotected. */
  get protection(): WorkbookProtectionData | null {
    return this.data.protection ?? null;
  }

  /** Sets or clears the workbook protection settings. Pass `null` to remove protection. */
  set protection(value: WorkbookProtectionData | null) {
    this.data.protection = value;
  }

  // --- Pivot / Slicer / Timeline caches ---

  /** Returns the pivot cache definitions at the workbook level. */
  get pivotCaches(): readonly PivotCacheDefinitionData[] {
    return this.data.pivotCaches ?? [];
  }

  /**
   * Adds a pivot cache definition at the workbook level.
   *
   * @param cache - The pivot cache definition to add.
   */
  addPivotCache(cache: PivotCacheDefinitionData): void {
    this.data.pivotCaches ??= [];
    this.data.pivotCaches.push(cache);
  }

  /** Returns the pivot cache records at the workbook level. */
  get pivotCacheRecords(): readonly PivotCacheRecordsData[] {
    return this.data.pivotCacheRecords ?? [];
  }

  /**
   * Adds pivot cache records at the workbook level.
   *
   * @param records - The pivot cache records to add.
   */
  addPivotCacheRecords(records: PivotCacheRecordsData): void {
    this.data.pivotCacheRecords ??= [];
    this.data.pivotCacheRecords.push(records);
  }

  /** Returns the slicer cache definitions at the workbook level. */
  get slicerCaches(): readonly SlicerCacheData[] {
    return this.data.slicerCaches ?? [];
  }

  /**
   * Adds a slicer cache at the workbook level.
   *
   * @param cache - The slicer cache definition to add.
   */
  addSlicerCache(cache: SlicerCacheData): void {
    this.data.slicerCaches ??= [];
    this.data.slicerCaches.push(cache);
  }

  /** Returns the timeline cache definitions at the workbook level. */
  get timelineCaches(): readonly TimelineCacheData[] {
    return this.data.timelineCaches ?? [];
  }

  /**
   * Adds a timeline cache at the workbook level.
   *
   * @param cache - The timeline cache definition to add.
   */
  addTimelineCache(cache: TimelineCacheData): void {
    this.data.timelineCaches ??= [];
    this.data.timelineCaches.push(cache);
  }

  // --- Serialization ---

  /**
   * Serializes the workbook to an XLSX `Uint8Array` via the WASM writer.
   *
   * @param options - Optional write options. Pass `{ password: '...' }` to encrypt
   *   the file with Agile Encryption (AES-256-CBC, SHA-512). An empty string
   *   password is treated as "no encryption".
   * @returns The XLSX file as a Uint8Array.
   *
   * @example
   * ```ts
   * const bytes = await wb.toBuffer();
   * await fs.writeFile('output.xlsx', bytes);
   * ```
   */
  async toBuffer(options?: WriteOptions): Promise<Uint8Array> {
    ensureInitialized();
    if (options?.password) {
      return wasmWriteWithPassword(this.data, options.password);
    }
    return wasmWrite(this.data);
  }

  /**
   * Write the workbook to a file on disk.
   * Only available in Node.js, Bun, and Deno environments.
   *
   * @param path - File path to write to.
   * @param options - Optional write options. Pass `{ password: '...' }` to encrypt.
   *
   * @example
   * ```ts
   * await wb.toFile('./report.xlsx');
   * await wb.toFile('./secret.xlsx', { password: 's3cret' });
   * ```
   */
  async toFile(path: string, options?: WriteOptions): Promise<void> {
    const buffer = await this.toBuffer(options);
    const { writeFile } = await import('node:fs/promises');
    await writeFile(path, buffer);
  }

  /**
   * Validate the workbook for OOXML compliance issues.
   * Returns a structured report with errors, warnings, and fix suggestions.
   *
   * @returns A ValidationReport with categorized issues and severity levels.
   *
   * @example
   * ```ts
   * const report = wb.validate();
   * console.log(report.issues.length, 'issues found');
   * ```
   */
  validate(): ValidationReport {
    ensureInitialized();
    return wasmValidate(this.data);
  }

  /**
   * Validate and auto-repair the workbook.
   * Fixes repairable issues (dangling style indices, missing default styles,
   * bad sheet names, missing theme, invalid metadata, row ordering).
   *
   * @returns An object containing the repaired Workbook, a validation report, and the number of repairs made.
   *
   * @example
   * ```ts
   * const { workbook, repairCount } = wb.repair();
   * console.log(`${repairCount} issues repaired`);
   * ```
   */
  repair(): { workbook: Workbook; report: ValidationReport; repairCount: number } {
    ensureInitialized();
    const result: RepairResult = wasmRepair(this.data);
    return {
      workbook: new Workbook(result.workbook),
      report: result.report,
      repairCount: result.repairCount,
    };
  }

  // --- Image embedding ---

  /** Track next image ID for unique naming. */
  private imageCounter = 0;

  /**
   * Add an image (PNG bytes) anchored to a cell range on the given sheet.
   * The image is embedded via `preservedEntries` and OOXML drawing XML.
   *
   * @param sheetName - Target sheet name.
   * @param anchor - Cell range to anchor the image to.
   * @param imageBytes - Raw image bytes.
   * @param format - Image format extension (default: `'png'`).
   *
   * @example
   * ```ts
   * const png = fs.readFileSync('logo.png');
   * wb.addImage('Sheet1', { from: 'B2', to: 'E10' }, png);
   * ```
   */
  addImage(
    sheetName: string,
    anchor: ImageAnchor,
    imageBytes: Uint8Array,
    format: 'png' | 'jpeg' | 'gif' = 'png',
  ): void {
    const sheetIndex = this.data.sheets.findIndex((s) => s.name === sheetName);
    if (sheetIndex === -1)
      throw new ModernXlsxError(SHEET_NOT_FOUND, `Sheet "${sheetName}" not found`);

    this.data.preservedEntries ??= {};

    this.imageCounter++;
    const imageId = this.imageCounter;
    const sheetNum = sheetIndex + 1;
    const mediaPath = `xl/media/image${imageId}.${format}`;
    const drawingPath = `xl/drawings/drawing${sheetNum}.xml`;
    const drawingRelsPath = `xl/drawings/_rels/drawing${sheetNum}.xml.rels`;
    const sheetRelsPath = `xl/worksheets/_rels/sheet${sheetNum}.xml.rels`;
    const rId = `rId${imageId}`;

    // Store image bytes
    this.data.preservedEntries[mediaPath] = Array.from(imageBytes);

    // Build or extend drawing XML
    const sheetAnchors = this.imageAnchors.get(drawingPath) ?? [];
    sheetAnchors.push({ anchor, imageIndex: imageId });
    this.imageAnchors.set(drawingPath, sheetAnchors);
    const drawingXml = generateDrawingXml(sheetAnchors);
    this.data.preservedEntries[drawingPath] = Array.from(TEXT_ENCODER.encode(drawingXml));

    // Build drawing rels
    const sheetRels = this.imageRels.get(drawingRelsPath) ?? [];
    sheetRels.push({ rId, target: `../media/image${imageId}.${format}` });
    this.imageRels.set(drawingRelsPath, sheetRels);
    const drawingRels = generateDrawingRels(sheetRels);
    this.data.preservedEntries[drawingRelsPath] = Array.from(TEXT_ENCODER.encode(drawingRels));

    // Build or extend sheet rels (add drawing relationship)
    const existingSheetRels = this.data.preservedEntries[sheetRelsPath];
    let sheetRelsXml: string;
    if (existingSheetRels) {
      // Merge: insert a new Relationship before closing </Relationships>
      const existing = new TextDecoder().decode(new Uint8Array(existingSheetRels));
      const drawingRel = `<Relationship Id="rIdDrawing${sheetNum}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${sheetNum}.xml"/>`;
      if (!existing.includes('relationships/drawing')) {
        sheetRelsXml = existing.replace('</Relationships>', `${drawingRel}</Relationships>`);
      } else {
        sheetRelsXml = existing;
      }
    } else {
      sheetRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing${sheetNum}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${sheetNum}.xml"/></Relationships>`;
    }
    this.data.preservedEntries[sheetRelsPath] = Array.from(TEXT_ENCODER.encode(sheetRelsXml));
  }

  /**
   * Generate a barcode and embed it as an image anchored to the given cell range.
   *
   * @param sheetName - Target sheet name.
   * @param anchor - Cell range to anchor the barcode image to.
   * @param value - Data to encode in the barcode.
   * @param options - Barcode type, rendering, and sizing options.
   *
   * @example
   * ```ts
   * wb.addBarcode('Sheet1', { from: 'A1', to: 'C5' }, '12345', {
   *   type: 'code128',
   * });
   * ```
   */
  addBarcode(
    sheetName: string,
    anchor: ImageAnchor,
    value: string,
    options: DrawBarcodeOptions,
  ): void {
    const pngBytes = generateBarcode(value, options);
    this.addImage(sheetName, anchor, pngBytes);
  }

  /**
   * Returns the raw internal `WorkbookData` for serialization or inspection.
   *
   * @returns The complete workbook data model as a plain JSON-serializable object.
   */
  toJSON(): WorkbookData {
    return this.data;
  }
}

// ---------------------------------------------------------------------------
// Worksheet
// ---------------------------------------------------------------------------

/**
 * Represents a single worksheet within a workbook.
 *
 * Provides access to cells, rows, columns, merge ranges, charts, pivot tables,
 * comments, hyperlinks, page setup, and other sheet-level features.
 *
 * @example
 * ```ts
 * const ws = wb.addSheet('Data');
 * ws.cell('A1').value = 'Name';
 * ws.cell('B1').value = 'Score';
 * ws.frozenPane = { rows: 1, cols: 0 };
 * ```
 */
export class Worksheet {
  private readonly data: SheetData;
  /** @internal */ readonly styles: StylesData | undefined;
  /** @internal */ private readonly workbookData: WorkbookData | undefined;

  /**
   * Wraps an existing sheet data object. Typically obtained via
   * `Workbook.getSheet()` or `Workbook.addSheet()`.
   *
   * @param data - The underlying sheet data model.
   * @param styles - Optional shared styles from the parent workbook.
   * @param workbookData - Optional parent workbook data for cross-sheet operations.
   */
  constructor(data: SheetData, styles?: StylesData, workbookData?: WorkbookData) {
    this.data = data;
    this.styles = styles;
    this.workbookData = workbookData;
  }

  /** Returns the sheet tab name. */
  get name(): string {
    return this.data.name;
  }

  /**
   * Renames the sheet tab. Validated per ECMA-376 rules.
   * @throws If the name is empty, exceeds 31 characters, or contains forbidden characters.
   */
  set name(value: string) {
    validateSheetName(value);
    this.data.name = value;
  }

  /** Returns the sheet visibility state. */
  get state(): SheetState {
    return this.data.state ?? 'visible';
  }

  /** Sets the sheet visibility state. */
  set state(value: SheetState) {
    this.data.state = value === 'visible' ? null : value;
  }

  /** Returns the worksheet dimension string from the source file, or `null` if not set. */
  get dimension(): string | null {
    return this.data.worksheet.dimension;
  }

  /** Returns the number of populated rows in the sheet. */
  get rowCount(): number {
    return this.data.worksheet.rows.length;
  }

  /** Returns all rows in the sheet, sorted by 1-based row index. */
  get rows(): readonly RowData[] {
    return this.data.worksheet.rows;
  }

  // --- Columns ---

  /** Returns the column definitions (width, visibility) for this sheet. */
  get columns(): readonly ColumnInfo[] {
    return this.data.worksheet.columns;
  }

  /** Replaces all column definitions for this sheet. */
  set columns(cols: ColumnInfo[]) {
    this.data.worksheet.columns = cols;
  }

  /**
   * Sets the width of a single 1-based column, creating or updating its definition.
   *
   * @param col - The 1-based column index.
   * @param width - The column width in Excel character units.
   *
   * @example
   * ```ts
   * ws.setColumnWidth(1, 20); // column A = 20 chars wide
   * ```
   */
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

  /** Returns all merged cell ranges (e.g., `"A1:C3"`). */
  get mergeCells(): readonly string[] {
    return this.data.worksheet.mergeCells;
  }

  /**
   * Adds a merged cell range (e.g., `"A1:C3"`). Duplicates are ignored.
   *
   * @param range - The A1-style range to merge (e.g., `'A1:C3'`).
   *
   * @example
   * ```ts
   * ws.addMergeCell('A1:D1'); // merge header across 4 columns
   * ```
   */
  addMergeCell(range: string): void {
    if (!this.data.worksheet.mergeCells.includes(range)) {
      this.data.worksheet.mergeCells.push(range);
    }
  }

  /**
   * Removes a merged cell range.
   *
   * @param range - The A1-style range to un-merge (e.g., `'A1:C3'`).
   * @returns `true` if the range was removed, `false` if not found.
   */
  removeMergeCell(range: string): boolean {
    const idx = this.data.worksheet.mergeCells.indexOf(range);
    if (idx === -1) return false;
    this.data.worksheet.mergeCells.splice(idx, 1);
    return true;
  }

  // --- Auto filter ---

  /** Returns the auto-filter configuration, or `null` if none is set. */
  get autoFilter(): AutoFilterData | null {
    return this.data.worksheet.autoFilter;
  }

  /** Sets the auto-filter. Accepts an `AutoFilterData` object, a range string (e.g., `"A1:D10"`), or `null` to clear. */
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

  /** Returns all hyperlinks defined in this sheet. */
  get hyperlinks(): readonly HyperlinkData[] {
    return this.data.worksheet.hyperlinks ?? [];
  }

  /**
   * Adds a hyperlink to a cell. The `location` is a URL or internal reference.
   *
   * @param cellRef - The cell reference to attach the hyperlink to (e.g., `'A1'`).
   * @param location - URL or internal reference (e.g., `'https://example.com'` or `'Sheet2!A1'`).
   * @param opts - Optional display text and tooltip.
   *
   * @example
   * ```ts
   * ws.addHyperlink('A1', 'https://example.com', {
   *   display: 'Visit Site',
   *   tooltip: 'Opens example.com',
   * });
   * ```
   */
  addHyperlink(
    cellRef: string,
    location: string,
    opts?: { display?: string; tooltip?: string },
  ): void {
    this.data.worksheet.hyperlinks ??= [];
    this.data.worksheet.hyperlinks.push({
      cellRef,
      location,
      display: opts?.display ?? null,
      tooltip: opts?.tooltip ?? null,
    });
  }

  /**
   * Removes the hyperlink on the given cell reference.
   *
   * @param cellRef - The cell reference to remove the hyperlink from (e.g., `'A1'`).
   * @returns `true` if the hyperlink was removed, `false` if not found.
   */
  removeHyperlink(cellRef: string): boolean {
    if (!this.data.worksheet.hyperlinks) return false;
    const idx = this.data.worksheet.hyperlinks.findIndex((h) => h.cellRef === cellRef);
    if (idx === -1) return false;
    this.data.worksheet.hyperlinks.splice(idx, 1);
    return true;
  }

  // --- Page setup ---

  /** Returns the page setup configuration (orientation, paper size, etc.), or `null` if unset. */
  get pageSetup(): PageSetupData | null {
    return this.data.worksheet.pageSetup ?? null;
  }

  /** Sets or clears the page setup configuration. */
  set pageSetup(setup: PageSetupData | null) {
    this.data.worksheet.pageSetup = setup;
  }

  // --- Sheet protection ---

  /** Returns the sheet protection settings, or `null` if the sheet is unprotected. */
  get sheetProtection(): SheetProtectionData | null {
    return this.data.worksheet.sheetProtection ?? null;
  }

  /** Sets or clears sheet protection. Pass `null` to remove protection. */
  set sheetProtection(protection: SheetProtectionData | null) {
    this.data.worksheet.sheetProtection = protection;
  }

  // --- Frozen pane ---

  /** Returns the frozen pane split position, or `null` if no panes are frozen. */
  get frozenPane(): FrozenPane | null {
    return this.data.worksheet.frozenPane;
  }

  /** Freezes rows and/or columns at the given split position, or clears the freeze with `null`. */
  set frozenPane(pane: FrozenPane | null) {
    this.data.worksheet.frozenPane = pane;
    if (pane) {
      this.data.worksheet.splitPane = null;
      this.data.worksheet.paneSelections = [];
    }
  }

  // --- Split pane ---

  /** Returns the split pane configuration, or `null` if no split pane is active. */
  get splitPane(): SplitPaneData | null {
    return this.data.worksheet.splitPane ?? null;
  }

  /** Sets a split pane. Setting a split pane clears any frozen pane. Pass `null` to clear. */
  set splitPane(pane: SplitPaneData | null) {
    this.data.worksheet.splitPane = pane;
    if (pane) {
      this.data.worksheet.frozenPane = null;
    } else {
      this.data.worksheet.paneSelections = [];
    }
  }

  // --- Pane selections ---

  /** Returns per-pane selection state, or empty array. */
  get paneSelections(): readonly PaneSelectionData[] {
    return this.data.worksheet.paneSelections ?? [];
  }

  /** Sets per-pane selection state. */
  set paneSelections(selections: PaneSelectionData[]) {
    this.data.worksheet.paneSelections = selections;
  }

  // --- Sheet view ---

  /** Returns the sheet view configuration, or `null` if using defaults. */
  get view(): SheetViewData | null {
    return this.data.worksheet.sheetView ?? null;
  }

  /** Sets sheet view configuration. Pass `null` to reset to defaults. */
  set view(sv: SheetViewData | null) {
    this.data.worksheet.sheetView = sv;
  }

  /** Returns the view mode: `'normal'`, `'pageBreakPreview'`, or `'pageLayout'`. */
  get viewMode(): ViewMode {
    return this.data.worksheet.sheetView?.view ?? 'normal';
  }

  /** Sets the view mode. Creates a sheet view if none exists. */
  set viewMode(mode: ViewMode) {
    this.data.worksheet.sheetView ??= {};
    this.data.worksheet.sheetView.view = mode === 'normal' ? null : mode;
  }

  // --- Tab color ---

  /** Returns the sheet tab color as an RGB hex string (e.g., `"FF0000"`) or `null`. */
  get tabColor(): string | null {
    return this.data.worksheet.tabColor ?? null;
  }

  /** Sets the sheet tab color. Pass `null` to clear. */
  set tabColor(color: string | null) {
    this.data.worksheet.tabColor = color;
  }

  // --- Used range ---

  /** Returns the computed used range (e.g., `"B2:D5"`) or `null` if the sheet has no cells. */
  get usedRange(): string | null {
    let minRow = Number.MAX_SAFE_INTEGER;
    let maxRow = 0;
    let minCol = Number.MAX_SAFE_INTEGER;
    let maxCol = 0;
    let hasCell = false;

    for (const row of this.data.worksheet.rows) {
      for (const cell of row.cells) {
        const { row: r, col: c } = decodeCellRef(cell.reference);
        hasCell = true;
        if (r < minRow) minRow = r;
        if (r > maxRow) maxRow = r;
        if (c < minCol) minCol = c;
        if (c > maxCol) maxCol = c;
      }
    }

    if (!hasCell) return null;
    return `${columnToLetter(minCol)}${minRow + 1}:${columnToLetter(maxCol)}${maxRow + 1}`;
  }

  // --- Data validations ---

  /** Returns all data validation rules applied to this sheet. */
  get validations(): readonly DataValidationData[] {
    return this.data.worksheet.dataValidations ?? [];
  }

  /**
   * Adds a data validation rule to the given cell range reference.
   *
   * @param ref - The cell range to validate (e.g., `'B2:B100'`).
   * @param rule - The validation rule (type, operator, formula, prompt, etc.).
   *
   * @example
   * ```ts
   * ws.addValidation('B2:B100', {
   *   type: 'list',
   *   formula1: '"Yes,No,Maybe"',
   * });
   * ```
   */
  addValidation(ref: string, rule: Omit<DataValidationData, 'sqref'>): void {
    this.data.worksheet.dataValidations ??= [];
    this.data.worksheet.dataValidations.push({ sqref: ref, ...rule });
  }

  /**
   * Removes the data validation rule for the given cell range reference.
   *
   * @param ref - The cell range whose validation rule should be removed.
   * @returns `true` if the rule was removed, `false` if not found.
   */
  removeValidation(ref: string): boolean {
    if (!this.data.worksheet.dataValidations) return false;
    const idx = this.data.worksheet.dataValidations.findIndex((v) => v.sqref === ref);
    if (idx === -1) return false;
    this.data.worksheet.dataValidations.splice(idx, 1);
    return true;
  }

  // --- Comments ---

  /** Returns all cell comments in this sheet. */
  get comments(): readonly CommentData[] {
    return this.data.worksheet.comments ?? [];
  }

  /**
   * Adds a comment to the given cell reference.
   *
   * @param cellRef - The cell reference (e.g., `'A1'`).
   * @param author - The comment author name.
   * @param text - The comment text.
   *
   * @example
   * ```ts
   * ws.addComment('A1', 'Alice', 'Please review this value');
   * ```
   */
  addComment(cellRef: string, author: string, text: string): void {
    this.data.worksheet.comments ??= [];
    this.data.worksheet.comments.push({ cellRef, author, text });
  }

  /**
   * Removes the comment on the given cell reference.
   *
   * @param cellRef - The cell reference whose comment should be removed.
   * @returns `true` if the comment was removed, `false` if not found.
   */
  removeComment(cellRef: string): boolean {
    if (!this.data.worksheet.comments) return false;
    const idx = this.data.worksheet.comments.findIndex((c) => c.cellRef === cellRef);
    if (idx === -1) return false;
    this.data.worksheet.comments.splice(idx, 1);
    return true;
  }

  // --- Threaded Comments ---

  /** Returns all threaded comments in this sheet. */
  get threadedComments(): readonly ThreadedCommentData[] {
    return this.data.worksheet.threadedComments ?? [];
  }

  /**
   * Adds a threaded comment to the given cell reference.
   * Creates a person entry in the workbook if the author does not exist yet.
   *
   * @param cell - The cell reference (e.g., `'A1'`).
   * @param text - The comment text.
   * @param author - The display name of the author.
   * @returns The unique ID of the created comment.
   *
   * @example
   * ```ts
   * const id = ws.addThreadedComment('A1', 'Looks good!', 'Alice');
   * ws.replyToComment(id, 'Thanks!', 'Bob');
   * ```
   */
  addThreadedComment(cell: string, text: string, author: string): string {
    const wb = this.workbookData;
    if (!wb) {
      throw new Error(
        'Worksheet must be created via Workbook.addSheet() or Workbook.getSheet() to use threaded comments',
      );
    }
    wb.persons ??= [];
    let person = wb.persons.find((p) => p.displayName === author);
    if (!person) {
      person = { id: crypto.randomUUID(), displayName: author };
      wb.persons.push(person);
    }
    const comment: ThreadedCommentData = {
      id: crypto.randomUUID(),
      refCell: cell,
      personId: person.id,
      text,
      timestamp: new Date().toISOString(),
    };
    this.data.worksheet.threadedComments ??= [];
    this.data.worksheet.threadedComments.push(comment);
    return comment.id;
  }

  /**
   * Replies to an existing threaded comment.
   * Creates a person entry in the workbook if the author does not exist yet.
   * @param commentId - The ID of the parent comment to reply to.
   * @param text - The reply text.
   * @param author - The display name of the reply author.
   * @returns The unique ID of the reply comment.
   * @throws If the parent comment is not found.
   */
  replyToComment(commentId: string, text: string, author: string): string {
    const parent = this.data.worksheet.threadedComments?.find((c) => c.id === commentId);
    if (!parent) {
      throw new ModernXlsxError(COMMENT_NOT_FOUND, `Comment ${commentId} not found`);
    }
    const wb = this.workbookData;
    if (!wb) {
      throw new Error(
        'Worksheet must be created via Workbook.addSheet() or Workbook.getSheet() to use threaded comments',
      );
    }
    wb.persons ??= [];
    let person = wb.persons.find((p) => p.displayName === author);
    if (!person) {
      person = { id: crypto.randomUUID(), displayName: author };
      wb.persons.push(person);
    }
    const reply: ThreadedCommentData = {
      id: crypto.randomUUID(),
      refCell: parent.refCell,
      personId: person.id,
      text,
      timestamp: new Date().toISOString(),
      parentId: commentId,
    };
    this.data.worksheet.threadedComments ??= [];
    this.data.worksheet.threadedComments.push(reply);
    return reply.id;
  }

  // --- Page margins ---

  /** Returns the page margins (top, bottom, left, right, header, footer), or `null` if unset. */
  get pageMargins(): PageMarginsData | null {
    return this.data.worksheet.pageSetup?.margins ?? null;
  }

  /** Sets or clears the page margins. Creates a page setup object if needed. */
  set pageMargins(margins: PageMarginsData | null) {
    this.data.worksheet.pageSetup ??= {};
    this.data.worksheet.pageSetup.margins = margins;
  }

  // --- Cell access ---

  /**
   * Returns the cell at the given A1-style reference, creating the row and cell if needed.
   * The returned `Cell` is a live wrapper -- mutations are reflected in the underlying sheet data.
   *
   * @param ref - A1-style cell reference (e.g., `'B3'`, `'AA100'`).
   * @returns A mutable Cell wrapper.
   *
   * @example
   * ```ts
   * const cell = ws.cell('B3');
   * cell.value = 42;
   * cell.formula = 'A3*2';
   * ```
   */
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

    return new Cell(cellData, this.styles);
  }

  // --- Row utilities ---

  /**
   * Sets the height (in points) of the given 1-based row, creating it if needed.
   *
   * @param rowIndex - The 1-based row index.
   * @param height - Row height in points (e.g., 24 for double-height).
   *
   * @example
   * ```ts
   * ws.setRowHeight(1, 30); // set header row to 30pt
   * ```
   */
  setRowHeight(rowIndex: number, height: number): void {
    const row = this.ensureRow(rowIndex);
    row.height = height;
  }

  /**
   * Sets the visibility of the given 1-based row, creating it if needed.
   *
   * @param rowIndex - The 1-based row index.
   * @param hidden - `true` to hide the row, `false` to show it.
   */
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

  // --- Tables (Excel ListObjects) ---

  /** Returns all table definitions on this sheet. */
  get tables(): readonly TableDefinitionData[] {
    return this.data.worksheet.tables ?? [];
  }

  /**
   * Returns the table with the given display name, or `undefined` if not found.
   *
   * @param displayName - The table display name to look up.
   * @returns The matching TableDefinitionData, or `undefined`.
   */
  getTable(displayName: string): TableDefinitionData | undefined {
    return this.data.worksheet.tables?.find((t) => t.displayName === displayName);
  }

  /**
   * Adds a table definition to the sheet.
   *
   * @param table - The table definition to add (range, columns, style, etc.).
   */
  addTable(table: TableDefinitionData): void {
    this.data.worksheet.tables ??= [];
    this.data.worksheet.tables.push(table);
  }

  /**
   * Removes the table with the given display name.
   *
   * @param displayName - The table display name to remove.
   * @returns `true` if the table was removed, `false` if not found.
   */
  removeTable(displayName: string): boolean {
    if (!this.data.worksheet.tables) return false;
    const idx = this.data.worksheet.tables.findIndex((t) => t.displayName === displayName);
    if (idx === -1) return false;
    this.data.worksheet.tables.splice(idx, 1);
    return true;
  }

  // --- Charts ---

  /** Returns all charts on this sheet. */
  get charts(): readonly WorksheetChartData[] {
    return this.data.worksheet.charts ?? [];
  }

  /** Returns all pivot tables on this sheet. */
  get pivotTables(): readonly PivotTableData[] {
    return this.data.worksheet.pivotTables ?? [];
  }

  /** Returns all slicers on this sheet. */
  get slicers(): readonly SlicerData[] {
    return this.data.worksheet.slicers ?? [];
  }

  /** Returns all timelines on this sheet. */
  get timelines(): readonly TimelineData[] {
    return this.data.worksheet.timelines ?? [];
  }

  // --- Pivot Table / Slicer / Timeline mutation ---

  /**
   * Adds a pivot table definition to this sheet.
   *
   * @param pt - The pivot table data to add.
   */
  addPivotTable(pt: PivotTableData): void {
    this.data.worksheet.pivotTables ??= [];
    this.data.worksheet.pivotTables.push(pt);
  }

  /**
   * Adds a pivot table using a fluent builder callback.
   *
   * @example
   * ```ts
   * ws.addPivotTableFromBuilder((b) => {
   *   b.name('SalesPivot')
   *    .cacheId(0)
   *    .location('A3:D20')
   *    .addRowField({ fieldIndex: 0, name: 'Region' })
   *    .addDataField({ fieldIndex: 2, subtotal: 'sum', name: 'Total Revenue' });
   * });
   * ```
   */
  addPivotTableFromBuilder(configure: (builder: PivotTableBuilder) => void): void {
    const builder = new PivotTableBuilder();
    configure(builder);
    this.addPivotTable(builder.build());
  }

  /**
   * Removes a pivot table by index.
   *
   * @param index - Zero-based index of the pivot table to remove.
   * @returns `true` if the pivot table was removed, `false` if index out of range.
   */
  removePivotTable(index: number): boolean {
    if (
      !this.data.worksheet.pivotTables ||
      index < 0 ||
      index >= this.data.worksheet.pivotTables.length
    )
      return false;
    this.data.worksheet.pivotTables.splice(index, 1);
    return true;
  }

  /**
   * Adds a slicer to this sheet.
   *
   * @param slicer - The slicer definition to add.
   */
  addSlicer(slicer: SlicerData): void {
    this.data.worksheet.slicers ??= [];
    this.data.worksheet.slicers.push(slicer);
  }

  /**
   * Removes a slicer by index.
   *
   * @param index - Zero-based index of the slicer to remove.
   * @returns `true` if the slicer was removed, `false` if index out of range.
   */
  removeSlicer(index: number): boolean {
    if (!this.data.worksheet.slicers || index < 0 || index >= this.data.worksheet.slicers.length)
      return false;
    this.data.worksheet.slicers.splice(index, 1);
    return true;
  }

  /**
   * Adds a timeline to this sheet.
   *
   * @param timeline - The timeline definition to add.
   */
  addTimeline(timeline: TimelineData): void {
    this.data.worksheet.timelines ??= [];
    this.data.worksheet.timelines.push(timeline);
  }

  /**
   * Removes a timeline by index.
   *
   * @param index - Zero-based index of the timeline to remove.
   * @returns `true` if the timeline was removed, `false` if index out of range.
   */
  removeTimeline(index: number): boolean {
    if (
      !this.data.worksheet.timelines ||
      index < 0 ||
      index >= this.data.worksheet.timelines.length
    )
      return false;
    this.data.worksheet.timelines.splice(index, 1);
    return true;
  }

  /**
   * Adds a chart to this worksheet using a builder callback.
   *
   * @param type - The chart type (e.g., `'bar'`, `'line'`, `'pie'`, `'scatter'`).
   * @param configure - A callback that receives a ChartBuilder for fluent configuration.
   *
   * @example
   * ```ts
   * ws.addChart('bar', (b) => {
   *   b.title('Sales')
   *    .addSeries({ valRef: 'Sheet1!$B$2:$B$5' })
   *    .legend('bottom');
   * });
   * ```
   */
  addChart(type: ChartType, configure: (builder: ChartBuilder) => void): void {
    const builder = new ChartBuilder(type);
    configure(builder);
    const chartData = builder.build();
    validateChartData(chartData);
    this.data.worksheet.charts ??= [];
    this.data.worksheet.charts.push(chartData);
  }

  /**
   * Adds a pre-built chart definition to this worksheet.
   *
   * @param chart - The complete chart data object (produced by `ChartBuilder.build()`).
   */
  addChartData(chart: WorksheetChartData): void {
    validateChartData(chart);
    this.data.worksheet.charts ??= [];
    this.data.worksheet.charts.push(chart);
  }

  /**
   * Removes a chart by index.
   *
   * @param index - Zero-based index of the chart to remove.
   * @returns `true` if removed, `false` if index out of range.
   */
  removeChart(index: number): boolean {
    if (!this.data.worksheet.charts || index < 0 || index >= this.data.worksheet.charts.length)
      return false;
    this.data.worksheet.charts.splice(index, 1);
    return true;
  }

  // --- Header / Footer ---

  /** Returns the header/footer configuration, or `null` if unset. */
  get headerFooter(): HeaderFooterData | null {
    return this.data.worksheet.headerFooter ?? null;
  }

  /** Sets or clears the header/footer configuration. */
  set headerFooter(hf: HeaderFooterData | null) {
    this.data.worksheet.headerFooter = hf;
  }

  // --- Page breaks ---

  /** Returns the page breaks configuration, or `null` if unset. */
  get pageBreaks(): PageBreaksData | null {
    return this.data.worksheet.pageBreaks ?? null;
  }

  /** Sets or clears the page breaks configuration. */
  set pageBreaks(pb: PageBreaksData | null) {
    this.data.worksheet.pageBreaks = pb;
  }

  /**
   * Adds a manual row page break (horizontal break before the given row).
   * @param row 1-based row index where the page break is inserted.
   */
  addRowBreak(row: number): void {
    this.data.worksheet.pageBreaks ??= {};
    this.data.worksheet.pageBreaks.rowBreaks ??= [];
    const brk: PageBreakData = { id: row, max: 16383, man: true };
    this.data.worksheet.pageBreaks.rowBreaks.push(brk);
  }

  /**
   * Adds a manual column page break (vertical break before the given column).
   * @param col 1-based column index where the page break is inserted.
   */
  addColBreak(col: number): void {
    this.data.worksheet.pageBreaks ??= {};
    this.data.worksheet.pageBreaks.colBreaks ??= [];
    const brk: PageBreakData = { id: col, max: 1048575, man: true };
    this.data.worksheet.pageBreaks.colBreaks.push(brk);
  }

  // --- Print area ---

  /**
   * Sets the print area for this sheet via a `_xlnm.Print_Area` defined name.
   * @param ref The range reference (e.g., `"Sheet1!$A$1:$D$50"`). Pass `null` to clear.
   */
  setPrintArea(ref: string | null): void {
    if (!this.workbookData) return;
    const sheetIndex = this.workbookData.sheets.indexOf(this.data);
    if (sheetIndex === -1) return;
    this.workbookData.definedNames ??= [];
    const idx = this.workbookData.definedNames.findIndex(
      (d) => d.name === '_xlnm.Print_Area' && d.sheetId === sheetIndex,
    );
    if (ref === null) {
      if (idx !== -1) this.workbookData.definedNames.splice(idx, 1);
    } else if (idx !== -1) {
      const entry = this.workbookData.definedNames[idx];
      if (entry) entry.value = ref;
    } else {
      this.workbookData.definedNames.push({
        name: '_xlnm.Print_Area',
        value: ref,
        sheetId: sheetIndex,
      });
    }
  }

  /**
   * Gets the print area for this sheet.
   * @returns The defined name value (e.g., `"Sheet1!$A$1:$D$50"`) or `null`.
   */
  getPrintArea(): string | null {
    if (!this.workbookData) return null;
    const sheetIndex = this.workbookData.sheets.indexOf(this.data);
    if (sheetIndex === -1) return null;
    const dn = this.workbookData.definedNames?.find(
      (d) => d.name === '_xlnm.Print_Area' && d.sheetId === sheetIndex,
    );
    return dn?.value ?? null;
  }

  // --- Outline properties ---

  /** Returns the outline (grouping) properties, or `null` if unset. */
  get outlineProperties(): OutlinePropertiesData | null {
    return this.data.worksheet.outlineProperties ?? null;
  }

  /** Sets or clears the outline properties (summaryBelow, summaryRight). */
  set outlineProperties(props: OutlinePropertiesData | null) {
    this.data.worksheet.outlineProperties = props;
  }

  // --- Row / Column grouping ---

  /**
   * Groups rows by setting their outline level (1-7).
   * @param startRow 1-based start row index (inclusive).
   * @param endRow 1-based end row index (inclusive).
   * @param level Outline level (1-7). Defaults to 1.
   */
  groupRows(startRow: number, endRow: number, level = 1): void {
    const clamped = Math.min(Math.max(level, 0), 7);
    for (let i = startRow; i <= endRow; i++) {
      const row = this.ensureRow(i);
      row.outlineLevel = clamped > 0 ? clamped : null;
    }
  }

  /**
   * Ungroups rows by clearing their outline level.
   * @param startRow 1-based start row index (inclusive).
   * @param endRow 1-based end row index (inclusive).
   */
  ungroupRows(startRow: number, endRow: number): void {
    for (let i = startRow; i <= endRow; i++) {
      const row = findRowBinary(this.data.worksheet.rows, i);
      if (row) {
        row.outlineLevel = null;
        row.collapsed = false;
      }
    }
  }

  /**
   * Groups columns by setting their outline level (1-7).
   * @param startCol 1-based start column index (inclusive).
   * @param endCol 1-based end column index (inclusive).
   * @param level Outline level (1-7). Defaults to 1.
   */
  groupColumns(startCol: number, endCol: number, level = 1): void {
    const clamped = Math.min(Math.max(level, 0), 7);
    const clampedLevel = clamped > 0 ? clamped : null;
    for (let col = startCol; col <= endCol; col++) {
      const existing = this.data.worksheet.columns.find((c) => c.min === col && c.max === col);
      if (existing) {
        existing.outlineLevel = clampedLevel;
      } else {
        this.data.worksheet.columns.push({
          min: col,
          max: col,
          width: 8.43,
          hidden: false,
          customWidth: false,
          outlineLevel: clampedLevel,
        });
      }
    }
  }

  /**
   * Ungroups columns by clearing their outline level.
   * @param startCol 1-based start column index (inclusive).
   * @param endCol 1-based end column index (inclusive).
   */
  ungroupColumns(startCol: number, endCol: number): void {
    for (let col = startCol; col <= endCol; col++) {
      const existing = this.data.worksheet.columns.find((c) => c.min === col && c.max === col);
      if (existing) {
        existing.outlineLevel = null;
        existing.collapsed = false;
      }
    }
  }

  /**
   * Collapses grouped rows at the given outline level.
   * @param startRow 1-based start row index (inclusive).
   * @param endRow 1-based end row index (inclusive).
   */
  collapseRows(startRow: number, endRow: number): void {
    for (let i = startRow; i <= endRow; i++) {
      const row = findRowBinary(this.data.worksheet.rows, i);
      if (row?.outlineLevel) {
        row.hidden = true;
      }
    }
    // Set collapsed flag on the summary row (row after group if summaryBelow).
    const summaryBelow = this.data.worksheet.outlineProperties?.summaryBelow !== false;
    const summaryIdx = summaryBelow ? endRow + 1 : startRow - 1;
    if (summaryIdx > 0) {
      const summary = this.ensureRow(summaryIdx);
      summary.collapsed = true;
    }
  }

  /**
   * Expands collapsed grouped rows.
   * @param startRow 1-based start row index (inclusive).
   * @param endRow 1-based end row index (inclusive).
   */
  expandRows(startRow: number, endRow: number): void {
    for (let i = startRow; i <= endRow; i++) {
      const row = findRowBinary(this.data.worksheet.rows, i);
      if (row?.outlineLevel) {
        row.hidden = false;
      }
    }
    const summaryBelow = this.data.worksheet.outlineProperties?.summaryBelow !== false;
    const summaryIdx = summaryBelow ? endRow + 1 : startRow - 1;
    if (summaryIdx > 0) {
      const summary = findRowBinary(this.data.worksheet.rows, summaryIdx);
      if (summary) summary.collapsed = false;
    }
  }

  // -------------------------------------------------------------------------
  // Sparklines
  // -------------------------------------------------------------------------

  /** Returns all sparkline groups on this sheet. */
  get sparklineGroups(): readonly SparklineGroupData[] {
    return this.data.worksheet.sparklineGroups ?? [];
  }

  /**
   * Adds a sparkline group to this sheet.
   *
   * @param group - The sparkline group definition.
   */
  addSparklineGroup(group: SparklineGroupData): void {
    this.data.worksheet.sparklineGroups ??= [];
    this.data.worksheet.sparklineGroups.push(group);
  }

  /** Removes all sparkline groups from this sheet. */
  clearSparklineGroups(): void {
    delete this.data.worksheet.sparklineGroups;
  }
}

// ---------------------------------------------------------------------------
// Cell
// ---------------------------------------------------------------------------

/**
 * Represents a single cell within a worksheet row.
 * Mutations are applied directly to the underlying data.
 *
 * Obtained via `Worksheet.cell()`. Setting `value` auto-detects the cell type
 * from the JS type (string, number, boolean).
 *
 * @example
 * ```ts
 * const cell = ws.cell('A1');
 * cell.value = 'Hello';
 * cell.formula = 'UPPER(A2)';
 * cell.styleIndex = boldStyleIdx;
 * ```
 */
export class Cell {
  private readonly data: CellData;
  private readonly styles: StylesData | undefined;

  /**
   * Wraps an existing cell data object. Typically obtained via `Worksheet.cell()`.
   *
   * @param data - The underlying cell data model.
   * @param styles - Optional shared styles for number format resolution.
   */
  constructor(data: CellData, styles?: StylesData) {
    this.data = data;
    this.styles = styles;
  }

  /** Returns the A1-style cell reference (e.g., `"B3"`). */
  get reference(): string {
    return this.data.reference;
  }

  /** Returns the cell's data type (e.g., `"number"`, `"sharedString"`, `"boolean"`, `"formulaStr"`). */
  get type(): CellType {
    return this.data.cellType;
  }

  /** Returns the 0-based index into `Workbook.styles.cellXfs`, or `null` if using the default style. */
  get styleIndex(): number | null {
    return this.data.styleIndex ?? null;
  }

  /** Sets the style index, or pass `null` to clear custom styling. */
  set styleIndex(value: number | null) {
    this.data.styleIndex = value ?? null;
  }

  /** Returns the cell value, coerced to the appropriate JS type based on `cellType`. */
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

  /**
   * Sets the cell value and auto-detects the cell type from the JS type.
   * Strings become shared strings, numbers stay numeric, booleans become `"1"`/`"0"`.
   * If a formula is set, the cell type is preserved (value acts as the cached result).
   */
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

  /** Returns the cell's formula string (without the leading `=`), or `null` if no formula is set. */
  get formula(): string | null {
    return this.data.formula ?? null;
  }

  /** Sets or clears the cell's formula. Setting a formula changes the cell type to `"formulaStr"`. */
  set formula(f: string | null) {
    this.data.formula = f;
    if (f !== null) {
      this.data.cellType = 'formulaStr';
    }
  }

  /**
   * Returns the resolved number format code (e.g., `"#,##0.00"`, `"yyyy-mm-dd"`),
   * or `null` if the cell uses the default General format.
   */
  get numberFormat(): string | null {
    if (!this.styles || this.data.styleIndex == null) return null;
    const xf = this.styles.cellXfs[this.data.styleIndex];
    if (!xf || xf.numFmtId === 0 || xf.numFmtId == null) return null;

    // Check builtin formats first (id 1–49)
    const builtin = getBuiltinFormat(xf.numFmtId);
    if (builtin && builtin !== 'General') return builtin;

    // Check custom numFmts
    const custom = this.styles.numFmts?.find((nf) => nf.id === xf.numFmtId);
    return custom?.formatCode ?? null;
  }

  /**
   * Returns a `Date` if this cell contains a date-formatted number, or `null` otherwise.
   * Uses the cell's number format to determine whether the numeric value represents a date.
   */
  get dateValue(): Date | null {
    if (this.data.cellType !== 'number' || this.data.value == null) return null;
    const fmt = this.numberFormat;
    if (!fmt || !isDateFormatCode(fmt)) return null;
    return serialToDate(Number.parseFloat(this.data.value));
  }

  /**
   * Returns the rich text runs for this cell, or `undefined` if the cell
   * has no rich text formatting. Each run contains a text segment and
   * optional formatting properties (bold, italic, underline, strike,
   * fontName, fontSize, color).
   */
  get richText(): readonly RichTextRun[] | undefined {
    return this.data.richText;
  }

  /**
   * Sets rich text runs on this cell, enabling mixed formatting within a
   * single cell. The cell type is automatically set to `"sharedString"` and
   * the plain-text value is updated to the concatenation of all run texts.
   *
   * Pass `undefined` to clear rich text.
   */
  set richText(runs: readonly RichTextRun[] | undefined) {
    if (runs === undefined) {
      delete this.data.richText;
      return;
    }
    this.data.richText = [...runs];
    // Auto-set the cell type and plain-text value.
    this.data.cellType = 'sharedString';
    this.data.value = runs.map((r) => r.text).join('');
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

// ---------------------------------------------------------------------------
// Input validation
// ---------------------------------------------------------------------------

const INVALID_SHEET_CHARS = /[\\/*?[\]:]/;

/** Validate an Excel sheet name per ECMA-376 §18.3.1.73 constraints. */
function validateSheetName(name: string): void {
  if (name.length === 0) {
    throw new ModernXlsxError(INVALID_ARGUMENT, 'Sheet name must not be empty');
  }
  if (name.length > 31) {
    throw new ModernXlsxError(
      INVALID_ARGUMENT,
      `Sheet name must be 31 characters or fewer (got ${name.length})`,
    );
  }
  if (INVALID_SHEET_CHARS.test(name)) {
    throw new ModernXlsxError(
      INVALID_ARGUMENT,
      `Sheet name contains invalid characters: \\ / * ? [ ] :`,
    );
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new ModernXlsxError(
      INVALID_ARGUMENT,
      'Sheet name must not start or end with an apostrophe',
    );
  }
}
