import type { DrawBarcodeOptions, ImageAnchor } from './barcode.js';
import { generateBarcode, generateDrawingRels, generateDrawingXml } from './barcode.js';
import { columnToLetter, decodeCellRef } from './cell-ref.js';
import { isDateFormatCode, serialToDate } from './dates.js';
import { getBuiltinFormat } from './format-cell.js';
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
  HeaderFooterData,
  HyperlinkData,
  OutlinePropertiesData,
  PageMarginsData,
  PageSetupData,
  PaneSelectionData,
  RepairResult,
  SheetViewData,
  RowData,
  SheetData,
  SheetProtectionData,
  SplitPaneData,
  StylesData,
  TableDefinitionData,
  ThemeColorsData,
  ValidationReport,
  ViewMode,
  WorkbookData,
  WorkbookViewData,
} from './types.js';
import { ensureInitialized, wasmRepair, wasmValidate, wasmWrite } from './wasm-loader.js';

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
// Workbook
// ---------------------------------------------------------------------------

/** Represents an Excel workbook containing sheets, styles, and metadata. */
export class Workbook {
  private readonly data: WorkbookData;
  private readonly imageAnchors = new Map<string, { anchor: ImageAnchor; imageIndex: number }[]>();
  private readonly imageRels = new Map<string, { rId: string; target: string }[]>();

  /** Create a new workbook, optionally seeded with existing data from a read operation. */
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

  /** Creates a new fluent style builder for constructing cell styles. */
  createStyle(): StyleBuilder {
    return new StyleBuilder();
  }

  /** Returns the worksheet with the given name, or `undefined` if not found. */
  getSheet(name: string): Worksheet | undefined {
    const sheet = this.data.sheets.find((s) => s.name === name);
    return sheet ? new Worksheet(sheet, this.data.styles) : undefined;
  }

  /** Returns the worksheet at the given 0-based index, or `undefined` if out of range. */
  getSheetByIndex(index: number): Worksheet | undefined {
    const sheet = this.data.sheets[index];
    return sheet ? new Worksheet(sheet, this.data.styles) : undefined;
  }

  /**
   * Adds a new empty sheet to the workbook and returns it.
   * The name is validated per ECMA-376 rules (max 31 chars, no special characters).
   * @throws If the name is invalid or a sheet with that name already exists.
   */
  addSheet(name: string): Worksheet {
    validateSheetName(name);
    if (this.data.sheets.some((s) => s.name === name)) {
      throw new Error(`Sheet "${name}" already exists`);
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
    return new Worksheet(sheetData, this.data.styles);
  }

  /**
   * Removes a sheet by name or 0-based index.
   * @returns `true` if the sheet was removed, `false` if not found.
   */
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

  /** Returns all defined names (named ranges) in the workbook. */
  get namedRanges(): readonly DefinedNameData[] {
    return this.data.definedNames ?? [];
  }

  /** Adds a named range. If `sheetId` is provided, the name is scoped to that sheet. */
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

  /** Returns the named range with the given name, or `undefined` if not found. */
  getNamedRange(name: string): DefinedNameData | undefined {
    return this.data.definedNames?.find((d) => d.name === name);
  }

  /**
   * Removes a named range by name.
   * @returns `true` if the named range was removed, `false` if not found.
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
    if (!this.data.definedNames) this.data.definedNames = [];
    const idx = this.data.definedNames.findIndex(
      (d) => d.name === '_xlnm.Print_Titles' && d.sheetId === sheetIndex,
    );
    if (value === null) {
      if (idx !== -1) this.data.definedNames.splice(idx, 1);
    } else if (idx !== -1) {
      this.data.definedNames[idx]!.value = value;
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
    if (!this.data.definedNames) this.data.definedNames = [];
    const idx = this.data.definedNames.findIndex(
      (d) => d.name === '_xlnm.Print_Area' && d.sheetId === sheetIndex,
    );
    if (value === null) {
      if (idx !== -1) this.data.definedNames.splice(idx, 1);
    } else if (idx !== -1) {
      this.data.definedNames[idx]!.value = value;
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

  // --- Serialization ---

  /** Serializes the workbook to an XLSX `Uint8Array` via the WASM writer. */
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

  /**
   * Validate the workbook for OOXML compliance issues.
   * Returns a structured report with errors, warnings, and fix suggestions.
   */
  validate(): ValidationReport {
    ensureInitialized();
    return wasmValidate(this.data);
  }

  /**
   * Validate and auto-repair the workbook.
   * Fixes repairable issues (dangling style indices, missing default styles,
   * bad sheet names, missing theme, invalid metadata, row ordering).
   * Returns a new Workbook with repairs applied, plus a validation report.
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
   * @param sheetName - Target sheet name
   * @param anchor - Cell range to anchor the image to
   * @param imageBytes - Raw PNG image bytes
   * @param format - Image format extension (default: `'png'`)
   */
  addImage(
    sheetName: string,
    anchor: ImageAnchor,
    imageBytes: Uint8Array,
    format: 'png' | 'jpeg' | 'gif' = 'png',
  ): void {
    const sheetIndex = this.data.sheets.findIndex((s) => s.name === sheetName);
    if (sheetIndex === -1) throw new Error(`Sheet "${sheetName}" not found`);

    if (!this.data.preservedEntries) {
      this.data.preservedEntries = {};
    }

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
   * @param sheetName - Target sheet name
   * @param anchor - Cell range to anchor the barcode image to
   * @param value - Data to encode in the barcode
   * @param options - Barcode type, rendering, and sizing options
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

  /** Returns the raw internal `WorkbookData` for serialization or inspection. */
  toJSON(): WorkbookData {
    return this.data;
  }
}

// ---------------------------------------------------------------------------
// Worksheet
// ---------------------------------------------------------------------------

/** Represents a single worksheet within a workbook. */
export class Worksheet {
  private readonly data: SheetData;
  /** @internal */ readonly styles: StylesData | undefined;

  /** Wraps an existing sheet data object. Typically obtained via `Workbook.getSheet()` or `Workbook.addSheet()`. */
  constructor(data: SheetData, styles?: StylesData) {
    this.data = data;
    this.styles = styles;
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

  /** Sets the width of a single 1-based column, creating or updating its definition. */
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

  /** Adds a merged cell range (e.g., `"A1:C3"`). Duplicates are ignored. */
  addMergeCell(range: string): void {
    if (!this.data.worksheet.mergeCells.includes(range)) {
      this.data.worksheet.mergeCells.push(range);
    }
  }

  /**
   * Removes a merged cell range.
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

  /** Adds a hyperlink to a cell. The `location` is a URL or internal reference (e.g., `"Sheet2!A1"`). */
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

  /**
   * Removes the hyperlink on the given cell reference.
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
  get paneSelections(): PaneSelectionData[] {
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
    return (this.data.worksheet.sheetView?.view as ViewMode) ?? 'normal';
  }

  /** Sets the view mode. Creates a sheet view if none exists. */
  set viewMode(mode: ViewMode) {
    if (!this.data.worksheet.sheetView) {
      this.data.worksheet.sheetView = {};
    }
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

  /** Adds a data validation rule to the given cell range reference. */
  addValidation(ref: string, rule: Omit<DataValidationData, 'sqref'>): void {
    if (!this.data.worksheet.dataValidations) {
      this.data.worksheet.dataValidations = [];
    }
    this.data.worksheet.dataValidations.push({ sqref: ref, ...rule });
  }

  /**
   * Removes the data validation rule for the given cell range reference.
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

  /** Adds a comment to the given cell reference. */
  addComment(cellRef: string, author: string, text: string): void {
    if (!this.data.worksheet.comments) {
      this.data.worksheet.comments = [];
    }
    this.data.worksheet.comments.push({ cellRef, author, text });
  }

  /**
   * Removes the comment on the given cell reference.
   * @returns `true` if the comment was removed, `false` if not found.
   */
  removeComment(cellRef: string): boolean {
    if (!this.data.worksheet.comments) return false;
    const idx = this.data.worksheet.comments.findIndex((c) => c.cellRef === cellRef);
    if (idx === -1) return false;
    this.data.worksheet.comments.splice(idx, 1);
    return true;
  }

  // --- Page margins ---

  /** Returns the page margins (top, bottom, left, right, header, footer), or `null` if unset. */
  get pageMargins(): PageMarginsData | null {
    return this.data.worksheet.pageSetup?.margins ?? null;
  }

  /** Sets or clears the page margins. Creates a page setup object if needed. */
  set pageMargins(margins: PageMarginsData | null) {
    if (!this.data.worksheet.pageSetup) {
      this.data.worksheet.pageSetup = {};
    }
    this.data.worksheet.pageSetup.margins = margins;
  }

  // --- Cell access ---

  /**
   * Returns the cell at the given A1-style reference (e.g., `"B3"`), creating the row and cell if needed.
   * The returned `Cell` is a live wrapper -- mutations are reflected in the underlying sheet data.
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

  /** Sets the height (in points) of the given 1-based row, creating it if needed. */
  setRowHeight(rowIndex: number, height: number): void {
    const row = this.ensureRow(rowIndex);
    row.height = height;
  }

  /** Sets the visibility of the given 1-based row, creating it if needed. */
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

  /** Returns the table with the given display name, or `undefined` if not found. */
  getTable(displayName: string): TableDefinitionData | undefined {
    return this.data.worksheet.tables?.find((t) => t.displayName === displayName);
  }

  /** Adds a table definition to the sheet. */
  addTable(table: TableDefinitionData): void {
    if (!this.data.worksheet.tables) {
      this.data.worksheet.tables = [];
    }
    this.data.worksheet.tables.push(table);
  }

  /**
   * Removes the table with the given display name.
   * @returns `true` if the table was removed, `false` if not found.
   */
  removeTable(displayName: string): boolean {
    if (!this.data.worksheet.tables) return false;
    const idx = this.data.worksheet.tables.findIndex((t) => t.displayName === displayName);
    if (idx === -1) return false;
    this.data.worksheet.tables.splice(idx, 1);
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
      const existing = this.data.worksheet.columns.find(
        (c) => c.min === col && c.max === col,
      );
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
      const existing = this.data.worksheet.columns.find(
        (c) => c.min === col && c.max === col,
      );
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
}

// ---------------------------------------------------------------------------
// Cell
// ---------------------------------------------------------------------------

/** Represents a single cell within a worksheet row. Mutations are applied directly to the underlying data. */
export class Cell {
  private readonly data: CellData;
  private readonly styles: StylesData | undefined;

  /** Wraps an existing cell data object. Typically obtained via `Worksheet.cell()`. */
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
    throw new Error('Sheet name must not be empty');
  }
  if (name.length > 31) {
    throw new Error(`Sheet name must be 31 characters or fewer (got ${name.length})`);
  }
  if (INVALID_SHEET_CHARS.test(name)) {
    throw new Error(`Sheet name contains invalid characters: \\ / * ? [ ] :`);
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new Error('Sheet name must not start or end with an apostrophe');
  }
}
