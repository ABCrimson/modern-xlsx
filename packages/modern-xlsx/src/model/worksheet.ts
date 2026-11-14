import { ChartBuilder } from '../authoring/chart-builder.js';
import { PivotTableBuilder } from '../authoring/pivot-builder.js';
import { validateChartData } from '../authoring/validate-chart.js';
import {
  type CellRange,
  columnIndexOrNull,
  columnToLetter,
  decodeCellRef,
  decodeRange,
  MAX_COL,
  MAX_ROW,
} from '../format/cell-ref.js';
import { quoteSheetName } from '../formula/serializer.js';
import type { Token, TokenType } from '../formula/tokenizer.js';
import { tokenize } from '../formula/tokenizer.js';
import { Cell } from './cell.js';
import { COMMENT_NOT_FOUND, INVALID_ARGUMENT, ModernXlsxError } from './errors.js';
import { excelGuid } from './guid.js';
import { nameKey, referenceNameProblem, workbookNameIndexes } from './name-rules.js';
import type {
  AutoFilterData,
  CellData,
  ChartType,
  ColumnInfo,
  CommentData,
  DataValidationData,
  DefinedNameData,
  FrozenPane,
  HeaderFooterData,
  HyperlinkData,
  OutlinePropertiesData,
  PageBreakData,
  PageBreaksData,
  PageMarginsData,
  PageSetupData,
  PaneSelectionData,
  PersonData,
  PivotTableData,
  RowData,
  SheetData,
  SheetProtectionData,
  SheetState,
  SheetViewData,
  SlicerData,
  SparklineGroupData,
  SplitPaneData,
  StylesData,
  TableColumnData,
  TableDefinitionData,
  ThreadedCommentData,
  TimelineData,
  ViewMode,
  WorkbookData,
  WorksheetChartData,
} from './types.js';
import type { Workbook } from './workbook.js';

/**
 * A URI scheme prefix (RFC 3986 §3.1). `addHyperlink` sends a matching value to the
 * external `target` unless it is an unqualified A1 range (`UNQUALIFIED_A1_RANGE`); the
 * Rust writer applies the same rule to a target-less `location`. A sheet-qualified
 * reference such as `Sheet2!A1` never matches.
 */
const URI_SCHEME = /^[A-Za-z][A-Za-z0-9+.-]*:/;

/**
 * A same-sheet cell range (`A1:B2`, `A1:$B$2`) or column range (`B:D`): its first
 * token fits `URI_SCHEME`, but it is an in-workbook `location`.
 */
const UNQUALIFIED_A1_RANGE =
  /^(?:[A-Za-z]{1,3}[0-9]+:[$]?[A-Za-z]{1,3}[$]?[0-9]+|[A-Za-z]{1,3}:[$]?[A-Za-z]{1,3})$/;

/** Rows whose cells are known to be in column order (checked once per row). */
const SORTED_ROWS = new WeakSet<RowData>();

/** 0-based column of an A1 reference's letters, or -1 when they do not parse. */
function refColumn(ref: string): number {
  let end = 0;
  while (end < ref.length) {
    const code = ref.charCodeAt(end) & ~0x20;
    if (code < 0x41 || code > 0x5a) break;
    end++;
  }
  return columnIndexOrNull(ref.slice(0, end)) ?? -1;
}

/** Stable-sorts a row's cells by column the first time the row is touched. */
function ensureSortedCells(row: RowData): void {
  if (SORTED_ROWS.has(row)) return;
  const keyed = row.cells.map((cell) => ({ cell, col: refColumn(cell.reference) }));
  if (keyed.some((k, i) => i > 0 && (keyed[i - 1]?.col ?? 0) > k.col)) {
    keyed.sort((a, b) => a.col - b.col);
    row.cells.splice(0, row.cells.length, ...keyed.map((k) => k.cell));
  }
  SORTED_ROWS.add(row);
}

/** Binary-searches a column-sorted row: the cell's index, or `-(insertionPoint + 1)`. */
function searchCell(cells: readonly CellData[], col: number): number {
  let lo = 0;
  let hi = cells.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    const c = refColumn(cells[mid]?.reference ?? '');
    if (c === col) return mid;
    if (c < col) lo = mid + 1;
    else hi = mid - 1;
  }
  return -(lo + 1);
}

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

/** Throws `INVALID_ARGUMENT` unless `row` is a 1-based row index (an integer from 1 to 1048576). */
function checkRowIndex(row: number): void {
  if (!(Number.isInteger(row) && row >= 1 && row <= MAX_ROW + 1)) {
    throw new ModernXlsxError(
      INVALID_ARGUMENT,
      `Row index must be an integer from 1 to ${MAX_ROW + 1} (got ${row})`,
    );
  }
}

/** Throws `INVALID_ARGUMENT` unless `col` is a 1-based column index (an integer from 1 to 16384). */
function checkColumnIndex(col: number): void {
  if (!(Number.isInteger(col) && col >= 1 && col <= MAX_COL + 1)) {
    throw new ModernXlsxError(
      INVALID_ARGUMENT,
      `Column index must be an integer from 1 to ${MAX_COL + 1} (got ${col})`,
    );
  }
}

/**
 * Removes `list[index]` when `index` is an in-range integer.
 *
 * @returns `true` if an element was removed, `false` if `list` is absent or
 *   `index` is not an in-range integer.
 */
function removeAt<T>(list: T[] | undefined, index: number): boolean {
  if (!list || !Number.isInteger(index) || index < 0 || index >= list.length) return false;
  list.splice(index, 1);
  return true;
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
  /** The owning workbook, through which the `name` setter renames. */
  readonly #workbook: Workbook | undefined;

  /**
   * Wraps an existing sheet data object. Typically obtained via
   * `Workbook.getSheet()` or `Workbook.addSheet()`.
   *
   * @param data - The underlying sheet data model.
   * @param styles - Optional shared styles from the parent workbook.
   * @param workbookData - Optional parent workbook data for cross-sheet operations.
   * @param workbook - Internal: the owning workbook, which passes itself so the
   *   `name` setter renames through {@link Workbook.renameSheet}.
   */
  constructor(
    data: SheetData,
    styles?: StylesData,
    workbookData?: WorkbookData,
    workbook?: Workbook,
  ) {
    this.data = data;
    this.styles = styles;
    this.workbookData = workbookData;
    this.#workbook = workbook;
  }

  /** Returns the sheet tab name. */
  get name(): string {
    return this.data.name;
  }

  /**
   * Renames the sheet tab, validated against Excel's sheet-name rules: 1-31
   * UTF-16 code units, none of `\ / * ? [ ] :` and no control character
   * (Unicode Cc), no leading or trailing apostrophe, and not `History` in any case.
   *
   * On a sheet that belongs to a workbook this is
   * {@link Workbook.renameSheet}: the name must be unique ignoring case and
   * canonical equivalence, and defined names that refer to the sheet are
   * rewritten. A standalone worksheet only validates the name.
   *
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if the name is empty,
   *   exceeds 31 UTF-16 code units, contains forbidden characters, starts or ends
   *   with an apostrophe, is the reserved name `History` (any case), or is already used by another sheet of the
   *   workbook (any case or normal form); nothing is changed then.
   */
  set name(value: string) {
    const index = this.workbookData?.sheets.indexOf(this.data) ?? -1;
    if (this.#workbook && index !== -1) {
      this.#workbook.renameSheet(index, value);
      return;
    }
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
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if `col` is not an
   *   integer from 1 to 16384.
   *
   * @example
   * ```ts
   * ws.setColumnWidth(1, 20); // column A = 20 chars wide
   * ```
   */
  setColumnWidth(col: number, width: number): void {
    const column = this.#ensureSingleColumn(col);
    column.width = width;
    column.customWidth = true;
  }

  /**
   * Returns the single-column definition for the 1-based `col`, creating it in
   * sorted position when no span covers the column and splitting the covering
   * span into up to three copies (each keeping every attribute) otherwise.
   *
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if `col` is not an
   *   integer from 1 to 16384.
   */
  #ensureSingleColumn(col: number): ColumnInfo {
    checkColumnIndex(col);
    const columns = this.data.worksheet.columns;
    const at = columns.findIndex((c) => c.min <= col && col <= c.max);
    const span = columns[at];
    if (!span) {
      const created: ColumnInfo = {
        min: col,
        max: col,
        width: 8.43,
        hidden: false,
        customWidth: false,
      };
      const next = columns.findIndex((c) => c.min > col);
      columns.splice(next === -1 ? columns.length : next, 0, created);
      return created;
    }
    if (span.min === col && span.max === col) return span;
    const middle: ColumnInfo = { ...span, min: col, max: col };
    const pieces: ColumnInfo[] = [];
    if (span.min < col) pieces.push({ ...span, max: col - 1 });
    pieces.push(middle);
    if (span.max > col) pieces.push({ ...span, min: col + 1 });
    columns.splice(at, 1, ...pieces);
    return middle;
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
   * Adds a hyperlink to a cell. A value that starts with a URI scheme
   * (`/^[A-Za-z][A-Za-z0-9+.-]*:/`: `https:`, `mailto:`, `file:`, and also a drive
   * path such as `C:\book.xlsx`) is stored as the external `target`, written as a
   * sheet relationship with `TargetMode="External"`; anything else, including a
   * same-sheet range such as `'A1:B2'` or `'B:D'`, is an in-workbook `location`
   * such as `'Sheet2!A1'` or a defined name.
   *
   * @param cellRef - The cell reference to attach the hyperlink to (e.g., `'A1'`).
   * @param location - External URL, `mailto:` address or file, or an in-workbook
   *   reference (e.g., `'https://example.com'` or `'Sheet2!A1'`).
   * @param opts - Optional display text and tooltip.
   *
   * @example
   * ```ts
   * ws.addHyperlink('A1', 'https://example.com', {
   *   display: 'Visit Site',
   *   tooltip: 'Opens example.com',
   * });
   * ws.addHyperlink('A2', 'Sheet2!A1', { display: 'Go to Sheet2' });
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
      ...(!UNQUALIFIED_A1_RANGE.test(location) && URI_SCHEME.test(location)
        ? { target: location }
        : { location }),
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

  /**
   * Sets or clears sheet protection. Pass `null` to remove protection. Each flag is
   * written as given: `true` locks the action in Excel, `false` allows it (see
   * {@link SheetProtectionData}); set `plainPassword` to require a password.
   */
  set sheetProtection(protection: SheetProtectionData | null) {
    this.data.worksheet.sheetProtection = protection;
  }

  // --- Frozen pane ---

  /** Returns the frozen pane split position, or `null` if no panes are frozen. */
  get frozenPane(): FrozenPane | null {
    return this.data.worksheet.frozenPane;
  }

  /**
   * Freezes rows and/or columns at the given split position, or clears the freeze
   * with `null`. Setting a frozen pane clears any split pane and the per-pane selections.
   */
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

  /**
   * Sets a split pane. Setting a split pane clears any frozen pane. Pass `null` to
   * clear it, which also clears the per-pane selections.
   */
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

  /**
   * Returns the sheet tab colour hex string: the ARGB value stored in the file
   * (e.g. `"FFFF0000"`) on a workbook read from a file, or the string last set;
   * `null` when no tab colour is set or it has no RGB value (a theme or indexed
   * colour).
   */
  get tabColor(): string | null {
    return this.data.worksheet.tabColor ?? null;
  }

  /**
   * Sets the sheet tab colour as hex RGB (`RRGGBB`) or ARGB (`AARRGGBB`), with
   * or without a leading `#`; written as uppercase ARGB. Pass `null` to clear.
   */
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
   * @param rule - The validation rule: `validationType`, `operator`, `formula1`,
   *   `formula2`, `allowBlank`, `showErrorMessage`, `errorTitle` and
   *   `errorMessage` (each `null` when unused), plus the optional
   *   `showInputMessage`, `promptTitle` and `prompt`.
   *
   * @example
   * ```ts
   * ws.addValidation('B2:B100', {
   *   validationType: 'list',
   *   operator: null,
   *   formula1: '"Yes,No,Maybe"',
   *   formula2: null,
   *   allowBlank: true,
   *   showErrorMessage: true,
   *   errorTitle: null,
   *   errorMessage: null,
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
   * Adds a threaded comment to the given cell reference: the root of a new
   * thread. A cell holds one thread (Excel shows only the first thread on a
   * cell and drops the rest), so a cell that already has a thread root throws;
   * add to that thread with {@link Worksheet.replyToComment} instead.
   * Creates a person entry in the workbook if the author does not exist yet.
   *
   * The comment and person ids are upper-case GUIDs in braces
   * (`{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`), the form Excel requires. On
   * write the sheet also gets the legacy placeholder note Excel needs to show
   * the thread (author `tc={id}`, text `[Threaded comment]` ... `Comment:`
   * and the replies), with its hidden note shape. It is added to the file
   * only, not to {@link Worksheet.comments}, and only when the cell has no
   * note: a note already on the cell is written as it is, and Excel then
   * shows that note, not the thread. A read of the file returns the
   * placeholder in {@link Worksheet.comments}, as it does for a file Excel
   * saved, and later writes keep it as it is.
   *
   * The comment's `timestamp` is the current time as an ISO 8601 UTC string
   * (`new Date().toISOString()` form). The writer writes it as `dT` in the
   * form Excel reads (UTC, hundredths of a second, no `Z`); a read returns
   * `dT` as the file holds it.
   *
   * @param cell - The cell reference (e.g., `'A1'`); it is stored without `$`
   *   and in upper case (`'$d$4'` becomes `'D4'`), and cells are compared that
   *   way (`'a1'`, `'$A$1'` and `'A1'` are one cell).
   * @param text - The comment text: at most 32,767 UTF-16 code units (its
   *   `length`; an astral character such as an emoji counts as two). Excel
   *   refuses a file with a longer threaded comment.
   * @param author - The display name of the author.
   * @returns The ID of the created comment, an upper-case GUID in braces.
   * @throws ModernXlsxError with code `INVALID_ARGUMENT`, leaving the worksheet
   *   unchanged, if the cell already has a thread root (the message names the
   *   cell and the root's ID to pass to `replyToComment`), if `text` is longer
   *   than 32,767 UTF-16 code units, or on a worksheet that was not obtained
   *   from a workbook.
   *
   * @example
   * ```ts
   * const id = ws.addThreadedComment('A1', 'Looks good!', 'Alice');
   * ws.replyToComment(id, 'Thanks!', 'Bob');
   * ```
   */
  addThreadedComment(cell: string, text: string, author: string): string {
    return this.#pushThreaded(cell, text, author);
  }

  /**
   * Replies to an existing threaded comment.
   * Creates a person entry in the workbook if the author does not exist yet.
   *
   * The reply's `parentId` is the thread root: a reply to a reply joins its
   * thread root, as Excel does (Excel refuses a file whose reply names
   * another reply as its parent).
   *
   * Excel shows a thread's replies in the order of their timestamps (`dT`,
   * which it keeps to the hundredth of a second). The reply's `timestamp`
   * is the current time as an ISO 8601 UTC string, or, when that is not at
   * least 10 ms after the latest timestamp already in the thread (replies
   * added within the same hundredth, or a thread stamped ahead of this
   * clock), that latest timestamp plus 10 ms, so replies show in the order
   * they were added. A timestamp without a zone, as Excel writes `dT`, is UTC.
   *
   * @param commentId - The ID of the comment to reply to (a thread root or a
   *   reply), exactly as {@link Worksheet.addThreadedComment} or
   *   {@link Worksheet.replyToComment} returned it or the file holds it.
   * @param text - The reply text: at most 32,767 UTF-16 code units (its
   *   `length`; an astral character such as an emoji counts as two). Excel
   *   refuses a file with a longer threaded comment.
   * @param author - The display name of the reply author.
   * @returns The ID of the reply comment, an upper-case GUID in braces.
   * @throws ModernXlsxError with code `COMMENT_NOT_FOUND` if no threaded comment on
   *   this sheet has the ID, or `INVALID_ARGUMENT`, leaving the worksheet
   *   unchanged, if `text` is longer than 32,767 UTF-16 code units or on a
   *   worksheet that was not obtained from a workbook.
   */
  replyToComment(commentId: string, text: string, author: string): string {
    const parent = this.#threadedById(commentId);
    if (!parent) {
      throw new ModernXlsxError(COMMENT_NOT_FOUND, `Comment ${commentId} not found`);
    }
    return this.#pushThreaded(parent.refCell, text, author, parent.parentId ?? parent.id);
  }

  /**
   * Appends a threaded comment by `author` (a reply when `parentId` is given)
   * and returns its new ID, in the braced upper-case form. The cell reference
   * is stored without `$` and in upper case (`'$d$4'` becomes `'D4'`): Excel
   * hides a thread whose `ref` holds `$`.
   *
   * @throws ModernXlsxError with code `INVALID_ARGUMENT`, before anything
   *   changes, on a worksheet that was not obtained from a workbook, for a
   *   `text` longer than {@link MAX_THREADED_COMMENT_LENGTH}, or for a root
   *   (no `parentId`) on a cell that already has a thread root.
   */
  #pushThreaded(refCell: string, text: string, author: string, parentId?: string): string {
    const wb = this.workbookData;
    if (!wb) {
      throw new ModernXlsxError(
        INVALID_ARGUMENT,
        'Worksheet must be created via Workbook.addSheet() or Workbook.getSheet() to use threaded comments',
      );
    }
    if (text.length > MAX_THREADED_COMMENT_LENGTH) {
      throw new ModernXlsxError(
        INVALID_ARGUMENT,
        `Threaded comment text is ${text.length} UTF-16 code units long: Excel refuses a threaded comment longer than ${MAX_THREADED_COMMENT_LENGTH} UTF-16 code units`,
      );
    }
    const ref = threadCellKey(refCell);
    if (parentId === undefined) {
      const root = this.#threadRootOn(ref);
      if (root) {
        throw new ModernXlsxError(
          INVALID_ARGUMENT,
          `Cell ${ref} already has a thread (root ${root.id}), and a cell holds one thread: add to it with replyToComment('${root.id}', text, author)`,
        );
      }
    }
    const person = this.#personFor(wb, author);
    this.data.worksheet.threadedComments ??= [];
    const threads = this.data.worksheet.threadedComments;
    const index = threadIndex(threads);
    const id = excelGuid();
    const threadId = parentId ?? id;
    // At least one step after the thread's latest stamp: Excel orders replies by dT.
    const at = Math.max(
      Date.now(),
      (index.latest.get(threadId) ?? Number.NEGATIVE_INFINITY) + THREAD_STAMP_STEP_MS,
    );
    const comment: ThreadedCommentData = {
      id,
      refCell: ref,
      personId: person.id,
      text,
      timestamp: new Date(at).toISOString(),
      ...(parentId !== undefined && { parentId }),
    };
    indexThreaded(index, comment, threads.push(comment) - 1);
    index.length = threads.length;
    return id;
  }

  /**
   * The thread root on cell `ref` (a {@link threadCellKey}), or `undefined`.
   * An indexed hit is checked against the array, and the index rebuilt when
   * the array changed under it.
   */
  #threadRootOn(ref: string): ThreadedCommentData | undefined {
    const threads = this.data.worksheet.threadedComments;
    if (!threads) return undefined;
    const isRoot = (c: ThreadedCommentData | undefined): c is ThreadedCommentData =>
      c !== undefined && c.parentId === undefined && threadCellKey(c.refCell) === ref;
    const position = threadIndex(threads).roots.get(ref);
    if (position === undefined) return undefined;
    const hit = threads[position];
    if (isRoot(hit)) return hit;
    const again = threadIndex(threads, true).roots.get(ref);
    return again === undefined ? undefined : threads[again];
  }

  /**
   * The threaded comment with this `id`, or `undefined`. An indexed hit is
   * checked against the array; a miss or a stale hit rebuilds the index
   * once, so an id the caller put into the array is always found.
   */
  #threadedById(id: string): ThreadedCommentData | undefined {
    const threads = this.data.worksheet.threadedComments;
    if (!threads) return undefined;
    const position = threadIndex(threads).ids.get(id);
    const hit = position === undefined ? undefined : threads[position];
    if (hit?.id === id) return hit;
    const again = threadIndex(threads, true).ids.get(id);
    return again === undefined ? undefined : threads[again];
  }

  /** Finds the workbook person with this display name, creating it when absent. */
  #personFor(wb: WorkbookData, author: string): PersonData {
    wb.persons ??= [];
    let person = wb.persons.find((p) => p.displayName === author);
    if (!person) {
      person = { id: excelGuid(), displayName: author };
      wb.persons.push(person);
    }
    return person;
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
   * @throws ModernXlsxError with code `INVALID_CELL_REF` if `ref` is not an
   *   uppercase A1 reference within `A1:XFD1048576`.
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

    const row = this.ensureRow(rowIndex);

    ensureSortedCells(row);
    const found = searchCell(row.cells, decoded.col);
    let cellData = found >= 0 ? row.cells[found] : undefined;
    if (!cellData) {
      cellData = {
        reference: normalRef,
        cellType: 'number',
        value: null,
        formula: null,
        styleIndex: null,
      };
      row.cells.splice(-found - 1, 0, cellData);
    }

    return new Cell(cellData, this.styles, () => this.workbookData?.dateSystem ?? 'date1900');
  }

  // --- Row utilities ---

  /**
   * Sets the height (in points) of the given 1-based row, creating it if needed.
   * The height is written as a custom height (`customHeight="1"`), so Excel
   * shows it instead of sizing the row to its content.
   *
   * @param rowIndex - The 1-based row index.
   * @param height - Row height in points (e.g., 24 for double-height).
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if `rowIndex` is not an
   *   integer from 1 to 1048576.
   *
   * @example
   * ```ts
   * ws.setRowHeight(1, 30); // set header row to 30pt
   * ```
   */
  setRowHeight(rowIndex: number, height: number): void {
    const row = this.ensureRow(rowIndex);
    row.height = height;
    row.customHeight = true;
  }

  /**
   * Sets the visibility of the given 1-based row, creating it if needed.
   *
   * @param rowIndex - The 1-based row index.
   * @param hidden - `true` to hide the row, `false` to show it.
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if `rowIndex` is not an
   *   integer from 1 to 1048576.
   */
  setRowHidden(rowIndex: number, hidden: boolean): void {
    const row = this.ensureRow(rowIndex);
    row.hidden = hidden;
  }

  /**
   * Find or create a row at the given 1-based index.
   *
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if `rowIndex` is not an
   *   integer from 1 to 1048576.
   */
  private ensureRow(rowIndex: number): RowData {
    checkRowIndex(rowIndex);
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
   * The `displayName` must be a name Excel opens: 1 to 255 UTF-16 code units;
   * no space, tab, line break, DEL or ASCII punctuation other than `_`, `.`,
   * `\` and `?`; not starting with a digit, `.`, `?` or a combining mark
   * (Unicode M; Excel refused U+0301 + `abc` and U+093F + `x`, while
   * `a` + U+0301 + `bc` opens; a leading Thai mark such as U+0E34 + `a`,
   * which Excel opens, is refused too); not `TRUE` or `FALSE`;
   * not a cell reference in the grid (`T2`, `xfd1`; `XFE1` and `A0` are fine);
   * and not an R1C1 reference or a name that starts with one (`R`, `C`, `RC`,
   * `R1x`, `C1_`, `RC2x`; a `.` right after the number, as in `R1.a`, is fine).
   * Case is ignored. It must also differ, ignoring case and canonical
   * equivalence, from the display name of every other table and from every
   * defined name in the workbook. An explicit `name` may be at most 255 code
   * units and must differ, compared the same way, from every other table's
   * explicit `name`.
   *
   * With `totalsRowCount` 1, Excel keeps a table's totals as ordinary cells in
   * the last row of its range, so this method also writes them: for each column
   * with a `totalsRowFunction` (`'sum'`, `'average'`, `'count'`, `'countNums'`,
   * `'max'`, `'min'`, `'stdDev'` or `'var'`) the `SUBTOTAL` formula Excel writes
   * (`SUBTOTAL(109,Table1[Amount])` for `'sum'`), for `'custom'` the column's
   * `totalsRowFormula`, and for a column with a `totalsRowLabel` and no function
   * the label text. The formulas have no cached value, so Excel computes them
   * when it opens the file. A table read from a file is not changed.
   *
   * A totals cell that already holds a value or a formula is kept as it is,
   * and Excel opens the file only if the kept cell agrees with its column
   * (Excel 365 probes):
   * - a `SUBTOTAL` column: a plain value (number, text, boolean), or the
   *   formula above, which Excel also takes without the table name
   *   (`SUBTOTAL(109,[Amount])`) and with the function or column name in
   *   another case; any other formula is refused, the table name in another
   *   case and a space inside the parentheses included;
   * - a `'custom'` column with a `totalsRowFormula`: any formula; a value or an
   *   empty cell is refused (without a `totalsRowFormula`, anything opens);
   * - a column with a `totalsRowLabel` and no function (or `'none'`): the label
   *   text, exactly, or nothing; other text, a number, a boolean or a formula
   *   is refused;
   * - a column with neither: text or nothing; a number, a boolean or a formula is refused.
   *
   * `Workbook.validate()` reports a totals cell Excel refuses as an error.
   *
   * A column's `calculatedColumnFormula` and `totalsRowFormula` are stored in
   * the form Excel writes in a file, which has no `@` this-row shorthand:
   * Excel refuses to open a table definition holding `[@Qty]`,
   * `[@[Unit Cost]]`, `Table1[@Qty]` or `[@]` (outside a string literal);
   * write `Table1[[#This Row],[Qty]]`, `Table1[[#This Row],[Unit Cost]]` or
   * `Table1[#This Row]` there. `[Qty]` and `[[#This Row],[Qty]]` without the
   * table name open. A `[@...]` that names a column of its table, ignoring
   * case, is a whole-column reference and opens: Excel writes the column
   * `@Home` of table `AtT` as `AtT[@Home]` (the names compared after their
   * `_xHHHH_` escapes are decoded). A `[@...]` of a table in another
   * workbook, which Excel stores as `[1]!Ext[@Home]`, is not checked. The
   * formulas are kept as given.
   *
   * @param table - The table definition to add (range, columns, style, etc.).
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if the display name or
   *   the `name` breaks one of the rules above (duplicates are found ignoring
   *   case and canonical equivalence: `café` with U+00E9 repeats `cafe` +
   *   U+0301), or a `calculatedColumnFormula`
   *   or `totalsRowFormula` uses the `@` shorthand (the message gives the form
   *   to write instead); the sheet is unchanged then.
   */
  addTable(table: TableDefinitionData): void {
    const key = this.#checkTableNames(table);
    const { tables } = workbookNameIndexes(this.workbookData ?? this.data);
    checkTableFormulas(table, key, (other) => tables.first(other));
    this.data.worksheet.tables ??= [];
    this.data.worksheet.tables.push(table);
    this.#writeTotalsRow(table);
  }

  /**
   * Throws unless `table`'s names are ones Excel opens beside the workbook's
   * tables and defined names; returns the {@link nameKey} of its display
   * name. The stored names' keys come from the workbook's name indexes
   * (`workbookNameIndexes`), brought up to date here, so the check computes
   * the keys of `table`'s own names alone.
   */
  #checkTableNames(table: TableDefinitionData): string {
    const invalid = (message: string) => new ModernXlsxError(INVALID_ARGUMENT, message);
    const problem = tableNameProblem(table.displayName);
    if (problem) {
      throw invalid(`Invalid table name ${JSON.stringify(table.displayName)}: ${problem}`);
    }
    if (table.name != null && table.name.length > MAX_TABLE_NAME_LENGTH) {
      throw invalid(
        `Invalid table name attribute ${JSON.stringify(table.name)}: it is longer than 255 characters`,
      );
    }
    const key = nameKey(table.displayName);
    // Excel opened an empty `name` (it then names the table itself), so only a set one can clash.
    const attributeKey = table.name ? nameKey(table.name) : undefined;
    const sheets = this.workbookData?.sheets ?? [this.data];
    const lists = sheets.map((sheet) => sheet.worksheet.tables);
    const indexes = workbookNameIndexes(this.workbookData ?? this.data);
    const tables = indexes.tables.sync(lists);
    tables.hint(table.displayName, key);
    // The first table, in sheet order, that clashes by either name.
    const byDisplay = tables.positionsOf(key)[0] ?? Number.POSITIVE_INFINITY;
    let byAttribute = Number.POSITIVE_INFINITY;
    if (attributeKey !== undefined && table.name) {
      const names = indexes.tableNames.sync(lists);
      names.hint(table.name, attributeKey);
      byAttribute = names.positionsOf(attributeKey)[0] ?? Number.POSITIVE_INFINITY;
    }
    if (byDisplay <= byAttribute && byDisplay !== Number.POSITIVE_INFINITY) {
      const other = tables.itemAt(byDisplay);
      throw invalid(
        `Table name ${JSON.stringify(table.displayName)} is already used by table ${JSON.stringify(other?.displayName)} (table names ignore case and canonical equivalence)`,
      );
    }
    if (byAttribute !== Number.POSITIVE_INFINITY) {
      const other = indexes.tableNames.itemAt(byAttribute);
      throw invalid(
        `Table name attribute ${JSON.stringify(table.name)} is already used by name attribute ${JSON.stringify(other?.name)} of table ${JSON.stringify(other?.displayName)} (table names ignore case and canonical equivalence)`,
      );
    }
    const defined = this.workbookData
      ? indexes.definedNames.sync([this.workbookData.definedNames]).first(key)
      : undefined;
    if (defined) {
      throw invalid(
        `Table name ${JSON.stringify(table.displayName)} is already used by defined name ${JSON.stringify(defined.name)} (names ignore case and canonical equivalence)`,
      );
    }
    return key;
  }

  /** Fills the empty cells of `table`'s totals row from its columns; see {@link Worksheet.addTable}. */
  #writeTotalsRow(table: TableDefinitionData): void {
    if (table.totalsRowCount < 1) return;
    const range = tryDecodeRange(table.ref);
    if (!range) return;
    const row = range.end.row + 1;
    table.columns.forEach((column, i) => {
      const col = range.start.col + i;
      if (col > range.end.col) return;
      const content = totalsCellContent(table.displayName, column);
      if (!content) return;
      const cell = this.cell(`${columnToLetter(col)}${row}`);
      if (cell.value !== null || cell.formula !== null) return;
      if ('formula' in content) cell.formula = content.formula;
      else cell.value = content.label;
    });
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
   * Adds a pivot table definition to this sheet. Its `cacheId` must name a pivot cache of
   * the workbook, such as one added with `Workbook.addPivotCache` (the write fails
   * otherwise). A row, column or page field with empty `items` is completed from that cache
   * on write, and the cache then refreshes when Excel opens the file (see `PivotTableData`).
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
   * @returns `true` if the pivot table was removed, `false` if `index` is not an
   *   in-range integer.
   */
  removePivotTable(index: number): boolean {
    return removeAt(this.data.worksheet.pivotTables, index);
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
   * @returns `true` if the slicer was removed, `false` if `index` is not an
   *   in-range integer.
   */
  removeSlicer(index: number): boolean {
    return removeAt(this.data.worksheet.slicers, index);
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
   * @returns `true` if the timeline was removed, `false` if `index` is not an
   *   in-range integer.
   */
  removeTimeline(index: number): boolean {
    return removeAt(this.data.worksheet.timelines, index);
  }

  /**
   * Adds a chart to this worksheet using a builder callback.
   *
   * @param type - The chart type (e.g., `'bar'`, `'line'`, `'pie'`, `'scatter'`).
   * @param configure - A callback that receives a ChartBuilder for fluent configuration.
   * @throws {RangeError} When the built chart fails validateChartData (a value outside its
   *   ECMA-376 range, a colour that is not RRGGBB, two series sharing an order, an anchor
   *   that ends before it starts, an empty series reference) or its secondary chart has a
   *   secondary chart of its own.
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
   * @throws {RangeError} When `chart` fails validateChartData (a value outside its ECMA-376
   *   range, a colour that is not RRGGBB, two series sharing an order, an anchor that ends
   *   before it starts, an empty series reference).
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
   * @returns `true` if removed, `false` if `index` is not an in-range integer.
   */
  removeChart(index: number): boolean {
    return removeAt(this.data.worksheet.charts, index);
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
   * Adds a manual row page break with `id` set to `row` and `max` 16383. ECMA-376
   * places a row break above the zero-based row `id`, so the break falls below the
   * 1-based row `row`.
   * @param row The break `id`: the 1-based number of the last row above the break.
   */
  addRowBreak(row: number): void {
    this.data.worksheet.pageBreaks ??= {};
    this.data.worksheet.pageBreaks.rowBreaks ??= [];
    const brk: PageBreakData = { id: row, max: 16383, man: true };
    this.data.worksheet.pageBreaks.rowBreaks.push(brk);
  }

  /**
   * Adds a manual column page break with `id` set to `col` and `max` 1048575. ECMA-376
   * places a column break left of the zero-based column `id`, so the break falls to the
   * right of the 1-based column `col`.
   * @param col The break `id`: the 1-based number of the last column left of the break.
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
   * Does nothing on a worksheet that does not belong to a workbook.
   *
   * Excel refuses a file whose print area has no sheet qualifier, so every
   * comma-separated area of `ref` that lacks one (`"A1:B2"`) is stored qualified
   * with this sheet's name, quoted when needed (`"'My Sheet'!A1:B2"`), and
   * otherwise as written; a qualified area (`"Sheet1!$A$1:$D$50"`) is stored as it is.
   * @param ref The reference or comma-separated list of references (cell, range,
   *   column span or row span). Pass `null` to clear.
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if `ref` is not a
   *   reference or a comma-separated list of references.
   */
  setPrintArea(ref: string | null): void {
    if (!this.workbookData) return;
    const sheetIndex = this.workbookData.sheets.indexOf(this.data);
    if (sheetIndex === -1) return;
    upsertBuiltinName(
      this.workbookData,
      '_xlnm.Print_Area',
      sheetIndex,
      ref === null ? undefined : qualifyBuiltinNameValue('_xlnm.Print_Area', this.data.name, ref),
    );
  }

  /**
   * Gets the print area for this sheet.
   * @returns The defined name value (e.g., `"Sheet1!$A$1:$D$50"`) or `null`.
   */
  getPrintArea(): string | null {
    if (!this.workbookData) return null;
    const sheetIndex = this.workbookData.sheets.indexOf(this.data);
    if (sheetIndex === -1) return null;
    return getBuiltinName(this.workbookData.definedNames, '_xlnm.Print_Area', sheetIndex);
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
   * Groups rows by setting their outline level (1-7), creating missing rows.
   * @param startRow 1-based start row index (inclusive).
   * @param endRow 1-based end row index (inclusive).
   * @param level Outline level, clamped to 0-7; 0 clears the level. Defaults to 1.
   * @throws ModernXlsxError with code `INVALID_ARGUMENT`, before any row changes,
   *   if `startRow` or `endRow` is not an integer from 1 to 1048576.
   */
  groupRows(startRow: number, endRow: number, level = 1): void {
    checkRowIndex(startRow);
    checkRowIndex(endRow);
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
   * Groups columns by setting their outline level (1-7), creating missing column definitions.
   * @param startCol 1-based start column index (inclusive).
   * @param endCol 1-based end column index (inclusive).
   * @param level Outline level, clamped to 0-7; 0 clears the level. Defaults to 1.
   * @throws ModernXlsxError with code `INVALID_ARGUMENT`, before any column
   *   changes, if `startCol` or `endCol` is not an integer from 1 to 16384.
   */
  groupColumns(startCol: number, endCol: number, level = 1): void {
    checkColumnIndex(startCol);
    checkColumnIndex(endCol);
    const clamped = Math.min(Math.max(level, 0), 7);
    const clampedLevel = clamped > 0 ? clamped : null;
    for (let col = startCol; col <= endCol; col++) {
      this.#ensureSingleColumn(col).outlineLevel = clampedLevel;
    }
  }

  /**
   * Ungroups columns by clearing their outline level.
   * @param startCol 1-based start column index (inclusive).
   * @param endCol 1-based end column index (inclusive).
   */
  ungroupColumns(startCol: number, endCol: number): void {
    for (let col = startCol; col <= endCol; col++) {
      if (this.data.worksheet.columns.some((c) => c.min <= col && col <= c.max)) {
        const column = this.#ensureSingleColumn(col);
        column.outlineLevel = null;
        column.collapsed = false;
      }
    }
  }

  /**
   * Collapses grouped rows: hides each row in the range that has an outline level
   * and marks the summary row `collapsed`. The summary row is the row after the
   * range, or the row before it when `outlineProperties.summaryBelow` is `false`.
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
    if (summaryIdx >= 1 && summaryIdx <= MAX_ROW + 1) {
      const summary = this.ensureRow(summaryIdx);
      summary.collapsed = true;
    }
  }

  /**
   * Expands collapsed grouped rows: shows each row in the range that has an outline
   * level and clears `collapsed` on the summary row (see {@link Worksheet.collapseRows}).
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
// Helpers shared with Workbook (internal: exports-shared.ts does not re-export them)
// ---------------------------------------------------------------------------

const INVALID_SHEET_CHARS = /[\\/*?[\]:\p{Cc}]/u;

/**
 * Validates a sheet name against Excel's sheet-name rules: 1-31 UTF-16 code
 * units (`name.length`), none of `\ / * ? [ ] :` and no control character
 * (Unicode Cc), no leading or trailing apostrophe, and not `History` in any case.
 *
 * @throws ModernXlsxError with code `INVALID_ARGUMENT` when a rule is broken.
 */
export function validateSheetName(name: string): void {
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
      'Sheet name contains invalid characters: \\ / * ? [ ] : or a control character',
    );
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new ModernXlsxError(
      INVALID_ARGUMENT,
      'Sheet name must not start or end with an apostrophe',
    );
  }
  if (name.toLowerCase() === 'history') {
    throw new ModernXlsxError(INVALID_ARGUMENT, "'History' is a reserved sheet name");
  }
}

/**
 * Longest threaded comment text Excel 16 opens, in UTF-16 code units (a
 * string's `length`): it refuses a file with a 32,768-unit root or reply, and
 * counts an astral character as two units (Excel 16 probed 2026-11-12).
 * The Rust `validate.rs` twin is `MAX_THREADED_COMMENT_TEXT`.
 */
const MAX_THREADED_COMMENT_LENGTH = 32_767;

/** The cell a threaded comment `refCell` names, as it is stored: without `$` and in upper case. */
function threadCellKey(refCell: string): string {
  return refCell.replaceAll('$', '').toUpperCase();
}

/**
 * The smallest step between two timestamps of a thread that Excel tells
 * apart: it keeps `dT` to the hundredth of a second (Excel 16 probed 2026-11-12).
 */
const THREAD_STAMP_STEP_MS = 10;

/** Lookups over one sheet's `threadedComments` array, by position in it. */
interface ThreadIndex {
  /** The array length the lookups were built or last extended for. */
  length: number;
  /** The {@link threadCellKey} of each thread root's cell: its position (the first root on a cell). */
  readonly roots: Map<string, number>;
  /** Each comment id: its position (the first comment with the id). */
  readonly ids: Map<string, number>;
  /** Each thread root id: the latest timestamp in its thread, in ms since the epoch. */
  readonly latest: Map<string, number>;
}

/**
 * The lookups of each `threadedComments` array, keyed by the array itself:
 * every Worksheet handle of a sheet (`getSheet()` makes one per call) shares
 * them, and a cloned or read sheet, whose array is another object, never does.
 */
const threadIndexes = new WeakMap<ThreadedCommentData[], ThreadIndex>();

/**
 * The lookups of `threads`, rebuilt in one pass when `rebuild` is set, when
 * there are none yet, or when the array's length is not the one they were
 * built for (the caller pushed, spliced or popped through the model). A
 * same-length edit in place (an element replaced, or a field of one changed)
 * is not seen here: callers check each hit against the array and rebuild when
 * it is stale, and an id lookup rebuilds on a miss. What only a full scan
 * would see waits for the next rebuild: a root such an edit put on a cell the
 * lookups have no root for, and a timestamp such an edit moved later.
 */
function threadIndex(threads: ThreadedCommentData[], rebuild = false): ThreadIndex {
  let index = threadIndexes.get(threads);
  if (index === undefined || rebuild || index.length !== threads.length) {
    const built: ThreadIndex = {
      length: threads.length,
      roots: new Map(),
      ids: new Map(),
      latest: new Map(),
    };
    for (const [position, c] of threads.entries()) indexThreaded(built, c, position);
    threadIndexes.set(threads, built);
    index = built;
  }
  return index;
}

/** Adds comment `c`, at `position` in its array, to `index`. */
function indexThreaded(index: ThreadIndex, c: ThreadedCommentData, position: number): void {
  if (!index.ids.has(c.id)) index.ids.set(c.id, position);
  if (c.parentId === undefined) {
    const key = threadCellKey(c.refCell);
    if (!index.roots.has(key)) index.roots.set(key, position);
  }
  const at = threadStampMs(c.timestamp);
  const threadId = c.parentId ?? c.id;
  if (at !== undefined && at > (index.latest.get(threadId) ?? Number.NEGATIVE_INFINITY)) {
    index.latest.set(threadId, at);
  }
}

/** `YYYY-MM-DDTHH:MM:SS[.f…][Z|±HH:MM]`, the timestamps the Rust writer converts to Excel's `dT` form. */
const THREAD_STAMP =
  /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?(Z|[+-]\d{2}:\d{2})?$/;

/**
 * Milliseconds since the epoch of a threaded comment timestamp, read as the
 * writer reads it (no zone is UTC, the zone Excel writes `dT` in), or
 * `undefined` when it is not such a timestamp.
 */
function threadStampMs(timestamp: string): number | undefined {
  const m = THREAD_STAMP.exec(timestamp);
  if (!m) return undefined;
  const [y, mo, d, h, mi, s] = m.slice(1, 7).map(Number) as [
    number,
    number,
    number,
    number,
    number,
    number,
  ];
  const day = new Date(Date.UTC(y, mo - 1, d));
  if (day.getUTCFullYear() !== y || day.getUTCMonth() !== mo - 1 || h > 23 || mi > 59 || s > 59) {
    return undefined;
  }
  const zone = m[8] ?? 'Z';
  const offsetMinutes =
    zone === 'Z'
      ? 0
      : (zone.startsWith('-') ? -1 : 1) * (Number(zone.slice(1, 3)) * 60 + Number(zone.slice(4)));
  const ms = Number((m[7] ?? '').padEnd(3, '0').slice(0, 3));
  return day.getTime() + ((h * 60 + mi - offsetMinutes) * 60 + s) * 1000 + ms;
}

/** The built-in print names, in the canonical spelling Excel writes. */
export type BuiltinPrintName = '_xlnm.Print_Titles' | '_xlnm.Print_Area';

/**
 * The canonical spelling of `name` when it is `_xlnm.Print_Area` or
 * `_xlnm.Print_Titles` in any letter case (Excel matches built-in names
 * without regard to case), else `undefined`.
 */
export function builtinPrintName(name: string): BuiltinPrintName | undefined {
  switch (name.toLowerCase()) {
    case '_xlnm.print_area':
      return '_xlnm.Print_Area';
    case '_xlnm.print_titles':
      return '_xlnm.Print_Titles';
    default:
      return undefined;
  }
}

/** Whether entry `d` is the built-in name `name`, in any letter case, scoped to `sheetIndex`. */
function isBuiltinEntry(d: DefinedNameData, name: BuiltinPrintName, sheetIndex: number): boolean {
  return d.sheetId === sheetIndex && builtinPrintName(d.name) === name;
}

/**
 * The value of the built-in defined name `name` scoped to `sheetIndex`, or
 * `null` when absent. An entry spelled in another letter case is the same
 * name, as Excel reads it; when several entries match, Excel reads the last,
 * and so does this.
 */
export function getBuiltinName(
  names: DefinedNameData[] | undefined,
  name: BuiltinPrintName,
  sheetIndex: number,
): string | null {
  return names?.findLast((d) => isBuiltinEntry(d, name, sheetIndex))?.value ?? null;
}

/**
 * Sets, replaces or (for `undefined`) removes the built-in defined name
 * `name` scoped to `sheetIndex`. An entry spelled in another letter case is
 * the same name, as Excel reads it, and Excel reads the last of several
 * matching entries. Setting updates that last entry in place (giving it the
 * canonical spelling), or appends one when none matches, and removes the
 * earlier matches, so exactly one entry remains; `undefined` removes every
 * matching entry.
 */
export function upsertBuiltinName(
  holder: { definedNames?: DefinedNameData[] | undefined },
  name: BuiltinPrintName,
  sheetIndex: number,
  value: string | undefined,
): void {
  holder.definedNames ??= [];
  const names = holder.definedNames;
  const matches = (d: DefinedNameData): boolean => isBuiltinEntry(d, name, sheetIndex);
  const last = value === undefined ? -1 : names.findLastIndex(matches);
  const kept = last === -1 ? undefined : names[last];
  for (let i = names.length - 1; i >= 0; i--) {
    const entry = names[i];
    if (entry && entry !== kept && matches(entry)) names.splice(i, 1);
  }
  if (value === undefined) return;
  if (kept) {
    kept.name = name;
    kept.value = value;
  } else {
    names.push({ name, value, sheetId: sheetIndex });
  }
}

/**
 * The single-sheet prefix at the start of a reference token's text: the
 * unquoted, `''`-unescaped sheet name and the prefix length up to (not
 * including) the `!`. `undefined` when the text has no prefix.
 */
export function readSheetPrefix(text: string): { sheet: string; length: number } | undefined {
  if (!text.startsWith("'")) {
    const bang = text.indexOf('!');
    return bang > 0 ? { sheet: text.slice(0, bang), length: bang } : undefined;
  }
  for (let i = 1; i < text.length; i++) {
    if (text.charAt(i) !== "'") continue;
    if (text.charAt(i + 1) === "'") {
      i++;
      continue;
    }
    return text.charAt(i + 1) === '!'
      ? { sheet: text.slice(1, i).replaceAll("''", "'"), length: i + 1 }
      : undefined;
  }
  return undefined;
}

/** Token types that make up one area of a print area or print titles value. */
const PRINT_AREA_PARTS: ReadonlySet<TokenType> = new Set<TokenType>([
  'cell_ref',
  'column_range',
  'row_range',
]);

/**
 * The offsets in `value` where an area with no sheet prefix starts, or
 * `undefined` when `tokens` are not a comma-separated list of areas. An area
 * is a cell reference, a `start:end` range of two cell references, a column
 * span (`A:B`) or a row span (`$1:$2`).
 */
function unqualifiedAreaStarts(tokens: readonly Token[], value: string): number[] | undefined {
  const starts: number[] = [];
  let i = 0;
  for (;;) {
    const first = tokens[i];
    if (!first || !PRINT_AREA_PARTS.has(first.type)) return undefined;
    if (!readSheetPrefix(value.slice(first.start, first.end))) starts.push(first.start);
    i += areaLength(tokens, i);
    if (i === tokens.length) return starts;
    if (tokens[i]?.type !== 'comma') return undefined;
    i++;
  }
}

/** The tokens of the area at `i`: 1, or 3 for `cell_ref colon cell_ref` (a `:` without a cell ref after it counts as 1 and fails at the caller). */
function areaLength(tokens: readonly Token[], i: number): number {
  return tokens[i]?.type === 'cell_ref' &&
    tokens[i + 1]?.type === 'colon' &&
    tokens[i + 2]?.type === 'cell_ref'
    ? 3
    : 1;
}

/**
 * Returns `value`, a built-in `_xlnm.Print_Area` or `_xlnm.Print_Titles` value,
 * with every comma-separated area that has no sheet prefix qualified with
 * `sheetName` (quoted when Excel needs it); the area is otherwise kept as written.
 * Excel refuses a file whose built-in name has an unqualified area.
 *
 * @throws ModernXlsxError with code `INVALID_ARGUMENT` if `value` is not a
 *   comma-separated list of references (cell, range, column span or row span).
 */
export function qualifyBuiltinNameValue(
  name: BuiltinPrintName,
  sheetName: string,
  value: string,
): string {
  const { tokens, errors } = tokenize(value);
  const starts =
    errors.length === 0 && !value.startsWith('=')
      ? unqualifiedAreaStarts(tokens, value)
      : undefined;
  if (!starts) {
    const label = name === '_xlnm.Print_Area' ? 'print area' : 'print titles';
    throw new ModernXlsxError(
      INVALID_ARGUMENT,
      `Invalid ${label} ${JSON.stringify(value)}: expected a reference or a comma-separated list of references such as "A1:D50"`,
    );
  }
  const prefix = `${quoteSheetName(sheetName)}!`;
  let result = value;
  for (const at of starts.toReversed()) result = result.slice(0, at) + prefix + result.slice(at);
  return result;
}

/** Longest table `displayName` or `name` Excel opens, in UTF-16 code units. */
export const MAX_TABLE_NAME_LENGTH = 255;

/** Characters Excel refuses anywhere in a table name: whitespace controls, space, DEL, and ASCII punctuation but `_ . \ ?`. */
export const TABLE_NAME_FORBIDDEN: RegExp = /[\t\n\r !"#$%&'()*+,\-/:;<=>@[\]^`{|}~\u007f]/;

/**
 * The first rule Excel applies to a table display name that `name` breaks,
 * as a short reason, or `undefined` when Excel opens it. The rules are the
 * ones an Excel 365 probe proved (see {@link Worksheet.addTable}); the last
 * three, shared with defined names, are {@link referenceNameProblem}. The
 * Rust `validate.rs` twin `name_problem` returns the same reasons, except
 * that it does not check a leading combining mark.
 */
function tableNameProblem(name: string): string | undefined {
  if (name === '') return 'it is empty';
  if (name.length > MAX_TABLE_NAME_LENGTH) return 'it is longer than 255 characters';
  if (TABLE_NAME_FORBIDDEN.test(name)) {
    return "it holds a character Excel refuses (a space, tab, line break, DEL, or ASCII punctuation other than '_', '.', '\\' and '?')";
  }
  if (/^[0-9.?]/.test(name)) return "it starts with a digit, '.' or '?'";
  if (/^\p{M}/u.test(name)) return 'it starts with a combining mark';
  return referenceNameProblem(name);
}

/** Index just past the quote closing the `"` or `'` at `pos` (a doubled quote stays inside), or -1. */
export function quotedEnd(formula: string, pos: number): number {
  const quote = formula.charAt(pos);
  for (let i = pos + 1; i < formula.length; i++) {
    if (formula.charAt(i) !== quote) continue;
    if (formula.charAt(i + 1) !== quote) return i + 1;
    i++;
  }
  return -1;
}

/** Index just past the `]` matching the `[` at `pos` (`'` escapes the next character), or -1. */
export function bracketEnd(formula: string, pos: number): number {
  let depth = 0;
  for (let i = pos; i < formula.length; i++) {
    const ch = formula.charAt(i);
    if (ch === "'") {
      i++;
    } else if (ch === '[') {
      depth++;
    } else if (ch === ']') {
      depth--;
      if (depth === 0) return i + 1;
    }
  }
  return -1;
}

/** A `@` this-row reference found by {@link thisRowShorthand}. */
interface ThisRowShorthand {
  /** The reference as written, table name included: `[@Qty]`, `Sales[@[Unit Cost]]`, `[@]`. */
  readonly ref: string;
  /** The table name before the `[`, or `''` when the reference names none. */
  readonly table: string;
  /** What follows the `@`: `Qty`, `[Unit Cost]`, or `''` for `[@]`. */
  readonly item: string;
}

/**
 * The first `@` this-row reference in `formula` (`[@Qty]`, `[@[Unit Cost]]`,
 * `Sales[@Qty]`, `[@]`) outside `"..."` string literals and `'...'` quoted
 * names, or `undefined` when there is none or a quote or bracket is never
 * closed. The table name is the run of characters a table name may hold
 * (none that {@link TABLE_NAME_FORBIDDEN} matches) just before the `[`.
 * A `[@...]` that names a column of that table (`columnsOf(table)`, which
 * gets `''` for no table name; see {@link namesColumn}) is a column
 * reference, not the shorthand: Excel writes a whole-column reference to a
 * column named `@Home` as `AtT[@Home]`. A table name right after `]!` names
 * a table of another workbook (Excel stores `Book.xlsx!Ext[@Home]` as
 * `[1]!Ext[@Home]`, whether that workbook is open or closed and whatever its
 * path), whose columns the workbook does not hold, so its `[@...]` is never
 * reported. One pass over `formula`, with one `columnsOf` call and one
 * lookup per `[@`. The Rust `validate.rs` twin is `this_row_shorthand`.
 */
function thisRowShorthand(
  formula: string,
  columnsOf: (table: string) => ReadonlySet<string>,
): ThisRowShorthand | undefined {
  let runStart = 0;
  let i = 0;
  while (i < formula.length) {
    const ch = formula.charAt(i);
    let end = i + 1;
    if (ch === '"' || ch === "'") {
      end = quotedEnd(formula, i);
    } else if (ch === '[') {
      end = bracketEnd(formula, i);
      const table = formula.slice(runStart, i);
      if (
        end > 0 &&
        formula.charAt(i + 1) === '@' &&
        !(runStart < i && runStart >= 2 && formula.slice(runStart - 2, runStart) === ']!') &&
        !namesColumn(formula.slice(i + 1, end - 1), columnsOf(table))
      ) {
        return {
          ref: formula.slice(runStart, end),
          table,
          item: formula.slice(i + 2, end - 1),
        };
      }
    } else if (!TABLE_NAME_FORBIDDEN.test(ch)) {
      i = end;
      continue;
    }
    if (end < 0) return undefined;
    i = end;
    runStart = end;
  }
  return undefined;
}

/**
 * `s` with each `_xHHHH_` escape (four hex digits, any case) turned into
 * U+HHHH in one left-to-right pass; an escape that names a surrogate stays as
 * written. A table part holds its column names and its formulas this way
 * (the text `_xHHHH_` itself as `_x005F_xHHHH_`). The Rust `validate.rs`
 * twin is `decode_xstring`.
 */
function decodeXstring(s: string): string {
  if (!s.includes('_x')) return s;
  return s.replace(/_x(?![dD][89a-fA-F])([0-9A-Fa-f]{4})_/g, (_, hex: string) =>
    String.fromCharCode(Number.parseInt(hex, 16)),
  );
}

/**
 * `true` when `inner`, the text of a structured-reference bracket without the
 * brackets, names a column whose name, decoded ({@link decodeXstring}) and
 * lower-cased, is in `columns`, ignoring case. `inner` is decoded as the
 * column names are, since the table part escapes both alike, and then loses
 * its `'` escapes; one lookup.
 */
function namesColumn(inner: string, columns: ReadonlySet<string>): boolean {
  return columns.has(decodeXstring(inner).replace(/'(.)?/gsu, '$1').toLowerCase());
}

/** No column names: what {@link checkTableFormulas} looks up for a table the workbook does not hold. */
const NO_COLUMNS: ReadonlySet<string> = new Set();

/**
 * Throws INVALID_ARGUMENT when a `calculatedColumnFormula` or
 * `totalsRowFormula` of `table` uses the `@` this-row shorthand, which Excel
 * 365 refuses to open in a table definition; the message gives the form
 * Excel writes there instead (`Sales[[#This Row],[Qty]]` for `[@Qty]`,
 * `Sales[#This Row]` for `[@]`). A `[@...]` naming a column of its table
 * (`table` when the reference names none, else the first table of the
 * workbook with that display name, ignoring case and canonical equivalence), or one of a
 * table in another workbook,
 * is a column reference and passes. `own` is the {@link nameKey} of
 * `table`'s display name, and `firstTable(key)` gives the first table of the
 * workbook, in sheet order, whose display name has the {@link nameKey} `key`.
 * The cost is linear in `table`'s columns and formulas plus the columns of the
 * tables they name:
 * a table's column names are decoded and lower-cased into a set once, the
 * first time a formula names it.
 */
function checkTableFormulas(
  table: TableDefinitionData,
  own: string,
  firstTable: (key: string) => TableDefinitionData | undefined,
): void {
  const tableNamed = (key: string): TableDefinitionData | undefined =>
    key === '' || key === own ? table : firstTable(key);
  const keys = new Map<TableDefinitionData, ReadonlySet<string>>();
  const columnsOf = (name: string): ReadonlySet<string> => {
    const found = tableNamed(nameKey(name));
    if (!found) return NO_COLUMNS;
    let set = keys.get(found);
    if (!set) {
      set = new Set(found.columns.map((column) => decodeXstring(column.name).toLowerCase()));
      keys.set(found, set);
    }
    return set;
  };
  for (const column of table.columns) {
    for (const field of ['calculatedColumnFormula', 'totalsRowFormula'] as const) {
      const found = column[field] ? thisRowShorthand(column[field], columnsOf) : undefined;
      if (!found) continue;
      throw new ModernXlsxError(
        INVALID_ARGUMENT,
        `Invalid ${field} of column ${JSON.stringify(column.name)} in table ${JSON.stringify(table.displayName)}: Excel refuses the @ shorthand ${found.ref} in a table definition; write ${thisRowFileForm(found.table || table.displayName, found.item)} instead`,
      );
    }
  }
}

/** The form Excel writes in a file for `[@item]` of table `name`: `name[[#This Row],[item]]`, or `name[#This Row]` for `[@]`. */
function thisRowFileForm(name: string, item: string): string {
  if (item === '') return `${name}[#This Row]`;
  return `${name}[[#This Row],${item.startsWith('[') ? item : `[${item}]`}]`;
}

/** `decodeRange(ref)`, or `undefined` when `ref` is not a range. */
function tryDecodeRange(ref: string): CellRange | undefined {
  try {
    return decodeRange(ref);
  } catch {
    return undefined;
  }
}

/** The `SUBTOTAL` function number Excel writes for each totals-row function (ECMA-376 `totalsRowFunction`). */
const SUBTOTAL_CODES: ReadonlyMap<string, number> = new Map([
  ['average', 101],
  ['countNums', 102],
  ['count', 103],
  ['max', 104],
  ['min', 105],
  ['stdDev', 107],
  ['sum', 109],
  ['var', 110],
]);

/**
 * A column name as a structured-reference item: `[`, `]`, `#` and `'` get a leading `'`, the
 * characters Excel requires escaped on input. Excel also shows `@` escaped (`[x'@y]`) but
 * accepts it unescaped, as written here.
 */
function escapeStructuredColumn(name: string): string {
  return name.replaceAll(/[[\]#']/g, "'$&");
}

/** What Excel stores in the totals-row cell of `column`, or `undefined` for an empty cell. */
function totalsCellContent(
  displayName: string,
  column: TableColumnData,
): { formula: string } | { label: string } | undefined {
  const fn = column.totalsRowFunction;
  const code = fn ? SUBTOTAL_CODES.get(fn) : undefined;
  if (code !== undefined) {
    return { formula: `SUBTOTAL(${code},${displayName}[${escapeStructuredColumn(column.name)}])` };
  }
  if (fn === 'custom') {
    const formula = column.totalsRowFormula?.replace(/^=/, '');
    return formula ? { formula } : undefined;
  }
  return column.totalsRowLabel ? { label: column.totalsRowLabel } : undefined;
}
