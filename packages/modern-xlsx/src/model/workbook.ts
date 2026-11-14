import { writeFileBytes } from '#runtime';
import { StyleBuilder } from '../authoring/style-builder.js';
import { assertImageAnchor } from '../barcode/common.js';
import type { DrawBarcodeOptions, ImageAnchor } from '../barcode/index.js';
import { generateBarcode, generateDrawingRels, generateDrawingXml } from '../barcode/index.js';
import { quoteSheetName } from '../formula/serializer.js';
import type { TokenType } from '../formula/tokenizer.js';
import { tokenize } from '../formula/tokenizer.js';
import type { WasmLoader } from '../wasm-backend.js';
import { INVALID_ARGUMENT, ModernXlsxError, SHEET_NOT_FOUND, WASM_INIT_FAILED } from './errors.js';
import { excelGuid } from './guid.js';
import { NameKeyIndex, nameKey, referenceNameProblem, workbookNameIndexes } from './name-rules.js';
import type {
  CalcChainEntryData,
  DateSystem,
  DefinedNameData,
  DocPropertiesData,
  PivotCacheDefinitionData,
  PivotCacheRecordsData,
  RepairResult,
  SheetData,
  SlicerCacheData,
  StylesData,
  TableDefinitionData,
  ThemeColorsData,
  ThreadedCommentData,
  TimelineCacheData,
  ValidationReport,
  WorkbookData,
  WorkbookProtectionData,
  WorkbookViewData,
  WriteOptions,
} from './types.js';
import {
  type BuiltinPrintName,
  bracketEnd,
  builtinPrintName,
  getBuiltinName,
  MAX_TABLE_NAME_LENGTH,
  qualifyBuiltinNameValue,
  quotedEnd,
  readSheetPrefix,
  TABLE_NAME_FORBIDDEN,
  upsertBuiltinName,
  validateSheetName,
  Worksheet,
} from './worksheet.js';

const TEXT_ENCODER = new TextEncoder();
const TEXT_DECODER = new TextDecoder();

/** Media parts `addImage` must not collide with: `xl/media/image<n>.<ext>`. */
const MEDIA_IMAGE_PATH = /^xl[/]media[/]image([0-9]+)[.][^/]+$/;

/** Optional `WorkbookData` fields a seeded workbook copies over when present. */
const OPTIONAL_WORKBOOK_KEYS = [
  'definedNames',
  'sharedStrings',
  'docProperties',
  'themeColors',
  'calcChain',
  'workbookViews',
  'protection',
  'calcPr',
  'fileVersion',
  'workbookPr',
  'externalReferences',
  'persons',
  'pivotCaches',
  'pivotCacheRecords',
  'slicerCaches',
  'timelineCaches',
  'preservedEntries',
  'preservedContentTypes',
] as const satisfies readonly (keyof WorkbookData)[];

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
 * This class is the unbound base: the `modern-xlsx` and `modern-xlsx/lite`
 * entries each export a subclass bound to their WASM build, and the WASM-backed
 * methods (`toBuffer`, `toFile`, `validate`, `repair`) of the base itself throw
 * `WASM_INIT_FAILED`.
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
  /** The WASM build an entry subclass binds; unset on the base. */
  protected static loader?: WasmLoader;

  private readonly data: WorkbookData;
  private readonly imageAnchors = new Map<string, { anchor: ImageAnchor; imageIndex: number }[]>();
  private readonly imageRels = new Map<string, { rId: string; target: string }[]>();
  private _sheetIndex = new Map<string, number>();
  /** The {@link nameKey} of each sheet name, kept beside `_sheetIndex` for the sheet-name check. */
  readonly #sheetKeys = new NameKeyIndex<SheetData>('name');

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
    if (data) {
      for (const key of OPTIONAL_WORKBOOK_KEYS) {
        if (data[key]) Object.assign(this.data, { [key]: data[key] });
      }
    }
    this._rebuildSheetIndex();
  }

  /**
   * The loader of the WASM build this instance's class is bound to.
   *
   * @throws ModernXlsxError with code `WASM_INIT_FAILED` on the unbound base class.
   */
  #wasm(): WasmLoader {
    const { loader } = this.constructor as typeof Workbook;
    if (!loader) {
      throw new ModernXlsxError(
        WASM_INIT_FAILED,
        "Workbook is not bound to a WASM build; import it from 'modern-xlsx' or 'modern-xlsx/lite'.",
      );
    }
    return loader;
  }

  private _rebuildSheetIndex(): void {
    this._sheetIndex = new Map<string, number>();
    for (let i = 0; i < this.data.sheets.length; i++) {
      const sheet = this.data.sheets[i];
      if (sheet) this._sheetIndex.set(sheet.name, i);
    }
    this.#sheetKeys.sync([this.data.sheets]);
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
   * The lookup is case-sensitive.
   *
   * @param name - The sheet tab name to look up, matched exactly.
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
    return sheet ? new Worksheet(sheet, this.data.styles, this.data, this) : undefined;
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
    return sheet ? new Worksheet(sheet, this.data.styles, this.data, this) : undefined;
  }

  /**
   * `true` when a sheet other than `exceptIndex` already has `name`. Sheet
   * names are compared by {@link nameKey}, ignoring case and canonical
   * equivalence: Excel refuses to open a file with two sheets `café` (U+00E9)
   * and `cafe` + U+0301. The stored names' keys come from an index
   * ({@link NameKeyIndex}), so the check computes the key of `name` alone.
   */
  #sheetNameTaken(name: string, exceptIndex?: number): boolean {
    const key = nameKey(name);
    const keys = this.#sheetKeys.sync([this.data.sheets]);
    keys.hint(name, key);
    return keys.first(key, (_, i) => i !== exceptIndex) !== undefined;
  }

  /**
   * Remap every stored sheet position after the sheet list was spliced.
   * `map` takes an old 0-based index to its new one, or to -1 when that sheet
   * was removed. Sheet-scoped defined names follow their sheet and are dropped
   * with it; workbook views follow their sheet and, when it was removed, are
   * clamped into the remaining range.
   */
  #remapSheetRefs(map: (index: number) => number): void {
    const names = this.data.definedNames;
    if (names) {
      for (let i = names.length - 1; i >= 0; i--) {
        const entry = names[i];
        if (entry?.sheetId == null) continue;
        const next = map(entry.sheetId);
        if (next === -1) names.splice(i, 1);
        else entry.sheetId = next;
      }
    }
    const last = this.data.sheets.length - 1;
    const remapView = (index: number): number => {
      const next = map(index);
      return next === -1 ? Math.max(0, Math.min(index, last)) : next;
    };
    for (const view of this.data.workbookViews ?? []) {
      view.activeTab = remapView(view.activeTab);
      view.firstSheet = remapView(view.firstSheet);
    }
  }

  /**
   * Validate a defined name before it is stored.
   *
   * @param name - The defined name.
   * @param sheetId - The 0-based scope sheet, or `null`/`undefined` for workbook scope.
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` when the name is not
   *   1-255 characters, does not start with a letter, `_` or `\`, has a later
   *   character other than a letter, combining mark, digit, `_`, `.`, `\` or
   *   `?`, is `TRUE` or `FALSE`, is an A1 cell reference inside the grid, is
   *   an R1C1 reference or starts with one, uses the `_xlnm.` prefix for
   *   anything but `Print_Area`, `Print_Titles` or `_FilterDatabase`, is
   *   scoped to a sheet index outside the workbook, or repeats a name in the
   *   same scope or a table's display name on any sheet, duplicates being
   *   found ignoring case and canonical equivalence (`café` with U+00E9
   *   repeats `cafe` + U+0301).
   */
  #validateDefinedName(name: string, sheetId: number | null | undefined): void {
    checkDefinedNameShape(name);
    const key = nameKey(name);
    checkReservedDefinedName(name, key);
    this.#checkDefinedNameScope(name, sheetId);
    const scope = sheetId ?? null;
    const indexes = workbookNameIndexes(this.data);
    const defined = indexes.definedNames.sync([this.data.definedNames]);
    defined.hint(name, key);
    if (defined.first(key, (d) => (d.sheetId ?? null) === scope)) {
      throw new ModernXlsxError(
        INVALID_ARGUMENT,
        `Defined name "${name}" already exists in this scope`,
      );
    }
    // Excel refuses a defined name equal to a table name at every scope (the
    // mirror of the check `Worksheet.addTable` makes).
    const table = indexes.tables
      .sync(this.data.sheets.map((sheet) => sheet.worksheet.tables))
      .first(key);
    if (table) {
      throw new ModernXlsxError(
        INVALID_ARGUMENT,
        `Defined name ${JSON.stringify(name)} is already used by table ${JSON.stringify(table.displayName)} (names ignore case and canonical equivalence)`,
      );
    }
  }

  /** Throws unless `sheetId` is `null`/`undefined` (workbook scope) or an existing sheet index. */
  #checkDefinedNameScope(name: string, sheetId: number | null | undefined): void {
    if (
      sheetId != null &&
      !(Number.isInteger(sheetId) && sheetId >= 0 && sheetId < this.data.sheets.length)
    ) {
      throw new ModernXlsxError(
        INVALID_ARGUMENT,
        `Invalid sheet index for defined name "${name}": ${sheetId}`,
      );
    }
  }

  /**
   * Give a sheet clone the ids that must be unique across the workbook: each
   * table gets the next table id and a `${displayName}_${n}` name, the display
   * name cut to keep it within 255 UTF-16 units and given a leading `_` when
   * the cut would make it an R1C1 reference (which the
   * clone's structured references follow), and each
   * threaded comment gets a fresh GUID in Excel's braced upper-case form,
   * with reply `parentId`s and legacy comment authors (`tc=${id}`) following.
   * A thread whose placeholder note exists only in the written file (one the
   * API added) gets its placeholder from the writer under the new id.
   * Call it before the clone joins the sheet list.
   */
  #reidClonedSheet(clone: SheetData): void {
    this.#reidClonedTables(clone);
    this.#reidClonedThreadedComments(clone);
  }

  /**
   * Re-ids the clone's tables against every table id and name, and defined
   * name, in the workbook, then points the clone's structured references at
   * the new names: in every cell formula of the clone and in its tables'
   * `calculatedColumnFormula` and `totalsRowFormula`.
   */
  #reidClonedTables(clone: SheetData): void {
    const tables = clone.worksheet.tables;
    if (!tables?.length) return;
    let maxId = 0;
    const taken = new Set<string>();
    for (const sheet of this.data.sheets) {
      for (const table of sheet.worksheet.tables ?? []) {
        maxId = Math.max(maxId, table.id);
        taken.add(nameKey(table.displayName));
        if (table.name) taken.add(nameKey(table.name));
      }
    }
    for (const entry of this.data.definedNames ?? []) taken.add(nameKey(entry.name));
    renameClonedTableRefs(clone, reidTables(tables, maxId, taken));
  }

  /** Gives the clone's threaded comments fresh GUIDs and updates the legacy `tc=${id}` comment authors. */
  #reidClonedThreadedComments(clone: SheetData): void {
    const threaded = clone.worksheet.threadedComments;
    if (!threaded?.length) return;
    const newIds = reidThreadedComments(threaded);
    // A legacy comment that mirrors a thread (the placeholder a read file
    // holds) names it as author `tc=${id}`.
    for (const comment of clone.worksheet.comments ?? []) {
      if (!comment.author.startsWith('tc=')) continue;
      const id = newIds.get(comment.author.slice(3));
      if (id !== undefined) comment.author = `tc=${id}`;
    }
  }

  /**
   * Adds a new empty sheet to the workbook and returns it.
   * The name is validated against Excel's sheet-name rules (see {@link Worksheet.name}).
   *
   * @param name - The sheet tab name. Must be 1-31 UTF-16 code units, free of
   *   `\ / * ? [ ] :` and control characters, not start or end with an apostrophe, and be unique
   *   ignoring case and canonical equivalence (`café` with U+00E9 and `cafe` + U+0301 are one
   *   name); `History` is reserved in any case.
   * @returns The newly created Worksheet.
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if the name is invalid
   *   or reserved, or a sheet with that name (in any case or normal form) already exists.
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
    if (this.#sheetNameTaken(name)) {
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
    return new Worksheet(sheetData, this.data.styles, this.data, this);
  }

  /**
   * Removes a sheet by name or 0-based index.
   *
   * Defined names scoped to the removed sheet are dropped, later sheets'
   * scoped names and the workbook views' `activeTab`/`firstSheet` follow
   * their sheets, and the calculation chain is dropped (Excel rebuilds it).
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
    const idx = this.#resolveSheet(nameOrIndex);
    if (!Number.isInteger(idx) || idx < 0 || idx >= this.data.sheets.length) return false;
    this.data.sheets.splice(idx, 1);
    this.#remapSheetRefs((i) => (i < idx ? i : i === idx ? -1 : i - 1));
    delete this.data.calcChain;
    this._rebuildSheetIndex();
    return true;
  }

  /**
   * Moves a sheet from one position to another.
   *
   * Sheet-scoped defined names and the workbook views' `activeTab`/`firstSheet`
   * follow their sheets. A real move (`fromIndex !== toIndex`) drops the
   * calculation chain, which Excel rebuilds.
   *
   * @param fromIndex - Zero-based source position.
   * @param toIndex - Zero-based target position.
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if either index is not
   *   an integer in range.
   *
   * @example
   * ```ts
   * wb.moveSheet(2, 0); // move third sheet to first position
   * ```
   */
  moveSheet(fromIndex: number, toIndex: number): void {
    const count = this.data.sheets.length;
    if (!Number.isInteger(fromIndex) || fromIndex < 0 || fromIndex >= count) {
      throw new ModernXlsxError(INVALID_ARGUMENT, `Invalid source index: ${fromIndex}`);
    }
    if (!Number.isInteger(toIndex) || toIndex < 0 || toIndex >= count) {
      throw new ModernXlsxError(INVALID_ARGUMENT, `Invalid target index: ${toIndex}`);
    }
    if (fromIndex === toIndex) return;
    const [sheet] = this.data.sheets.splice(fromIndex, 1);
    if (!sheet) return;
    this.data.sheets.splice(toIndex, 0, sheet);
    this.#remapSheetRefs((i) => {
      if (i === fromIndex) return toIndex;
      if (fromIndex < toIndex && i > fromIndex && i <= toIndex) return i - 1;
      if (toIndex <= i && i < fromIndex) return i + 1;
      return i;
    });
    delete this.data.calcChain;
    this._rebuildSheetIndex();
  }

  /**
   * Clones a sheet and inserts the copy at the end (or given position).
   *
   * The copy's tables get new workbook-unique ids and names (`Sales` becomes
   * `Sales_2`, or the next free suffix; a source name too long for the
   * suffix is cut so the new name stays within Excel's 255 UTF-16 units,
   * never inside a surrogate pair, and a cut name that would read as a cell
   * or R1C1 reference gets a leading `_`), and its threaded comments get new
   * ids (upper-case GUIDs in braces, the form Excel requires) with replies
   * and legacy `tc=` comment authors still pointing at them. Inserting before an existing sheet shifts later sheet-scoped names
   * and workbook views, and drops the calculation chain.
   *
   * Structured references to a renamed table (`Sales[Amount]`, matched
   * ignoring case and canonical equivalence, and names such as `a\b`, `a?b` or `Ventes€` the formula
   * tokenizer cannot scan) are rewritten to the new name in the copy only:
   * in every cell formula of the copy (its totals row included) and in the
   * `calculatedColumnFormula` and `totalsRowFormula` of its tables. Text in
   * string literals, references to tables on other sheets, formulas on other
   * sheets or in defined names that name the source table, and a whole
   * formula the tokenizer rejects for a reason other than a table name (a
   * sheet-qualified `Data!Sales[Amount]`, an unclosed string, or another name
   * it cannot scan, such as a defined name `\rate`) keep naming the source
   * table.
   *
   * @param sourceIndex - Zero-based index of the sheet to clone.
   * @param newName - Tab name for the cloned sheet (unique ignoring case and
   *   canonical equivalence).
   * @param insertIndex - Optional zero-based position for the clone, from 0 to
   *   the sheet count; defaults to the end.
   * @returns The newly created Worksheet (deep copy of the source).
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if either index is
   *   invalid, or the new name is invalid, reserved or already taken.
   *
   * @example
   * ```ts
   * const copy = wb.cloneSheet(0, 'Sheet1 Copy');
   * ```
   */
  cloneSheet(sourceIndex: number, newName: string, insertIndex?: number): Worksheet {
    const count = this.data.sheets.length;
    const source = Number.isInteger(sourceIndex) ? this.data.sheets[sourceIndex] : undefined;
    if (!source) {
      throw new ModernXlsxError(INVALID_ARGUMENT, `Invalid source index: ${sourceIndex}`);
    }
    if (
      insertIndex !== undefined &&
      !(Number.isInteger(insertIndex) && insertIndex >= 0 && insertIndex <= count)
    ) {
      throw new ModernXlsxError(INVALID_ARGUMENT, `Invalid insert index: ${insertIndex}`);
    }
    validateSheetName(newName);
    if (this.#sheetNameTaken(newName)) {
      throw new ModernXlsxError(INVALID_ARGUMENT, `Sheet "${newName}" already exists`);
    }
    const clone: SheetData = structuredClone(source);
    clone.name = newName;
    this.#reidClonedSheet(clone);
    const idx = insertIndex ?? count;
    this.data.sheets.splice(idx, 0, clone);
    if (idx < count) {
      this.#remapSheetRefs((i) => (i >= idx ? i + 1 : i));
      delete this.data.calcChain;
    }
    this._rebuildSheetIndex();
    return new Worksheet(clone, this.data.styles, this.data, this);
  }

  /**
   * Renames a sheet and rewrites the defined names that refer to it.
   *
   * Every defined-name value (print areas and print titles included) has its
   * sheet prefixes that name the old sheet, quoted or not and compared
   * ignoring case (final sigma folded; U+0130 is not matched by `i` +
   * U+0307), not canonical equivalence, replaced
   * with the new name, quoted when needed, as Excel's own rename stores them:
   * `Data!$A$1` becomes `'My Data'!$A$1`, and an end of a 3-D span is
   * rewritten too (`Data:Data3!A1` becomes `'My Data:Data3'!A1`; renaming a
   * sheet between the ends changes nothing). A prefix with a colon is a span
   * only when both ends name sheets of the workbook (compared the same way,
   * before the rename). Otherwise only a quoted one is a path into another
   * workbook, which Excel's own rename keeps, and so does this one:
   * `'C:/dir/Book.xlsx'!A1` and `'C:Nosuch'!A1` are kept when sheet `C` is
   * renamed and no sheet `Nosuch` exists. An unquoted one whose first end
   * is no sheet is a name, the range operator and a sheet-qualified
   * reference, and its last end follows the rename as in Excel:
   * `First:Data!$C$10` becomes `First:'My Data'!$C$10`, while
   * `SUM(Data:Nosuch!$A$1)` is kept. A sheet name right after the range
   * operator is written quoted, as Excel 365 stores it (2026-11-14):
   * renaming Data to First turns `First:Data!$C$10` and
   * `First:'Data'!$C$10` into `First:'First'!$C$10` (unquoted, the text
   * would be the one-sheet span First:First), and renaming Sheet3 to First
   * turns `First:Sheet1!$C$10` into `First:'Sheet1'!$C$10`. A name and a
   * sheet spelled alike are told apart by those quotes alone, in Excel too:
   * hand-written `First:First!$C$10` is a span. A prefix spelled in another
   * normal form (`cafe` + U+0301 for a sheet `café` with U+00E9) is kept,
   * as Excel 365's own rename keeps it (2026-11-13): Excel does not take it
   * for the sheet.
   * A reference into another workbook (`[1]Data!$A$1`, the form Excel stores)
   * keeps naming that workbook's sheet, as do text in string literals and a
   * longer sheet name (`Data10`). A value the formula tokenizer rejects is
   * left unchanged as a whole, among them one with a sheet-qualified defined
   * name such as `Data!Rate`, which Excel rewrites. Cell formulas, chart
   * series, pivot cache sources, data validations, conditional formats and
   * hyperlinks are not rewritten. Only values change: the names themselves
   * are kept as they are and not checked, so a name read from a file that
   * `addNamedRange` would refuse (`a€b`) does not stop the rename.
   *
   * @param nameOrIndex - The current sheet name or zero-based index.
   * @param newName - The new tab name (validated against Excel's sheet-name
   *   rules, unique ignoring case and canonical equivalence; a rename of the
   *   same sheet to another case or normal form of its name is allowed).
   * @throws ModernXlsxError with code `SHEET_NOT_FOUND` if the sheet is not
   *   found, or `INVALID_ARGUMENT` if the new name is invalid, reserved or
   *   taken by another sheet (nothing is changed then).
   *
   * @example
   * ```ts
   * wb.renameSheet('Sheet1', 'January');
   * ```
   */
  renameSheet(nameOrIndex: string | number, newName: string): void {
    const idx = this.#resolveSheet(nameOrIndex);
    const sheet = this.data.sheets[idx];
    if (!sheet) {
      throw new ModernXlsxError(SHEET_NOT_FOUND, `Sheet not found: ${nameOrIndex}`);
    }
    validateSheetName(newName);
    if (this.#sheetNameTaken(newName, idx)) {
      throw new ModernXlsxError(INVALID_ARGUMENT, `Sheet "${newName}" already exists`);
    }
    let sheetCases: ReadonlySet<string> | undefined;
    const isSheet = (end: string): boolean => {
      // A name that is not a string (data built by hand) names no sheet here.
      sheetCases ??= new Set(
        this.data.sheets.flatMap((s) => (typeof s?.name === 'string' ? [caseKey(s.name)] : [])),
      );
      return sheetCases.has(caseKey(end));
    };
    // A sheet whose own stored name is not a string (data built by hand) is
    // named by no prefix, so the defined names are kept as they are.
    if (typeof sheet.name === 'string') {
      const rename: SheetRename = {
        target: caseKey(sheet.name),
        newName,
        newKey: caseKey(newName),
        isSheet,
      };
      for (const entry of this.data.definedNames ?? []) {
        const value = renameSheetPrefixes(entry.value, rename);
        if (value !== undefined) entry.value = value;
      }
    }
    sheet.name = newName;
    this._rebuildSheetIndex();
  }

  /**
   * Hides a sheet. At least one sheet must remain visible.
   *
   * @param nameOrIndex - The sheet name or zero-based index to hide.
   * @throws ModernXlsxError with code `SHEET_NOT_FOUND` if the sheet is not
   *   found, or `INVALID_ARGUMENT` if it is the last visible sheet.
   *
   * @example
   * ```ts
   * wb.hideSheet('Internal');
   * ```
   */
  hideSheet(nameOrIndex: string | number): void {
    const idx = this.#resolveSheet(nameOrIndex);
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
   * @throws ModernXlsxError with code `SHEET_NOT_FOUND` if the sheet is not found.
   *
   * @example
   * ```ts
   * wb.unhideSheet('Internal');
   * ```
   */
  unhideSheet(nameOrIndex: string | number): void {
    const idx = this.#resolveSheet(nameOrIndex);
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
   * @param name - The defined name (e.g., `'TotalSales'`): 1-255 characters,
   *   starting with a letter, `_` or `\`, then letters, combining marks (the
   *   vowel signs of Thai and Indic scripts, Arabic harakat, Hebrew points,
   *   decomposed accents), digits, `_`, `.`, `\` or `?`; not `TRUE` or
   *   `FALSE`; not a cell reference inside the grid
   *   (`A1`, `T2`; `XFE1` and `A0` are fine); not an R1C1 reference or a name
   *   that starts with one (`R`, `C`, `R1C1`, `R1x`, `C1_total`, `R2D2`; a
   *   `.` right after the number, as in `R1.a`, is fine); and, ignoring case
   *   and canonical equivalence (`café` with U+00E9 is `cafe` + U+0301), not
   *   the display name of a table on any sheet. `Worksheet.addTable`
   *   applies the same reference rules to a table name, whose characters may
   *   also be symbols such as `€` that a defined name may not hold. The
   *   `_xlnm.` prefix is accepted only for `Print_Area`, `Print_Titles` and
   *   `_FilterDatabase`. Excel matches built-in names without regard to
   *   letter case, so `_xlnm.Print_Area` and `_xlnm.Print_Titles` in any case
   *   (`'_xlnm.print_area'`) are stored in the spelling Excel writes,
   *   `_xlnm.Print_Area` and `_xlnm.Print_Titles`, and the print-area and
   *   print-titles getters and setters see them.
   * @param value - The range reference (e.g., `'Sheet1!$A$1:$A$10'`). For
   *   `_xlnm.Print_Area` and `_xlnm.Print_Titles` (in any letter case) with a
   *   `sheetId`, an area without a sheet prefix is stored qualified with that
   *   sheet's name.
   * @param sheetId - Optional zero-based sheet index to scope the name to.
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if the name breaks
   *   those rules, `sheetId` is not a sheet index of this workbook, the
   *   name (ignoring case and canonical equivalence) already exists in the
   *   same scope or names a
   *   table, or the `value`
   *   of a built-in print name with a `sheetId` is not a comma-separated list
   *   of references.
   *
   * @example
   * ```ts
   * wb.addNamedRange('Revenue', 'Sheet1!$B$2:$B$100');
   * ```
   */
  addNamedRange(name: string, value: string, sheetId?: number): void {
    this.#validateDefinedName(name, sheetId);
    const printName = builtinPrintName(name);
    if (printName !== undefined) {
      name = printName;
      if (typeof sheetId === 'number') value = this.#qualifyBuiltin(printName, sheetId, value);
    }
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

  /** Resolves a sheet name (exact, case-sensitive) or index to a zero-based index; an unknown name gives -1. */
  #resolveSheet(nameOrIndex: string | number): number {
    return typeof nameOrIndex === 'string'
      ? this.data.sheets.findIndex((s) => s.name === nameOrIndex)
      : nameOrIndex;
  }

  // --- Print titles & print area ---

  /**
   * Gets the print titles (repeat rows/columns) for the given sheet.
   * @param sheetIndex 0-based sheet index.
   * @returns The defined name value (e.g., `"Sheet1!$1:$2,Sheet1!$A:$B"`) or `null`.
   */
  getPrintTitles(sheetIndex: number): string | null {
    return getBuiltinName(this.data.definedNames, '_xlnm.Print_Titles', sheetIndex);
  }

  /**
   * Sets print titles (repeat rows/columns) for the given sheet.
   *
   * Excel refuses a file whose print titles have no sheet qualifier, so every
   * comma-separated area of `value` that lacks one (`"$1:$2"`) is stored qualified
   * with the sheet's name, quoted when needed, and otherwise as written; a
   * qualified area (`"Sheet1!$1:$2"`) is stored as it is.
   * @param sheetIndex 0-based sheet index.
   * @param value The row span, column span or comma-separated list of them (e.g.
   *   `"$1:$2"` for the first 2 rows, `"$A:$B,$1:$2"`). Pass `null` to clear.
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if `sheetIndex` is not
   *   the index of an existing sheet, or `value` is not a reference or a
   *   comma-separated list of references.
   */
  setPrintTitles(sheetIndex: number, value: string | null): void {
    this.#checkDefinedNameScope('_xlnm.Print_Titles', sheetIndex);
    upsertBuiltinName(
      this.data,
      '_xlnm.Print_Titles',
      sheetIndex,
      value === null ? undefined : this.#qualifyBuiltin('_xlnm.Print_Titles', sheetIndex, value),
    );
  }

  /**
   * Gets the print area for the given sheet.
   * @param sheetIndex 0-based sheet index.
   * @returns The defined name value (e.g., `"Sheet1!$A$1:$D$50"`) or `null`.
   */
  getPrintArea(sheetIndex: number): string | null {
    return getBuiltinName(this.data.definedNames, '_xlnm.Print_Area', sheetIndex);
  }

  /**
   * Sets the print area for the given sheet.
   *
   * Excel refuses a file whose print area has no sheet qualifier, so every
   * comma-separated area of `value` that lacks one (`"A1:B2"`) is stored qualified
   * with the sheet's name, quoted when needed (`"'My Sheet'!A1:B2"`), and otherwise
   * as written; a qualified area (`"Sheet1!$A$1:$D$50"`) is stored as it is.
   * @param sheetIndex 0-based sheet index.
   * @param value The reference or comma-separated list of references (cell, range,
   *   column span or row span). Pass `null` to clear.
   * @throws ModernXlsxError with code `INVALID_ARGUMENT` if `sheetIndex` is not
   *   the index of an existing sheet, or `value` is not a reference or a
   *   comma-separated list of references.
   */
  setPrintArea(sheetIndex: number, value: string | null): void {
    this.#checkDefinedNameScope('_xlnm.Print_Area', sheetIndex);
    upsertBuiltinName(
      this.data,
      '_xlnm.Print_Area',
      sheetIndex,
      value === null ? undefined : this.#qualifyBuiltin('_xlnm.Print_Area', sheetIndex, value),
    );
  }

  /** `value` of a built-in name for the sheet at `sheetIndex`, with unqualified areas qualified with the sheet's name. */
  #qualifyBuiltin(name: BuiltinPrintName, sheetIndex: number, value: string): string {
    return qualifyBuiltinNameValue(name, this.data.sheets[sheetIndex]?.name ?? '', value);
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
   * When a pivot table that uses this cache has a row, column or page field with empty
   * `items`, the writer completes that table from the cache and writes the cache with
   * `refreshOnLoad="1"`, so Excel recalculates the pivot when it opens the file (see
   * `PivotTableData`).
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
   * @returns The XLSX file as a Uint8Array backed by its own ArrayBuffer, so it can be
   *   passed straight to `new Response()` or `new Blob()`.
   * @throws ModernXlsxError with code `WASM_INIT_FAILED` if WASM is not initialized,
   *   `INVALID_ARGUMENT` for a password in the `modern-xlsx/lite` build,
   *   `INVALID_WORKBOOK` for the causes listed with that code in `errors.ts`,
   *   `INVALID_CELL_REF` for a frozen pane that puts its top-left cell past XFD1048576,
   *   `INVALID_CELL_VALUE` for a `sharedString` cell without a value or a `cellMetadata`
   *   other than 1 in a workbook without a metadata part, and
   *   `SIZE_LIMIT` when completing pivot tables built in code would pass the pivot item
   *   ceilings, the notes the write generates (VML note shapes and threaded-comment
   *   placeholders) would pass their ceiling, or an encrypted package stream would pass
   *   2 GiB (the numbers are in the `SIZE_LIMIT` entry of `errors.ts`).
   *
   * @example
   * ```ts
   * const bytes = await wb.toBuffer();
   * await fs.writeFile('output.xlsx', bytes);
   * ```
   */
  async toBuffer(options?: WriteOptions): Promise<Uint8Array<ArrayBuffer>> {
    const wasm = this.#wasm();
    wasm.ensureInitialized();
    if (options?.password) {
      return wasm.wasmWriteWithPassword(this.data, options.password);
    }
    return wasm.wasmWrite(this.data);
  }

  /**
   * Write the workbook to a file on disk.
   * Only available in Node.js, Bun, and Deno environments.
   *
   * @param path - File path to write to.
   * @param options - Optional write options. Pass `{ password: '...' }` to encrypt.
   * @throws ModernXlsxError with code `WASM_INIT_FAILED` if WASM is not initialized,
   *   `INVALID_ARGUMENT` for a password in the `modern-xlsx/lite` build,
   *   `INVALID_WORKBOOK` for the causes listed with that code in `errors.ts`,
   *   `INVALID_CELL_REF` for a frozen pane that puts its top-left cell past XFD1048576,
   *   `INVALID_CELL_VALUE` for a `sharedString` cell without a value or a `cellMetadata`
   *   other than 1 in a workbook without a metadata part,
   *   `SIZE_LIMIT` when completing pivot tables built in code would pass the pivot item
   *   ceilings, the notes the write generates (VML note shapes and threaded-comment
   *   placeholders) would pass their ceiling, or an encrypted package stream would pass
   *   2 GiB (the numbers are in the `SIZE_LIMIT` entry of `errors.ts`), and
   *   `IO_ERROR` outside Node.js-compatible runtimes (browsers, workers, the IIFE build).
   *
   * @example
   * ```ts
   * await wb.toFile('./report.xlsx');
   * await wb.toFile('./secret.xlsx', { password: 's3cret' });
   * ```
   */
  async toFile(path: string, options?: WriteOptions): Promise<void> {
    await writeFileBytes(path, await this.toBuffer(options));
  }

  /**
   * Validate the workbook for OOXML compliance issues.
   * Returns a structured report with errors, warnings, and fix suggestions.
   * The report stores the issues found first, at most 10,000 of them and at
   * most 16 MiB of message and location text; its counts cover every issue,
   * and `omittedIssueCount` says how many were left out.
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
    const wasm = this.#wasm();
    wasm.ensureInitialized();
    return wasm.wasmValidate(this.data);
  }

  /**
   * Validate and auto-repair the workbook.
   * Fixes the repairable issues the validator reports: invalid and duplicate
   * sheet names, missing default styles and dangling style ids, shared-string
   * cells without a value, a missing theme, malformed document-property dates
   * and whitespace-only document-property text fields, row order, and reversed
   * merge ranges. This workbook itself is left unchanged.
   *
   * @returns An object containing the repaired Workbook (an instance of this workbook's class),
   *   a validation report, and the number of repairs made.
   *
   * @example
   * ```ts
   * const { workbook, repairCount } = wb.repair();
   * console.log(`${repairCount} issues repaired`);
   * ```
   */
  repair(): { workbook: Workbook; report: ValidationReport; repairCount: number } {
    const wasm = this.#wasm();
    wasm.ensureInitialized();
    const result: RepairResult = wasm.wasmRepair(this.data);
    return {
      workbook: new (this.constructor as typeof Workbook)(result.workbook),
      report: result.report,
      repairCount: result.repairCount,
    };
  }

  // --- Image embedding ---

  /** Track next image ID for unique naming. */
  private imageCounter = 0;

  /**
   * Adds an image (PNG, JPEG or GIF bytes, matching `format`) anchored to a
   * cell range on the given sheet.
   * The image is embedded via `preservedEntries` and OOXML drawing XML.
   *
   * @param sheetName - Target sheet name.
   * @param anchor - Cell range to anchor the image to, as 0-based column and row indexes (`fromCol`, `fromRow`, `toCol`, `toRow`).
   * @param imageBytes - Raw image bytes.
   * @param format - Image format extension (default: `'png'`).
   * @throws ModernXlsxError with code `SHEET_NOT_FOUND` if no sheet has the
   *   exact name `sheetName`.
   * @throws {RangeError} If an anchor column is not an integer 0-16383, a row is not an integer
   *   0-1048575, an offset is not a non-negative safe integer, or the anchor ends before it
   *   starts; the workbook is left unchanged.
   *
   * @example
   * ```ts
   * const png = fs.readFileSync('logo.png');
   * // Cells B2 to E10: 0-based column and row indexes.
   * wb.addImage('Sheet1', { fromCol: 1, fromRow: 1, toCol: 4, toRow: 9 }, png);
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
    assertImageAnchor(anchor);

    this.data.preservedEntries ??= {};

    let highest = this.imageCounter;
    for (const key of Object.keys(this.data.preservedEntries)) {
      const n = MEDIA_IMAGE_PATH.exec(key)?.[1];
      if (n) highest = Math.max(highest, Number(n));
    }
    const imageId = highest + 1;
    this.imageCounter = imageId;
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
      const existing = TEXT_DECODER.decode(new Uint8Array(existingSheetRels));
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
   * @param anchor - Cell range to anchor the barcode image to, as 0-based column and row indexes (`fromCol`, `fromRow`, `toCol`, `toRow`).
   * @param value - Data to encode in the barcode.
   * @param options - Barcode type, rendering, and sizing options.
   * @throws {RangeError} If the anchor is invalid (see {@link Workbook.addImage}) or a rendering
   *   option is out of range (see `generateBarcode`).
   *
   * @example
   * ```ts
   * // Cells A1 to C5: 0-based column and row indexes.
   * wb.addBarcode('Sheet1', { fromCol: 0, fromRow: 0, toCol: 2, toRow: 4 }, '12345', {
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
   * Returns the live internal WorkbookData; mutations change this workbook.
   *
   * @returns The complete workbook data model as a plain JSON-serializable object.
   */
  toJSON(): WorkbookData {
    return this.data;
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
        name: 'Calibri',
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
  };
}

// ---------------------------------------------------------------------------
// Input validation
// ---------------------------------------------------------------------------

/** First character of a defined name: a letter, `_` or `\`. */
const DEFINED_NAME_START = /^[\p{L}_\\]/u;

/**
 * Every character of a defined name: a letter, combining mark (Unicode M,
 * never first: see {@link DEFINED_NAME_START}), digit, `_`, `.`, `\` or
 * `?`. Excel 365 (build 20527, 2026-11-12) takes `a?b`, and every name
 * probed with a combining mark after the first character (Thai, Devanagari,
 * Bengali and Tamil vowel signs, Arabic harakat, Hebrew points, decomposed
 * Latin and Vietnamese accents, U+20DD, U+FE0F) opens with the name kept in
 * a file modern-xlsx wrote. `Names.Add` takes each one except `\` + U+0301
 * (as with `\a`), and Excel's own save drops a lone U+FE0F.
 */
const DEFINED_NAME_REST = /^[\p{L}\p{M}\p{Nd}_.\\?]+$/u;

/** The `_xlnm.` built-in names `addNamedRange` accepts, lower-cased. */
const BUILTIN_DEFINED_NAMES: ReadonlySet<string> = new Set([
  '_xlnm.print_area',
  '_xlnm.print_titles',
  '_xlnm._filterdatabase',
]);

/**
 * Throws INVALID_ARGUMENT when `name` is not a well-formed defined name
 * (length, first and later characters, then the reference rules defined
 * names share with table names), checked in that order. Excel's `Names.Add`
 * refuses an A1 cell reference as a defined name, as `addTable` does for a
 * table, though Excel opens a file holding one; the character rules are
 * narrower than a table name's (no `€`, `·` or `°`, which Excel takes). As
 * for a table name, a combining mark may not come first, though Excel takes
 * a leading Thai mark in both.
 */
function checkDefinedNameShape(name: string): void {
  const invalid = (why: string) =>
    new ModernXlsxError(INVALID_ARGUMENT, `Invalid defined name "${name}": ${why}`);
  if (name.length < 1 || name.length > 255) {
    throw invalid('it must be 1-255 characters long');
  }
  if (!DEFINED_NAME_START.test(name)) {
    throw invalid('it must start with a letter, "_" or "\\"');
  }
  if (!DEFINED_NAME_REST.test(name)) {
    throw invalid('it may contain only letters, combining marks, digits, "_", ".", "\\" and "?"');
  }
  const problem = referenceNameProblem(name);
  if (problem) throw invalid(problem);
}

/**
 * Throws INVALID_ARGUMENT when `name` uses the reserved `_xlnm.` prefix for
 * anything but `Print_Area`, `Print_Titles` and `_FilterDatabase` (`key`
 * is the {@link nameKey} of `name`).
 */
function checkReservedDefinedName(name: string, key: string): void {
  if (!key.startsWith('_xlnm.')) return;
  const invalid = (why: string) =>
    new ModernXlsxError(INVALID_ARGUMENT, `Invalid defined name "${name}": ${why}`);
  if (!BUILTIN_DEFINED_NAMES.has(key)) {
    throw invalid("'_xlnm.' is reserved for Print_Area, Print_Titles and _FilterDatabase");
  }
}

/**
 * Token types whose source text can start with a sheet prefix that
 * `renameSheet` rewrites. An `error` token's value omits its prefix
 * (`Data!#REF!` has value `#REF!`), so prefixes are read from the source text.
 */
const SHEET_PREFIXED_TOKENS: ReadonlySet<TokenType> = new Set<TokenType>([
  'cell_ref',
  'column_range',
  'row_range',
  'name',
  'error',
]);

/**
 * A sheet name ignoring case (final sigma folded), not canonical
 * equivalence: its lower case with each U+03C2 made U+03C3, and each U+0130
 * kept. `toLowerCase` ends a word in U+03C2 where U+03A3 stood, and Excel
 * takes the two as one letter: its own rename of sheet U+039F U+0394 U+039F
 * U+03A3 rewrites a prefix that ends in U+03C3 as well as one that ends in
 * U+03C2. `toLowerCase` turns U+0130 into `i` + U+0307, a spelling Excel 365
 * does not take for U+0130 (2026-11-13): its rename of sheet U+0130 `a`
 * keeps a prefix `i` + U+0307 `a`. A name without U+0130 is lower-cased
 * whole, without splitting.
 */
const caseKey = (name: string): string =>
  name.includes('\u0130')
    ? name
        .split('\u0130')
        .map((part) => part.toLowerCase().replaceAll('\u03C2', '\u03C3'))
        .join('\u0130')
    : name.toLowerCase().replaceAll('\u03C2', '\u03C3');

/** One `renameSheet` call, built once for all the defined names it rewrites. */
interface SheetRename {
  /** The {@link caseKey} of the old sheet name. */
  readonly target: string;
  readonly newName: string;
  /** The {@link caseKey} of `newName`. */
  readonly newKey: string;
  /** Whether a span end names a sheet of the workbook, before the rename. */
  readonly isSheet: (end: string) => boolean;
}

/**
 * Rewrite the sheet prefixes in a defined-name value that name the old sheet
 * (`rename.target`; ignoring case (final sigma folded), not canonical
 * equivalence: Excel's own rename keeps a prefix spelled in another normal
 * form, which it does not take for the sheet, so {@link nameKey} is not used
 * here; see {@link caseKey}) to `rename.newName`, quoted when needed, as
 * Excel's own rename does: each end of a 3-D span that names it is rewritten
 * too, and the last end of an unquoted `Name:Sheet` prefix
 * ({@link renamedSheetSpan}). A single sheet name (not a 3-D span) right
 * after the range operator `:` is written quoted, as Excel stores it
 * (`First:'Data'!$C$10`). A prefix right after a `]` names a sheet of
 * another workbook (`[1]Sheet1!$A$1`, the form Excel stores for a reference
 * into an external workbook) and is kept; the quoted forms Excel stores,
 * `'[1]My Sheet'!$A$1` and `'C:/dir/[Book.xlsx]Sheet1'!$A$1`, hold the
 * bracket inside the quotes, and {@link renamedSheetSpan} keeps every prefix
 * that holds a bracket, 3-D spans (`'[1]My Sheet:Sheet3'!$A$1`) included.
 * Returns `undefined` when nothing changes,
 * including for values the tokenizer rejects. The parts are joined once, so
 * the time is linear in the value length.
 */
function renameSheetPrefixes(value: string, rename: SheetRename): string | undefined {
  const { tokens, errors } = tokenize(value);
  if (errors.length > 0) return undefined;
  const parts: string[] = [];
  let copied = 0;
  for (const token of tokens) {
    if (!SHEET_PREFIXED_TOKENS.has(token.type)) continue;
    if (value.charAt(token.start - 1) === ']') continue;
    const text = value.slice(token.start, token.end);
    // An unprefixed error literal (`#REF!`) has a `!` of its own.
    if (token.type === 'error' && text === token.value) continue;
    const prefix = readSheetPrefix(text);
    if (!prefix) continue;
    const inQuotes = text.startsWith("'");
    const afterColon = value.charAt(token.start - 1) === ':';
    const replacement = renamedSheetSpan(prefix.sheet, rename, inQuotes, afterColon);
    if (replacement === undefined) continue;
    parts.push(value.slice(copied, token.start), replacement);
    copied = token.start + prefix.length;
  }
  if (parts.length === 0) return undefined;
  parts.push(value.slice(copied));
  const result = parts.join('');
  return result === value ? undefined : result;
}

/**
 * The prefix (without its `!`) that replaces `sheet`, the unquoted sheet name
 * or 3-D span of a prefix, when the sheet `target` (its {@link caseKey}) is
 * renamed to `newName` (both from `rename`, as is `isSheet`); `undefined`
 * when `sheet` does not name `target`. A
 * prefix that holds `[` or `]` names another workbook (no sheet name holds a
 * bracket) and gives `undefined`. A span
 * (`Sheet1:Sheet3`; sheet names cannot contain `:`) has each end that names
 * `target` replaced, and a sheet between its ends is not named by it; the
 * span is quoted as one unit when either end needs quoting
 * (`'Sheet1:Q 3'`), the form Excel stores. Text with a colon is a span only
 * when both ends name sheets of the workbook (`isSheet`), as Excel reads it.
 * Otherwise a quoted prefix (`inQuotes`), as in `'C:/dir/Book.xlsx'` or
 * `'Nosuch:Sheet1'`, is a path into another workbook, and an unquoted one
 * whose first end is no sheet, as in `First:Sheet1!$C$10`, is a name, the
 * range operator and a sheet-qualified reference. Its last end is written
 * quoted, as Excel 365 stores it (2026-11-14), so the text cannot be read
 * as a span once a sheet takes the name: it is replaced when it names
 * `target` (`First:'First'` when Sheet1 is renamed to First), and kept but
 * quoted when the new name is the first end (`First:'Sheet1'` when Sheet3
 * is renamed to First); every other such prefix gives `undefined`. A sheet
 * name without a colon is replaced quoted when it follows the range
 * operator (`afterColon`, as in `First:'Sheet1'!$C$10` or
 * `A1:'Sheet1'!$B$2`), the form Excel stores, and otherwise quoted only
 * when needed.
 */
function renamedSheetSpan(
  sheet: string,
  rename: SheetRename,
  inQuotes: boolean,
  afterColon: boolean,
): string | undefined {
  const { target, newName, isSheet } = rename;
  if (sheet.includes('[') || sheet.includes(']')) return undefined;
  const colon = sheet.indexOf(':');
  if (colon < 0) {
    if (caseKey(sheet) !== target) return undefined;
    return afterColon ? quotedSheetName(newName) : quoteSheetName(newName);
  }
  const first = sheet.slice(0, colon);
  const last = sheet.slice(colon + 1);
  const firstKey = caseKey(first);
  const lastHit = caseKey(last) === target;
  if (!isSheet(first)) {
    return inQuotes ? undefined : renamedNameOperand(first, firstKey, last, lastHit, rename);
  }
  const firstHit = firstKey === target;
  if (!isSheet(last) || (!firstHit && !lastHit)) return undefined;
  return spanPrefix(firstHit ? newName : first, lastHit ? newName : last);
}

/** The 3-D span `start:end`, quoted as one unit when either end needs quoting. */
function spanPrefix(start: string, end: string): string {
  const span = `${start}:${end}`;
  const quoted = quoteSheetName(start) !== start || quoteSheetName(end) !== end;
  return quoted ? quotedSheetName(span) : span;
}

/**
 * The prefix that replaces the unquoted `first:last` of {@link renamedSheetSpan} when `first`
 * (its {@link caseKey} is `firstKey`) names no sheet: the name `first`, the range operator and a
 * reference on sheet `last`. The last end is written quoted, as Excel stores it: the new name
 * when `last` is the renamed sheet (`lastHit`), and `last` itself when the new name is spelled as
 * `first` and `last` is a sheet; `undefined` otherwise.
 */
function renamedNameOperand(
  first: string,
  firstKey: string,
  last: string,
  lastHit: boolean,
  rename: SheetRename,
): string | undefined {
  if (lastHit) return `${first}:${quotedSheetName(rename.newName)}`;
  if (firstKey === rename.newKey && rename.isSheet(last))
    return `${first}:${quotedSheetName(last)}`;
  return undefined;
}

/** `name` in single quotes with each inner `'` doubled, whether or not it needs quoting. */
const quotedSheetName = (name: string): string => `'${name.replaceAll("'", "''")}'`;

/**
 * Give cloned tables ids after `maxId` and {@link cloneTableName} names
 * (`${base}_${n}`, at most 255 UTF-16 units, `_` first when a cut base
 * would make it a reference), with the smallest `n >= 2` whose
 * {@link nameKey} is not in `taken` (the keys of the table and defined
 * names), which grows as names are handed out. `name` follows `displayName` when set.
 * Returns the {@link nameKey} of each old display name mapped to its new name.
 */
function reidTables(
  tables: TableDefinitionData[],
  maxId: number,
  taken: Set<string>,
): Map<string, string> {
  const renames = new Map<string, string>();
  let nextId = maxId;
  for (const table of tables) {
    nextId += 1;
    table.id = nextId;
    let n = 2;
    while (taken.has(nameKey(cloneTableName(table.displayName, n)))) n++;
    const unique = cloneTableName(table.displayName, n);
    taken.add(nameKey(unique));
    renames.set(nameKey(table.displayName), unique);
    table.displayName = unique;
    if (table.name != null) table.name = unique;
  }
  return renames;
}

/**
 * `${base}_${n}`, where `base` is `displayName` cut so the whole name is at
 * most {@link MAX_TABLE_NAME_LENGTH} UTF-16 units, one unit shorter when the
 * cut would split a surrogate pair. When the cut name is a reference (the
 * cut dropped the `.` that made `R0…01.a` a name, leaving the R1C1
 * reference `R0…01_2`), `base` is cut one unit more, by the same surrogate
 * rule, and put after a `_`, which no reference starts with.
 */
function cloneTableName(displayName: string, n: number): string {
  const suffix = `_${n}`;
  const cutAt = (limit: number): number => {
    const high = displayName.charCodeAt(limit - 1);
    return limit < displayName.length && high >= 0xd800 && high <= 0xdbff ? limit - 1 : limit;
  };
  const cut = cutAt(Math.min(displayName.length, MAX_TABLE_NAME_LENGTH - suffix.length));
  const name = displayName.slice(0, cut) + suffix;
  if (cut === displayName.length || referenceNameProblem(name) === undefined) return name;
  return `_${displayName.slice(0, cutAt(cut - 1))}${suffix}`;
}

/**
 * Rewrite the table name of each structured reference (`Table1[Amount]`) in
 * `formula` that `renames` maps (keyed by {@link nameKey}: Excel matches table
 * names ignoring case and canonical equivalence, so `cafe` + U+0301`[Qty]`
 * names table `café` with U+00E9) to its new name. Text inside string literals and names that
 * merely contain a table name are not references, so they stay. When the
 * tokenizer rejects the formula, it reads it again with the table names
 * masked ({@link maskTableNames}), so a name holding a character the
 * tokenizer cannot scan (`a\b`, `a?b`, `Ventes€`) is still found; a formula
 * rejected either way is returned as it is. The parts are joined once, so
 * the time is linear in the formula length.
 */
function renameTableRefs(formula: string, renames: ReadonlyMap<string, string>): string {
  let { tokens, errors } = tokenize(formula);
  if (errors.length > 0) {
    const masked = maskTableNames(formula);
    if (masked === undefined) return formula;
    ({ tokens, errors } = tokenize(masked));
    if (errors.length > 0) return formula;
  }
  const parts: string[] = [];
  let copied = 0;
  for (const token of tokens) {
    if (token.type !== 'structured_ref') continue;
    // A bare reference (`[@Amount]`) names no table: its bracket comes first.
    const nameLength = token.value.indexOf('[');
    if (nameLength < 1) continue;
    const nameEnd = token.start + nameLength;
    const renamed = renames.get(nameKey(formula.slice(token.start, nameEnd)));
    if (renamed === undefined) continue;
    parts.push(formula.slice(copied, token.start), renamed);
    copied = nameEnd;
  }
  if (parts.length === 0) return formula;
  parts.push(formula.slice(copied));
  return parts.join('');
}

/**
 * `formula` with each table-name position replaced by as many `_` as it has
 * UTF-16 units, so the tokenizer reads every table name as one it can scan
 * and its token positions are positions in `formula`; `undefined` when a
 * string literal, quoted sheet name or bracket is never closed. A table-name
 * position is the run of characters a table name may hold (none that
 * {@link TABLE_NAME_FORBIDDEN} matches) just before a `[` outside `"..."`
 * string literals, `'...'` quoted names and brackets. Inside brackets `'`
 * escapes the next character, as the tokenizer reads them.
 */
function maskTableNames(formula: string): string | undefined {
  const parts: string[] = [];
  let copied = 0;
  let runStart = 0;
  let i = 0;
  while (i < formula.length) {
    const ch = formula.charAt(i);
    let end = i + 1;
    if (ch === '"' || ch === "'") {
      end = quotedEnd(formula, i);
    } else if (ch === '[') {
      if (runStart < i) {
        parts.push(formula.slice(copied, runStart), '_'.repeat(i - runStart));
        copied = i;
      }
      end = bracketEnd(formula, i);
    } else if (!TABLE_NAME_FORBIDDEN.test(ch)) {
      i = end;
      continue;
    }
    if (end < 0) return undefined;
    i = end;
    runStart = end;
  }
  parts.push(formula.slice(copied));
  return parts.join('');
}

/**
 * Point the structured references of a sheet clone at its renamed tables
 * (`renames` from {@link reidTables}): every cell formula, and each table
 * column's `calculatedColumnFormula` and `totalsRowFormula`.
 */
function renameClonedTableRefs(clone: SheetData, renames: ReadonlyMap<string, string>): void {
  for (const cell of clone.worksheet.rows.flatMap((row) => row.cells)) {
    if (cell.formula) cell.formula = renameTableRefs(cell.formula, renames);
  }
  for (const column of (clone.worksheet.tables ?? []).flatMap((table) => table.columns)) {
    if (column.calculatedColumnFormula) {
      column.calculatedColumnFormula = renameTableRefs(column.calculatedColumnFormula, renames);
    }
    if (column.totalsRowFormula) {
      column.totalsRowFormula = renameTableRefs(column.totalsRowFormula, renames);
    }
  }
}

/**
 * Give cloned threaded comments new GUIDs in the braced upper-case form Excel
 * writes and requires (`{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`, see
 * {@link excelGuid}), whatever form the source ids had, remapping reply
 * `parentId`s. Returns the old-to-new id map.
 */
function reidThreadedComments(comments: ThreadedCommentData[]): Map<string, string> {
  const newIds = new Map<string, string>();
  for (const comment of comments) {
    const id = excelGuid();
    newIds.set(comment.id, id);
    comment.id = id;
  }
  for (const comment of comments) {
    if (comment.parentId !== undefined) {
      comment.parentId = newIds.get(comment.parentId) ?? comment.parentId;
    }
  }
  return newIds;
}
