import { isDateFormatCode, isDateFormatId, serialToDate } from '../format/dates.js';
import { getBuiltinFormat } from '../format/format-cell.js';
import { INVALID_CELL_VALUE, ModernXlsxError } from './errors.js';
import type { CellData, CellType, DateSystem, RichTextRun, StylesData } from './types.js';

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
  /** @internal */ private readonly dateSystem: (() => DateSystem) | undefined;

  /**
   * Wraps an existing cell data object. Typically obtained via `Worksheet.cell()`.
   *
   * @param data - The underlying cell data model.
   * @param styles - Optional shared styles for number format resolution.
   * @param dateSystem - Internal: reads the owning workbook's date system for
   *   {@link Cell.dateValue}; absent means the 1900 system.
   */
  constructor(data: CellData, styles?: StylesData, dateSystem?: () => DateSystem) {
    this.data = data;
    this.styles = styles;
    this.dateSystem = dateSystem;
  }

  /** Returns the A1-style cell reference (e.g., `"B3"`). */
  get reference(): string {
    return this.data.reference;
  }

  /**
   * Returns the cell's data type (e.g., `"number"`, `"sharedString"`, `"boolean"`, `"error"`).
   * A formula cell takes the type of its cached value: `"formulaStr"` only when
   * that value is a string.
   */
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
   * Any rich text on the cell is cleared. With a formula set, the value is the
   * cached result and the type follows it: `"formulaStr"` for a string,
   * `"number"` for a number or `null`, `"boolean"` for a boolean; an `"error"`
   * formula cell keeps its type. A `null` value on a cell without a formula makes it a
   * `"number"` cell.
   *
   * On a shared-formula child (no formula text, but `formulaType` `"shared"` or a
   * `sharedIndex`), any value, `null` included, replaces the formula: `formulaType`,
   * `formulaRef` and `sharedIndex` are removed. A cell with formula text keeps them.
   *
   * @throws ModernXlsxError with code `INVALID_CELL_VALUE` for `NaN` or an infinity.
   */
  set value(val: string | number | boolean | null) {
    if (typeof val === 'number' && !Number.isFinite(val)) {
      throw new ModernXlsxError(
        INVALID_CELL_VALUE,
        `Cell value must be a finite number, got ${val}`,
      );
    }
    delete this.data.richText;
    if (this.#isSharedChild()) this.#dropFormulaKeys();
    if (val === null) {
      this.data.value = null;
      this.data.cellType = 'number';
    } else if (typeof val === 'string') {
      this.data.value = val;
      // An error formula cell keeps its type; any other string cell is a shared string.
      if (!(this.data.formula && this.data.cellType === 'error'))
        this.data.cellType = 'sharedString';
    } else if (typeof val === 'number') {
      this.data.value = val.toString();
      this.data.cellType = 'number';
    } else {
      this.data.value = val ? '1' : '0';
      this.data.cellType = 'boolean';
    }
    this.#syncFormulaType();
  }

  /**
   * With a formula set, gives the cell the type of its cached value: a string
   * is `"formulaStr"`, a number or no value is `"number"`, a boolean stays
   * `"boolean"`, and `"error"` is kept.
   */
  #syncFormulaType(): void {
    if (!this.data.formula || this.data.cellType === 'error') return;
    if (this.data.value == null) this.data.cellType = 'number';
    else if (this.data.cellType === 'sharedString') this.data.cellType = 'formulaStr';
  }

  /**
   * A shared-formula child for the value and rich-text setters: no formula text,
   * but `formulaType` `"shared"` or a `sharedIndex`, so its formula is the master's.
   */
  #isSharedChild(): boolean {
    return (
      !this.data.formula && (this.data.formulaType === 'shared' || this.data.sharedIndex != null)
    );
  }

  /**
   * A shared-formula child for the formula setter: `formulaType` `"shared"` or a
   * `sharedIndex`, and no `formulaRef` (which only a master has), with or without
   * formula text of its own.
   */
  #isSharedFormulaChild(): boolean {
    return (
      (this.data.formulaType === 'shared' || this.data.sharedIndex != null) && !this.data.formulaRef
    );
  }

  /** Removes the `formulaType`, `formulaRef` and `sharedIndex` of the `<f>`. */
  #dropFormulaKeys(): void {
    delete this.data.formulaType;
    delete this.data.formulaRef;
    delete this.data.sharedIndex;
  }

  /** Returns the cell's formula text as stored, or `null` if no formula is set. A formula read from a file has no leading `=`; text given to the setter is kept as given, and the writer drops one leading `=` when it writes `<f>`. */
  get formula(): string | null {
    return this.data.formula ?? null;
  }

  /**
   * Sets or clears the cell's formula. A formula cell takes the type of its
   * cached value (see {@link Cell.value}): `"formulaStr"` for a string,
   * `"number"` for a number or no value. Clearing the formula (`null`, or an empty
   * string, which is stored as `null` so the getter returns `null`) removes
   * `formulaType`, `formulaRef` and `sharedIndex` with it, and turns a `"formulaStr"`
   * cell into `"sharedString"`, or `"number"` when it holds no value.
   * Formula text set on a shared-formula child (`formulaType` `"shared"` or a
   * `sharedIndex`, and no `formulaRef`, whether or not it carries formula text of its
   * own) gives that child a plain formula of its own: its `formulaType` and
   * `sharedIndex` are removed. New text on any other cell, a shared-formula master
   * (one with a `formulaRef`) included, keeps those keys. The children of a master
   * follow its text: a new text is theirs on write, and once the master's formula is
   * cleared, a child without text of its own is written with its cached value alone
   * (empty when it holds none) and a child with text of its own as a plain formula.
   */
  set formula(f: string | null) {
    const child = this.#isSharedFormulaChild();
    this.data.formula = f || null;
    if (f) {
      if (child) this.#dropFormulaKeys();
      this.#syncFormulaType();
      return;
    }
    this.#dropFormulaKeys();
    if (this.data.cellType === 'formulaStr') {
      this.data.cellType = this.data.value == null ? 'number' : 'sharedString';
    }
  }

  /**
   * Returns the resolved number format code (e.g., `"#,##0.00"`, `"yyyy-mm-dd"`),
   * or `null` if the cell uses the default General format.
   *
   * The workbook's own `numFmts` win, then the built-in table. A built-in date id
   * the table has no code for (27-36, 50-55, 57 and 58) resolves to the code of id
   * 14. Other ids without a code (5-8, 23-26 and 41-44) return `null` unless the
   * file defines them.
   */
  get numberFormat(): string | null {
    if (!this.styles || this.data.styleIndex == null) return null;
    const xf = this.styles.cellXfs[this.data.styleIndex];
    if (!xf || xf.numFmtId === 0 || xf.numFmtId == null) return null;

    const custom = this.styles.numFmts?.find((nf) => nf.id === xf.numFmtId);
    if (custom) return custom.formatCode;
    const builtin = getBuiltinFormat(xf.numFmtId);
    if (builtin && builtin !== 'General') return builtin;
    return isDateFormatId(xf.numFmtId) ? (getBuiltinFormat(14) ?? null) : null;
  }

  /**
   * Returns a `Date` if this cell contains a date-formatted number, or `null` otherwise.
   * Uses the cell's number format to determine whether the numeric value
   * represents a date, and the workbook's date system (1900 or 1904) for the epoch.
   */
  get dateValue(): Date | null {
    if (this.data.cellType !== 'number' || this.data.value == null) return null;
    const fmt = this.numberFormat;
    if (!fmt || !isDateFormatCode(fmt)) return null;
    return serialToDate(Number.parseFloat(this.data.value), this.dateSystem?.() ?? 'date1900');
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
   * On a shared-formula child (no formula text, but `formulaType` `"shared"` or a
   * `sharedIndex`), the rich text replaces the formula, as a value does:
   * `formulaType`, `formulaRef` and `sharedIndex` are removed.
   *
   * Pass `undefined` to clear rich text.
   */
  set richText(runs: readonly RichTextRun[] | undefined) {
    if (runs === undefined) {
      delete this.data.richText;
      return;
    }
    if (this.#isSharedChild()) this.#dropFormulaKeys();
    this.data.richText = runs.map((r) => ({ ...r }));
    // Auto-set the cell type and plain-text value.
    this.data.cellType = 'sharedString';
    this.data.value = runs.map((r) => r.text).join('');
  }
}
