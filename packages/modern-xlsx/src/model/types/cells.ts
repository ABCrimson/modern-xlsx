import type { RichTextRun } from './workbook.js';

// ---------------------------------------------------------------------------
// Read Options
// ---------------------------------------------------------------------------

/**
 * Security and size limits for reading a workbook.
 *
 * Raise a limit only for files you trust. Every field must be a finite number
 * greater than zero; the integer fields must also be safe integers. A value
 * above a field's ceiling is clamped to the ceiling, not rejected.
 */
export interface ReadLimits {
  /**
   * Maximum total decompressed size of all ZIP entries, in bytes.
   * Default and ceiling: 2,147,483,648 (2 GiB).
   */
  maxDecompressedSize?: number;
  /**
   * Maximum ratio of decompressed to compressed size per ZIP entry (unitless).
   * Default: 100. No ceiling. A sheet padded with whitespace can compress past 100:1.
   */
  maxCompressionRatio?: number;
  /** Maximum number of ZIP entries. Default: 100,000. Ceiling: 1,048,576. */
  maxEntries?: number;
  /**
   * Maximum length of the JSON the reader builds for the workbook, in UTF-8 bytes.
   * Default and ceiling: 536,870,888 (V8's maximum string length). UTF-8 bytes
   * are an upper bound on UTF-16 units, so a workbook with non-ASCII text can be
   * rejected slightly before V8's own limit. Exceeding it rejects with code
   * `SIZE_LIMIT`. It also bounds the total text a read copies from shared strings,
   * comment authors, hyperlink targets and worksheet namespace declarations into
   * the cells, comments, hyperlinks and kept worksheet extensions that name them:
   * the read rejects with `SIZE_LIMIT` as soon as those copies pass it. It also
   * bounds the memory of the records a read keeps in its lists (the defined names,
   * the shared string entries and rich text runs, the fonts, fills, borders, cell and
   * differential formats and other records of `xl/styles.xml`, the chart series and
   * axes, the threaded comments and mentions, then the worksheet lists: data
   * validations, conditional-format rules, selections, columns, hyperlinks, merged
   * ranges and the like), counted together over all of them and all sheets at each
   * record's size in memory: a part of millions of tiny records rejects with
   * `SIZE_LIMIT` before its JSON is built, and a workbook whose JSON
   * fits can be refused for it (in the npm build, measured on 2026-11-13,
   * 1,000,000 one-rule conditional formats use about 55% of the default).
   */
  maxJsonLength?: number;
}

/** Options for reading XLSX files. */
export interface ReadOptions {
  /** Password for encrypted XLSX files. */
  password?: string;
  /** Security and size limits; omitted fields keep their defaults. */
  limits?: ReadLimits;
}

/** Options for writing XLSX files. */
export interface WriteOptions {
  /** Password to encrypt the XLSX file with Agile Encryption (AES-256-CBC, SHA-512). */
  password?: string;
}

/** The date epoch system used by the workbook. `'date1900'` is the default for Windows Excel. */
export type DateSystem = 'date1900' | 'date1904';

/** The data type stored in a cell. */
export type CellType =
  | 'sharedString'
  | 'number'
  | 'boolean'
  | 'error'
  | 'formulaStr'
  | 'inlineStr'
  | 'stub';

/** Fill pattern type for cell backgrounds. */
export type PatternType =
  | 'none'
  | 'solid'
  | 'gray125'
  | 'darkGray'
  | 'mediumGray'
  | 'lightGray'
  | 'darkHorizontal'
  | 'darkVertical'
  | 'darkDown'
  | 'darkUp'
  | 'darkGrid'
  | 'darkTrellis'
  | 'lightHorizontal'
  | 'lightVertical'
  | 'lightDown'
  | 'lightUp'
  | 'lightGrid'
  | 'lightTrellis';

/** Border line style for cell borders. */
export type BorderStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

// ---------------------------------------------------------------------------
// Colour
// ---------------------------------------------------------------------------

/**
 * A SpreadsheetML colour (`CT_Color`): an ARGB value, a theme slot, a legacy
 * palette index or the automatic colour, with an optional tint. Every attribute
 * is optional and absent ones are omitted.
 *
 * Properties named `colorSpec`, `fgColorSpec`, `bgColorSpec` and `colorSpecs`
 * carry the full colour next to the legacy rgb-only string property
 * (`color`, `fgColor`, `tabColor`, ...). On write the spec wins only while
 * `spec.rgb` equals that string (an empty string counts as absent); otherwise
 * the string is written as plain `rgb`. So to change a colour, change the
 * string and drop or update the spec.
 */
export interface ColorData {
  /** ARGB hex, e.g. `'FF4472C4'`. */
  rgb?: string;
  /** Zero-based index into the theme's colour scheme. */
  theme?: number;
  /** Index into the legacy indexed palette (64 is the system foreground, 65 the system background). */
  indexed?: number;
  /** Lightens (positive) or darkens (negative) the colour, from -1 to 1. */
  tint?: number;
  /** The automatic (system) colour. */
  auto?: boolean;
}

// ---------------------------------------------------------------------------
// Cell & Row
// ---------------------------------------------------------------------------

/** Raw cell data as stored in the workbook data model. */
export interface CellData {
  /** A1-style cell reference (e.g., `'B3'`). */
  reference: string;
  /** The data type of the cell value. */
  cellType: CellType;
  /** Zero-based index into `StylesData.cellXfs`, or `null` for default. */
  styleIndex: number | null;
  /** The cell value as a string (numbers are stored as string representations). */
  value: string | null;
  /** The formula string (without leading `=`), or `null`. */
  formula: string | null;
  /**
   * The `t` attribute of `<f>`, read and written verbatim (such as `'shared'`, `'array'` or
   * `'dataTable'`); omitted on read when `<f>` has none.
   */
  formulaType?: string | null;
  /** The `ref` range of a shared, array or data-table formula. */
  formulaRef?: string | null;
  /** Shared formula group index. */
  sharedIndex?: number | null;
  /** Inline string content (used for `inlineStr` cell type). */
  inlineString?: string | null;
  /**
   * `<c cm>`: the 1-based index of the cell's record in the workbook's cell
   * metadata (`xl/metadata.xml`). Excel gives a dynamic-array formula `cm="1"`.
   */
  cellMetadata?: number | null;
  /**
   * Whether this is a dynamic array formula: set on read whenever `<c>` carries
   * `cm`; on write it yields `cm="1"` when `cellMetadata` is unset.
   *
   * @deprecated Removed in 2.0; use {@link CellData.cellMetadata}.
   */
  dynamicArray?: boolean | null;
  /** First input cell of a data table (`r1` on `<f>`). */
  formulaR1?: string | null;
  /** Second input cell of a data table (`r2` on `<f>`). */
  formulaR2?: string | null;
  /** Two-dimensional data table flag. */
  formulaDt2d?: boolean | null;
  /** One-dimensional data table is a row (`dtr` on `<f>`): `formulaR1` is a row input cell. */
  formulaDtr?: boolean | null;
  /** Data table input cell 1 was deleted (`del1` on `<f>`). */
  formulaDel1?: boolean | null;
  /** Data table input cell 2 was deleted (`del2` on `<f>`). */
  formulaDel2?: boolean | null;
  /**
   * Set on read when `<f>` carries `del1="1"` or the non-schema `dtr1="1"`; on
   * write it yields `del1="1"` when `formulaDel1` is unset.
   *
   * @deprecated Removed in 2.0; use {@link CellData.formulaDel1}.
   */
  formulaDtr1?: boolean | null;
  /**
   * Set on read when `<f>` carries `del2="1"` or the non-schema `dtr2="1"`; on
   * write it yields `del2="1"` when `formulaDel2` is unset.
   *
   * @deprecated Removed in 2.0; use {@link CellData.formulaDel2}.
   */
  formulaDtr2?: boolean | null;
  /**
   * Rich text runs for mixed-format text within a cell. On read, a shared-string
   * cell gets the runs of the shared string its index names; inline-string runs
   * are not modelled. Set it to write rich text.
   */
  richText?: RichTextRun[];
}

/** Raw row data containing cells and row-level properties. */
export interface RowData {
  /** 1-based row index in the worksheet. */
  index: number;
  /** The cells in this row. */
  readonly cells: CellData[];
  /** Row height in points, or `null` for default height. */
  height: number | null;
  /**
   * Whether the height was set by hand. Excel keeps `height` only when this is
   * true; without it Excel sizes the row to its content.
   */
  customHeight?: boolean;
  /** Whether this row is hidden. */
  hidden: boolean;
  /** Outline (grouping) level (1-7), or `null` if ungrouped. */
  outlineLevel?: number | null;
  /** Whether this grouped row is collapsed. */
  collapsed?: boolean;
  /**
   * Zero-based index into `StylesData.cellXfs` of the row's format (`<row s>`),
   * which Excel applies to the row's empty cells. Read only when the row has
   * `customFormat="1"`, and written back with it.
   */
  styleIndex?: number | null;
}
