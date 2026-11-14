import type { ColorData, RowData } from './cells.js';
import type { WorksheetChartData } from './charts.js';
import type { PivotTableData, SlicerData, TimelineData } from './pivot.js';

// ---------------------------------------------------------------------------
// Worksheet structures
// ---------------------------------------------------------------------------

/** Frozen pane configuration for locking rows/columns. */
export interface FrozenPane {
  /** Number of rows to freeze from the top. */
  rows: number;
  /** Number of columns to freeze from the left. */
  cols: number;
}

/** A split (not frozen) `<pane>` of a `<sheetView>`, dividing the view into two or four scrollable regions. */
export interface SplitPaneData {
  /** Horizontal split position in twips (ySplit). */
  horizontal?: number | null;
  /** Vertical split position in twips (xSplit). */
  vertical?: number | null;
  /** Cell reference for top-left cell in bottom-right pane. */
  topLeftCell?: string | null;
  /** Active pane: `"topLeft"` | `"topRight"` | `"bottomLeft"` | `"bottomRight"`. */
  activePane?: string | null;
}

/** A `<selection>` of a `<sheetView>`: the active cell and selected range of one pane. */
export interface PaneSelectionData {
  /** Which pane this selection belongs to. */
  pane?: string | null;
  /** Active cell reference, e.g. `"A1"`. */
  activeCell?: string | null;
  /** Selected range, e.g. `"A1:C5"`. */
  sqref?: string | null;
}

/** The view mode for a worksheet tab. */
export type ViewMode = 'normal' | 'pageBreakPreview' | 'pageLayout';

/** The visibility state of a worksheet tab. */
export type SheetState = 'visible' | 'hidden' | 'veryHidden';

/** The `<sheetView>` attributes of a worksheet (ECMA-376 Part 1 §18.3.1.87). */
export interface SheetViewData {
  /** Whether grid lines are visible (default: true). */
  showGridLines?: boolean;
  /** Whether row and column headers are visible (default: true). */
  showRowColHeaders?: boolean;
  /** Whether zero values are displayed (default: true). */
  showZeros?: boolean;
  /** Right-to-left display mode (default: false). */
  rightToLeft?: boolean;
  /** Whether this sheet tab is selected (default: false). */
  tabSelected?: boolean;
  /** Whether the ruler is shown in Page Layout view (default: true). */
  showRuler?: boolean;
  /** Whether outline (grouping) symbols are shown (default: true). */
  showOutlineSymbols?: boolean;
  /** Whether white space around the page is shown in Page Layout view (default: true). */
  showWhiteSpace?: boolean;
  /** Whether the default grid color is used (default: true). */
  defaultGridColor?: boolean;
  /** Zoom percentage (10–400). */
  zoomScale?: number | null;
  /** Zoom percentage for Normal view. */
  zoomScaleNormal?: number | null;
  /** Zoom percentage for Page Layout view. */
  zoomScalePageLayoutView?: number | null;
  /** Zoom percentage for Page Break Preview. */
  zoomScaleSheetLayoutView?: number | null;
  /**
   * `colorId`: the gridline colour as an index into the legacy indexed palette (see
   * {@link ColorData.indexed}), not a theme colour. Read and written verbatim.
   */
  colorId?: number | null;
  /** View mode. */
  view?: ViewMode | null;
}

/** Column definition for width, visibility, and grouping. */
export interface ColumnInfo {
  /** 1-based first column index of this span. */
  min: number;
  /** 1-based last column index of this span. */
  max: number;
  /** Column width in Excel character units. */
  width: number;
  /** Whether this column is hidden. */
  hidden: boolean;
  /** Whether the width was explicitly set (vs. default). */
  customWidth: boolean;
  /** Outline (grouping) level (1-7), or `null` if ungrouped. */
  outlineLevel?: number | null;
  /** Whether this grouped column is collapsed. */
  collapsed?: boolean;
  /** Index into `StylesData.cellXfs` that Excel applies to the column's empty cells and shows as its format. */
  styleIndex?: number;
  /** Whether the column is "best fit": Excel widens it to show numbers typed into its cells. */
  bestFit?: boolean;
  /** Whether Excel shows the phonetic (furigana) guide for the column's cells. */
  phonetic?: boolean;
}

/** Auto-filter dropdown configuration for a worksheet range. */
export interface AutoFilterData {
  /** The A1-style range covered by the auto-filter (e.g., `'A1:D10'`). */
  range: string;
  /** Per-column filter criteria. */
  filterColumns?: readonly FilterColumnData[];
}

/** A `<filterColumn>` of an auto-filter: a list of values to show, or custom filter conditions. */
export interface FilterColumnData {
  colId: number;
  filters?: readonly string[];
  customFilters?: CustomFiltersData | null;
}

/** The `<customFilters>` of a filter column: conditions joined by AND or OR. */
export interface CustomFiltersData {
  /**
   * The `and` attribute: `true` requires every condition to match (AND); `false` or
   * omitted, any one (OR). Written as `and="1"` only when `true`.
   */
  andOp?: boolean;
  filters: readonly CustomFilterData[];
}

/** A `<customFilter>` condition: an operator and the value it compares against. */
export interface CustomFilterData {
  /** The `operator` attribute, read and written verbatim; omitted on read when absent. */
  operator?: string | null;
  val: string;
}

/** Hyperlink attached to a cell. */
export interface HyperlinkData {
  /** The cell reference this hyperlink is attached to (e.g., `'A1'`). */
  cellRef: string;
  /**
   * External URL, `mailto:` address or file (e.g., `'https://example.com'`), stored as a
   * sheet relationship with `TargetMode="External"` and referenced by `r:id`.
   */
  target?: string | null;
  /**
   * In-workbook reference (e.g., `'Sheet2!A1'` or a defined name), or, together with
   * `target`, a fragment within that target.
   */
  location?: string | null;
  /** Display text shown in the cell. */
  display?: string | null;
  /** Tooltip text shown on hover. */
  tooltip?: string | null;
}

/**
 * Page margins in inches (`<pageMargins>`). The writer fills an unset side with Excel's
 * Normal preset: left and right 0.7, top and bottom 0.75, header and footer 0.3.
 */
export interface PageMarginsData {
  top?: number | null;
  bottom?: number | null;
  left?: number | null;
  right?: number | null;
  header?: number | null;
  footer?: number | null;
}

/** Print setup: the `<pageSetup>` attributes, the sheet's `<pageMargins>` (`margins`) and `<sheetPr><pageSetUpPr fitToPage>` (`fitToPage`). */
export interface PageSetupData {
  paperSize?: number | null;
  orientation?: 'portrait' | 'landscape' | null;
  fitToWidth?: number | null;
  fitToHeight?: number | null;
  scale?: number | null;
  firstPageNumber?: number | null;
  horizontalDpi?: number | null;
  verticalDpi?: number | null;
  /**
   * `<sheetPr><pageSetUpPr fitToPage>`. Excel applies `fitToWidth` and
   * `fitToHeight` only when this is `true`; the writer emits the element for
   * `true` only.
   */
  fitToPage?: boolean | null;
  margins?: PageMarginsData | null;
}

/** A `<brk>` page break in `<rowBreaks>` or `<colBreaks>`. */
export interface PageBreakData {
  id: number;
  min?: number | null;
  max?: number | null;
  /** `man`: a manual page break. */
  man?: boolean;
}

/** The worksheet's row and column page breaks (`<rowBreaks>` and `<colBreaks>`). */
export interface PageBreaksData {
  rowBreaks?: PageBreakData[];
  colBreaks?: PageBreakData[];
}

/**
 * Sheet protection: the `<sheetProtection>` element (ECMA-376 Part 1 §18.3.1.85).
 *
 * Every flag except `sheet` says whether an action is **locked** while the sheet is
 * protected: `true` blocks the action in Excel, `false` allows it. Excel enforces the
 * flags only when `sheet` is `true`. An optional flag left out takes the default shown
 * on it, and a workbook read back omits an optional flag that is at its default.
 *
 * To protect with a password, set `plainPassword`. `password` and the `algorithmName`,
 * `hashValue`, `saltValue` and `spinCount` group hold the hash a file carries; they are
 * written back unchanged, and a `password` that is not 1 to 4 hex digits fails the write.
 */
export interface SheetProtectionData {
  /** `true` protects the sheet: locked cells cannot be edited and the flags below apply. */
  sheet: boolean;
  /** `true` locks shapes, charts and other objects; `false` allows editing them (schema default `false`). */
  objects: boolean;
  /** `true` locks scenarios; `false` allows editing them (schema default `false`). */
  scenarios: boolean;
  /**
   * The legacy 16-bit password hash a file stores (1 to 4 hex digits), never the password
   * itself: any other value makes the write throw `INVALID_WORKBOOK` and `validate()` report
   * it, and it is not hashed for you, because a password such as `'cafe'` is also a valid hash.
   */
  password?: string | null;
  /** `true` locks formatting cells; `false` allows it (schema default `true`). */
  formatCells: boolean;
  /** `true` locks formatting columns; `false` allows it (schema default `true`). */
  formatColumns: boolean;
  /** `true` locks formatting rows; `false` allows it (schema default `true`). */
  formatRows: boolean;
  /** `true` locks inserting columns; `false` allows it (schema default `true`). */
  insertColumns: boolean;
  /** `true` locks inserting rows; `false` allows it (schema default `true`). */
  insertRows: boolean;
  /** `true` locks deleting columns; `false` allows it (schema default `true`). */
  deleteColumns: boolean;
  /** `true` locks deleting rows; `false` allows it (schema default `true`). */
  deleteRows: boolean;
  /** `true` locks sorting; `false` allows it (schema default `true`). */
  sort: boolean;
  /** `true` locks AutoFilter use; `false` allows it (schema default `true`). */
  autoFilter: boolean;
  /** `true` locks inserting hyperlinks; `false` allows it. Omitted means `true`. */
  insertHyperlinks?: boolean;
  /** `true` locks PivotTable use; `false` allows it. Omitted means `true`. */
  pivotTables?: boolean;
  /** `true` blocks selecting locked cells; `false` allows it. Omitted means `false`. */
  selectLockedCells?: boolean;
  /** `true` blocks selecting unlocked cells; `false` allows it. Omitted means `false`. */
  selectUnlockedCells?: boolean;
  /** Hash algorithm of `hashValue` as stored in the file, such as `'SHA-512'`. */
  algorithmName?: string | null;
  /** Base64 iterated password hash as stored in the file. */
  hashValue?: string | null;
  /** Base64 salt of `hashValue` as stored in the file. */
  saltValue?: string | null;
  /** Hash iterations of `hashValue` as stored in the file. */
  spinCount?: number | null;
  /**
   * The password to protect the sheet with, the only field that takes one: on write it
   * replaces `password` and the stored hash with a fresh hash (SHA-512, random 16-byte salt,
   * 100,000 iterations, as Excel writes it; `modern-xlsx/lite` writes the legacy 16-bit
   * `password` hash and accepts ASCII passwords only). Write-only: a workbook read back
   * carries the hash, never this field.
   */
  plainPassword?: string | null;
}

/** A `<dataValidation>` rule applied to the cells in `sqref`. */
export interface DataValidationData {
  sqref: string;
  /** The `type` attribute, such as `'list'` or `'whole'`. */
  validationType: string | null;
  operator: string | null;
  /**
   * The `errorStyle` attribute: what invalid input does. `'stop'` rejects it;
   * `'warning'` and `'information'` let the user keep it. Absent means `'stop'`.
   */
  errorStyle?: 'stop' | 'warning' | 'information' | null;
  /**
   * The `imeMode` attribute: the input method editor mode while a cell is
   * selected, such as `'off'`, `'on'` or `'hiragana'`. Absent means `'noControl'`.
   */
  imeMode?: string | null;
  /**
   * The `showDropDown` attribute of a `'list'` validation. Despite its name,
   * `true` HIDES the in-cell drop-down arrow; absent or `false` shows it.
   */
  showDropDown?: boolean | null;
  formula1: string | null;
  formula2: string | null;
  allowBlank: boolean | null;
  showErrorMessage: boolean | null;
  errorTitle: string | null;
  errorMessage: string | null;
  showInputMessage?: boolean | null;
  promptTitle?: string | null;
  prompt?: string | null;
}

// ---------------------------------------------------------------------------
// Conditional formatting
// ---------------------------------------------------------------------------

/** A `<cfvo>` threshold of a colour scale, data bar or icon set: its `type` and optional `val`. */
export interface CfvoData {
  cfvoType: string;
  val?: string | null;
}

/** A `<colorScale>` conditional format. */
export interface ColorScaleData {
  readonly cfvos: CfvoData[];
  /** One entry per colour stop: its `rgb`, or `''` when the stop has none. */
  readonly colors: string[];
  /**
   * The full colour of every stop, in stop order, filled only when some stop has
   * a theme, indexed, tint or auto attribute. Stop `i` uses `colorSpecs[i]` on
   * write while its `rgb` equals `colors[i]`; otherwise `colors[i]` is written
   * as plain `rgb`. A stop with neither is written with the default
   * `FF638EC6`, so the stop count is kept.
   */
  readonly colorSpecs?: ColorData[];
}

/** A `<dataBar>` conditional format. */
export interface DataBarData {
  readonly cfvos: CfvoData[];
  /** The `rgb` of the bar colour, or `''` when it has none. */
  color: string;
  /**
   * The full bar colour when it has a theme, indexed, tint or auto attribute;
   * same write rule as {@link ColorScaleData.colorSpecs}. A bar with neither is
   * written with Excel's default `FF638EC6`.
   */
  colorSpec?: ColorData | null;
}

/** An `<iconSet>` conditional format. */
export interface IconSetData {
  /** The `iconSet` attribute, the name of the icon set. */
  iconSetType?: string | null;
  readonly cfvos: CfvoData[];
}

/** A `<cfRule>` of a conditional-formatting block. */
export interface ConditionalFormattingRuleData {
  /** The `type` attribute. */
  ruleType: string;
  priority: number;
  operator: string | null;
  /** The first `<formula>` child. */
  formula: string | null;
  /**
   * The second and third `<formula>` children (a `cellIs` `'between'` or `'notBetween'` rule's
   * upper bound is the second), written after `formula` in order. A `<cfRule>` holds at most
   * three formulas: a rule with more fails the write with `INVALID_WORKBOOK` and is reported by
   * `validate()`. An empty `<formula/>` reads as `''` in its position. Absent when the rule has
   * one formula or none.
   */
  additionalFormulas?: string[];
  /** Zero-based index into `StylesData.dxfs` of the format the rule applies. */
  dxfId: number | null;
  /** `stopIfTrue`: no rule of lower priority applies once this one is true. Absent means false. */
  stopIfTrue?: boolean | null;
  /** `aboveAverage` of an `'aboveAverage'` rule: `false` highlights cells below the average. Absent means true. */
  aboveAverage?: boolean | null;
  /** `percent` of a `'top10'` rule: `rank` is a percentage, not a count. Absent means false. */
  percent?: boolean | null;
  /** `bottom` of a `'top10'` rule: the bottom `rank` values, not the top. Absent means false. */
  bottom?: boolean | null;
  /** `text` of a `'containsText'`, `'notContainsText'`, `'beginsWith'` or `'endsWith'` rule. */
  text?: string | null;
  /** `timePeriod` of a `'timePeriod'` rule, such as `'lastWeek'` or `'today'`. */
  timePeriod?: string | null;
  /** `rank` of a `'top10'` rule: how many values (or what percentage, with `percent`) match. */
  rank?: number | null;
  /** `stdDev` of an `'aboveAverage'` rule: how many standard deviations from the average. */
  stdDev?: number | null;
  /** `equalAverage` of an `'aboveAverage'` rule: values equal to the average also match. Absent means false. */
  equalAverage?: boolean | null;
  colorScale?: ColorScaleData | null;
  dataBar?: DataBarData | null;
  iconSet?: IconSetData | null;
}

/** A `<conditionalFormatting>` block: the rules applied to the cells in `sqref`. */
export interface ConditionalFormattingData {
  sqref: string;
  readonly rules: ConditionalFormattingRuleData[];
}

// ---------------------------------------------------------------------------
// Table definitions (Excel ListObjects)
// ---------------------------------------------------------------------------

/** Totals row aggregate functions supported by Excel tables. */
export type TotalsRowFunction =
  | 'none'
  | 'sum'
  | 'min'
  | 'max'
  | 'average'
  | 'count'
  | 'countNums'
  | 'stdDev'
  | 'var'
  | 'custom';

/** A `<tableColumn>` of a table. */
export interface TableColumnData {
  /** Column id, unique within the table. */
  id: number;
  /** Column name, the text of its header cell. */
  name: string;
  /**
   * Totals-row function: `'sum'`, `'min'`, `'max'`, `'average'`, `'count'`, `'countNums'`,
   * `'stdDev'`, `'var'`, `'custom'` (uses `totalsRowFormula`) or `'none'`.
   * `Worksheet.addTable` writes the matching totals-row cell when it is empty: the
   * `SUBTOTAL` formula Excel writes for a function (`SUBTOTAL(109,Table1[Amount])` for
   * `'sum'`), or the `totalsRowFormula` for `'custom'`.
   */
  totalsRowFunction?: string | null;
  /**
   * `totalsRowLabel`: the text of the column's totals-row cell. `Worksheet.addTable`
   * writes it into the cell when the column has no function and the cell is empty.
   */
  totalsRowLabel?: string | null;
  calculatedColumnFormula?: string | null;
  /** `<totalsRowFormula>`: the formula of the column's totals-row cell (without a leading `=`). */
  totalsRowFormula?: string | null;
  /** Zero-based index into `StylesData.dxfs` of the header cell's format. */
  headerRowDxfId?: number | null;
  /** Zero-based index into `StylesData.dxfs` of the data cells' format. */
  dataDxfId?: number | null;
  /** Zero-based index into `StylesData.dxfs` of the totals-row cell's format. */
  totalsRowDxfId?: number | null;
}

/** A table's `<tableStyleInfo>`: its style and which stripes and emphasis columns show. */
export interface TableStyleInfoData {
  /** Table style name, such as `'TableStyleMedium2'`. */
  name?: string | null;
  showFirstColumn: boolean;
  showLastColumn: boolean;
  showRowStripes: boolean;
  showColumnStripes: boolean;
}

/**
 * An Excel table (ListObject), `xl/tables/table{n}.xml`. Attributes and child elements
 * not listed here, such as `<sortState>`, are dropped on read.
 */
export interface TableDefinitionData {
  /** Table id, unique across the workbook. */
  id: number;
  name?: string | null;
  /** The name structured references use. */
  displayName: string;
  /** The table's range, header row included (e.g. `'A1:D10'`). */
  ref: string;
  /** Number of header rows; `0` for a table without a header row. */
  headerRowCount: number;
  totalsRowCount: number;
  totalsRowShown: boolean;
  readonly columns: TableColumnData[];
  styleInfo?: TableStyleInfoData | null;
  /** The range of the table's `<autoFilter>`. */
  autoFilterRef?: string | null;
}

// ---------------------------------------------------------------------------
// Headers & Footers
// ---------------------------------------------------------------------------

/** The worksheet's `<headerFooter>`: print header and footer text for odd, even and first pages. */
export interface HeaderFooterData {
  oddHeader?: string | null;
  oddFooter?: string | null;
  evenHeader?: string | null;
  evenFooter?: string | null;
  firstHeader?: string | null;
  firstFooter?: string | null;
  /** Use `evenHeader`/`evenFooter` on even pages. Omitted means `false`. */
  differentOddEven?: boolean;
  /** Use `firstHeader`/`firstFooter` on the first page. Omitted means `false`. */
  differentFirst?: boolean;
  /** Scale the header and footer with the document. Omitted means `true`. */
  scaleWithDoc?: boolean;
  /** Align the header and footer with the page margins. Omitted means `true`. */
  alignWithMargins?: boolean;
}

// ---------------------------------------------------------------------------
// Outline (Grouping) Properties
// ---------------------------------------------------------------------------

/** The worksheet's `<sheetPr><outlinePr>`: where grouped rows and columns keep their summary. */
export interface OutlinePropertiesData {
  /** Summary rows sit below their detail rows. Omitted means `true`. */
  summaryBelow?: boolean;
  /** Summary columns sit right of their detail columns. Omitted means `true`. */
  summaryRight?: boolean;
}

// ---------------------------------------------------------------------------
// Sparklines
// ---------------------------------------------------------------------------

/** Sparkline kind: a line, columns, or `'stacked'` (win/loss). */
export type SparklineType = 'line' | 'column' | 'stacked';

/** One `<x14:sparkline>`: the data it plots and the cell it is drawn in. */
export interface SparklineData {
  /** The data range (e.g. `'Sheet1!A1:A10'`). */
  formula: string;
  /** The cell the sparkline is drawn in (e.g. `'B1'`). */
  sqref: string;
}

/**
 * An `<x14:sparklineGroup>` from the worksheet's `<extLst>`: sparklines that share a type,
 * colours and options. Each `color*` string is hex RGB (RRGGBB) or ARGB (AARRGGBB), with or
 * without a leading #; written as uppercase ARGB (FF alpha added for RRGGBB); any other
 * non-empty string fails the write with `INVALID_WORKBOOK` and is reported by `validate()`
 * ('' writes no colour). Each
 * `color*Spec` holds the full colour when it has a theme, indexed, tint or auto attribute;
 * same write rule as {@link WorksheetData.tabColorSpec}: the spec is used while its `rgb`
 * equals the string, otherwise the string is written as plain `rgb`.
 */
export interface SparklineGroupData {
  /** Omitted means `'line'`. */
  sparklineType?: SparklineType;
  readonly sparklines: SparklineData[];
  colorSeries?: string | null;
  colorSeriesSpec?: ColorData | null;
  colorNegative?: string | null;
  colorNegativeSpec?: ColorData | null;
  colorAxis?: string | null;
  colorAxisSpec?: ColorData | null;
  colorMarkers?: string | null;
  colorMarkersSpec?: ColorData | null;
  colorFirst?: string | null;
  colorFirstSpec?: ColorData | null;
  colorLast?: string | null;
  colorLastSpec?: ColorData | null;
  colorHigh?: string | null;
  colorHighSpec?: ColorData | null;
  colorLow?: string | null;
  colorLowSpec?: ColorData | null;
  /** Line weight in points. */
  lineWeight?: number | null;
  markers?: boolean;
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  negative?: boolean;
  displayXAxis?: boolean;
  displayEmptyCellsAs?: 'gap' | 'zero' | 'span' | null;
  /** Vertical axis minimum; applies on a `'custom'` axis, which is written when `minAxisType` is unset. */
  manualMin?: number | null;
  /** Vertical axis maximum; applies on a `'custom'` axis, which is written when `maxAxisType` is unset. */
  manualMax?: number | null;
  rightToLeft?: boolean;
  /** How the axis minimum is chosen: per sparkline (`'individual'`, the default), shared by the group, or `manualMin`. */
  minAxisType?: 'individual' | 'group' | 'custom' | null;
  /** How the axis maximum is chosen: per sparkline (`'individual'`, the default), shared by the group, or `manualMax`. */
  maxAxisType?: 'individual' | 'group' | 'custom' | null;
  /** The horizontal axis is a date axis over `dateRange`. */
  dateAxis?: boolean;
  /** Data in hidden rows and columns is plotted. */
  displayHidden?: boolean;
  /** The range holding a date axis's dates (the group's `<xm:f>`), e.g. `'Sheet1!A20:J20'`. */
  dateRange?: string | null;
}

// ---------------------------------------------------------------------------
// WorksheetData
// ---------------------------------------------------------------------------

/**
 * A legacy cell comment (note) from `xl/comments{n}.xml`: plain text, with rich text
 * flattened on read. The writer writes the comments part and the VML note drawing Excel
 * shows notes through: it keeps a note drawing read from the file when its notes still
 * match, merges added, removed or moved notes into it, and generates one when there is
 * none. A note drawing it cannot merge (not UTF-8, not well-formed XML, or binding the `v`,
 * `o` or `x` prefix to another namespace) is written as it is, and one that several sheets
 * share is merged for the first of them with comments only.
 */
export interface CommentData {
  cellRef: string;
  author: string;
  text: string;
}

/** An @-mention inside a threaded comment's text. */
export interface MentionData {
  /** The `personId` of the mentioned person. */
  mentionPersonId: string;
  /** The mention's own ID. */
  mentionId: string;
  /** Start of the mention in the comment text. */
  startIndex: number;
  /** Length of the mention in the comment text. */
  length: number;
}

/** A threaded (modern) comment, `xl/threadedComments/threadedComment{n}.xml`. */
export interface ThreadedCommentData {
  /**
   * The comment's GUID in braces, `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`:
   * the form Excel requires (it writes upper case and reads either case) and
   * refuses the file without; `validate()` reports any other form as an error.
   */
  id: string;
  /** The cell the comment is on (e.g. `'A1'`). */
  refCell: string;
  /** The author's {@link PersonData.id} (a GUID in braces). */
  personId: string;
  text: string;
  /** ISO 8601 timestamp, written as `dT`. */
  timestamp: string;
  /**
   * The `id` of the thread root this comment replies to (a GUID in braces;
   * Excel drops a reply whose `parentId` has none).
   */
  parentId?: string;
  /** Whether the thread is resolved (`done`). */
  done?: boolean | null;
  /** The @-mentions in `text`. */
  mentions?: MentionData[];
}

/** A `<person>` of the persons list `xl/persons/person.xml`, the authors of threaded comments. */
export interface PersonData {
  /**
   * The person's GUID in braces, `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`, the
   * form Excel requires (as for {@link ThreadedCommentData.id}).
   */
  id: string;
  displayName: string;
  /** The person's `userId`. */
  userId?: string | null;
  providerId?: string;
}

/** The content of a worksheet part, `xl/worksheets/sheet{n}.xml`, with its tables, charts and comments. */
export interface WorksheetData {
  /** The `<dimension ref>` read from the file. The writer recomputes the used range from `rows`. */
  dimension: string | null;
  rows: RowData[];
  mergeCells: string[];
  autoFilter: AutoFilterData | null;
  frozenPane: FrozenPane | null;
  splitPane?: SplitPaneData | null;
  paneSelections?: PaneSelectionData[];
  sheetView?: SheetViewData | null;
  columns: ColumnInfo[];
  dataValidations?: DataValidationData[];
  conditionalFormatting?: ConditionalFormattingData[];
  hyperlinks?: HyperlinkData[];
  pageSetup?: PageSetupData | null;
  sheetProtection?: SheetProtectionData | null;
  comments?: CommentData[];
  /** Tab color as hex RGB (RRGGBB) or ARGB (AARRGGBB), with or without a leading #; written as uppercase ARGB and read back in that form; any other non-empty string fails the write with `INVALID_WORKBOOK` and is reported by `validate()`. */
  tabColor?: string | null;
  /**
   * The full tab colour when it has a theme, indexed, tint or auto attribute.
   * Used on write while its `rgb` equals `tabColor`; otherwise `tabColor` is
   * written as plain `rgb`.
   */
  tabColorSpec?: ColorData | null;
  tables?: TableDefinitionData[];
  headerFooter?: HeaderFooterData | null;
  pageBreaks?: PageBreaksData | null;
  outlineProperties?: OutlinePropertiesData | null;
  sparklineGroups?: SparklineGroupData[];
  charts?: WorksheetChartData[];
  pivotTables?: PivotTableData[];
  threadedComments?: ThreadedCommentData[];
  slicers?: SlicerData[];
  timelines?: TimelineData[];
  /**
   * Every worksheet `<extLst><ext>` other than the sparkline one, verbatim XML
   * kept for round-trip. Each entry is well-formed on its own and is written
   * inside `<extLst>` after the sparkline extension.
   */
  preservedExtensions?: string[];
}
