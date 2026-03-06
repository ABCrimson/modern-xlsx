// ---------------------------------------------------------------------------
// Read Options
// ---------------------------------------------------------------------------

/** Options for reading XLSX files. */
export interface ReadOptions {
  /** Password for encrypted XLSX files. */
  password?: string;
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
  /** Formula type: `'shared'`, `'array'`, or `null`. */
  formulaType?: string | null;
  /** The range reference for shared/array formulas. */
  formulaRef?: string | null;
  /** Shared formula group index. */
  sharedIndex?: number | null;
  /** Inline string content (used for `inlineStr` cell type). */
  inlineString?: string | null;
  /** Whether this is a dynamic array formula. */
  dynamicArray?: boolean | null;
  /** R1-style formula reference (for cross-sheet formulas). */
  formulaR1?: string | null;
  /** R2-style formula reference (for cross-sheet formulas). */
  formulaR2?: string | null;
  /** Two-dimensional data table flag. */
  formulaDt2d?: boolean | null;
  /** Data table row 1 flag. */
  formulaDtr1?: boolean | null;
  /** Data table row 2 flag. */
  formulaDtr2?: boolean | null;
  /** Rich text runs for mixed-format text within a cell. */
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
  /** Whether this row is hidden. */
  hidden: boolean;
  /** Outline (grouping) level (1-7), or `null` if ungrouped. */
  outlineLevel?: number | null;
  /** Whether this grouped row is collapsed. */
  collapsed?: boolean;
}

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
  /** Theme color ID for the grid color. */
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
}

/** Auto-filter dropdown configuration for a worksheet range. */
export interface AutoFilterData {
  /** The A1-style range covered by the auto-filter (e.g., `'A1:D10'`). */
  range: string;
  /** Per-column filter criteria. */
  filterColumns?: readonly FilterColumnData[];
}

export interface FilterColumnData {
  colId: number;
  filters?: readonly string[];
  customFilters?: CustomFiltersData | null;
}

export interface CustomFiltersData {
  andOp?: boolean;
  filters: readonly CustomFilterData[];
}

export interface CustomFilterData {
  operator?: string | null;
  val: string;
}

/** Hyperlink attached to a cell. */
export interface HyperlinkData {
  /** The cell reference this hyperlink is attached to (e.g., `'A1'`). */
  cellRef: string;
  /** The URL or internal reference target (e.g., `'https://example.com'` or `'Sheet2!A1'`). */
  location?: string | null;
  /** Display text shown in the cell. */
  display?: string | null;
  /** Tooltip text shown on hover. */
  tooltip?: string | null;
}

export interface PageMarginsData {
  top?: number | null;
  bottom?: number | null;
  left?: number | null;
  right?: number | null;
  header?: number | null;
  footer?: number | null;
}

export interface PageSetupData {
  paperSize?: number | null;
  orientation?: 'portrait' | 'landscape' | null;
  fitToWidth?: number | null;
  fitToHeight?: number | null;
  scale?: number | null;
  firstPageNumber?: number | null;
  horizontalDpi?: number | null;
  verticalDpi?: number | null;
  margins?: PageMarginsData | null;
}

export interface PageBreakData {
  id: number;
  min?: number | null;
  max?: number | null;
  man?: boolean;
}

export interface PageBreaksData {
  rowBreaks?: PageBreakData[];
  colBreaks?: PageBreakData[];
}

export interface SheetProtectionData {
  sheet: boolean;
  objects: boolean;
  scenarios: boolean;
  password?: string | null;
  formatCells: boolean;
  formatColumns: boolean;
  formatRows: boolean;
  insertColumns: boolean;
  insertRows: boolean;
  deleteColumns: boolean;
  deleteRows: boolean;
  sort: boolean;
  autoFilter: boolean;
}

export interface DataValidationData {
  sqref: string;
  validationType: string | null;
  operator: string | null;
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

export interface CfvoData {
  cfvoType: string;
  val?: string | null;
}

export interface ColorScaleData {
  readonly cfvos: CfvoData[];
  readonly colors: string[];
}

export interface DataBarData {
  readonly cfvos: CfvoData[];
  color: string;
}

export interface IconSetData {
  iconSetType?: string | null;
  readonly cfvos: CfvoData[];
}

export interface ConditionalFormattingRuleData {
  ruleType: string;
  priority: number;
  operator: string | null;
  formula: string | null;
  dxfId: number | null;
  colorScale?: ColorScaleData | null;
  dataBar?: DataBarData | null;
  iconSet?: IconSetData | null;
}

export interface ConditionalFormattingData {
  sqref: string;
  readonly rules: ConditionalFormattingRuleData[];
}

// ---------------------------------------------------------------------------
// Table definitions (Excel ListObjects)
// ---------------------------------------------------------------------------

export interface TableColumnData {
  id: number;
  name: string;
  totalsRowFunction?: string | null;
  totalsRowLabel?: string | null;
  calculatedColumnFormula?: string | null;
  headerRowDxfId?: number | null;
  dataDxfId?: number | null;
  totalsRowDxfId?: number | null;
}

export interface TableStyleInfoData {
  name?: string | null;
  showFirstColumn: boolean;
  showLastColumn: boolean;
  showRowStripes: boolean;
  showColumnStripes: boolean;
}

export interface TableDefinitionData {
  id: number;
  name?: string | null;
  displayName: string;
  ref: string;
  headerRowCount: number;
  totalsRowCount: number;
  totalsRowShown: boolean;
  readonly columns: TableColumnData[];
  styleInfo?: TableStyleInfoData | null;
  autoFilterRef?: string | null;
}

// ---------------------------------------------------------------------------
// Headers & Footers
// ---------------------------------------------------------------------------

export interface HeaderFooterData {
  oddHeader?: string | null;
  oddFooter?: string | null;
  evenHeader?: string | null;
  evenFooter?: string | null;
  firstHeader?: string | null;
  firstFooter?: string | null;
  differentOddEven?: boolean;
  differentFirst?: boolean;
  scaleWithDoc?: boolean;
  alignWithMargins?: boolean;
}

// ---------------------------------------------------------------------------
// Outline (Grouping) Properties
// ---------------------------------------------------------------------------

export interface OutlinePropertiesData {
  summaryBelow?: boolean;
  summaryRight?: boolean;
}

// ---------------------------------------------------------------------------
// Sparklines
// ---------------------------------------------------------------------------

export type SparklineType = 'line' | 'column' | 'stacked';

export interface SparklineData {
  formula: string;
  sqref: string;
}

export interface SparklineGroupData {
  sparklineType?: SparklineType;
  readonly sparklines: SparklineData[];
  colorSeries?: string | null;
  colorNegative?: string | null;
  colorAxis?: string | null;
  colorMarkers?: string | null;
  colorFirst?: string | null;
  colorLast?: string | null;
  colorHigh?: string | null;
  colorLow?: string | null;
  lineWeight?: number | null;
  markers?: boolean;
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  negative?: boolean;
  displayXAxis?: boolean;
  displayEmptyCellsAs?: 'gap' | 'zero' | 'span' | null;
  manualMin?: number | null;
  manualMax?: number | null;
  rightToLeft?: boolean;
}

// ---------------------------------------------------------------------------
// Chart definitions
// ---------------------------------------------------------------------------

export type ChartType =
  | 'bar'
  | 'column'
  | 'line'
  | 'pie'
  | 'doughnut'
  | 'scatter'
  | 'area'
  | 'radar'
  | 'bubble'
  | 'stock';
export type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';
export type ScatterStyle = 'lineMarker' | 'line' | 'marker' | 'smooth' | 'smoothMarker';
export type RadarStyle = 'standard' | 'marker' | 'filled';
export type MarkerStyleType =
  | 'circle'
  | 'square'
  | 'diamond'
  | 'triangle'
  | 'star'
  | 'x'
  | 'plus'
  | 'dash'
  | 'dot'
  | 'none';
export type TickLabelPosition = 'high' | 'low' | 'nextTo' | 'none';
export type TickMark = 'cross' | 'in' | 'out' | 'none';
export type AxisPosition = 'bottom' | 'top' | 'left' | 'right';
export type LegendPosition = 'top' | 'bottom' | 'left' | 'right' | 'topRight';

export type TrendlineType =
  | 'linear'
  | 'exponential'
  | 'logarithmic'
  | 'polynomial'
  | 'power'
  | 'movingAverage';
export type ErrorBarType = 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr' | 'custom';
export type ErrorBarDirection = 'both' | 'plus' | 'minus';

export interface TrendlineData {
  trendType: TrendlineType;
  order?: number | null;
  period?: number | null;
  forward?: number | null;
  backward?: number | null;
  displayEq?: boolean;
  displayRSqr?: boolean;
}

export interface ErrorBarsData {
  errType: ErrorBarType;
  direction?: ErrorBarDirection;
  value?: number | null;
}

export interface View3DData {
  rotX?: number | null;
  rotY?: number | null;
  perspective?: number | null;
  rAngAx?: boolean | null;
}

export interface ChartTitleData {
  text: string;
  overlay?: boolean;
  fontSize?: number | null;
  bold?: boolean | null;
  color?: string | null;
}

export interface ChartAxisData {
  id: number;
  crossAx: number;
  title?: ChartTitleData | null;
  numFmt?: string | null;
  sourceLinked?: boolean;
  min?: number | null;
  max?: number | null;
  majorUnit?: number | null;
  minorUnit?: number | null;
  logBase?: number | null;
  reversed?: boolean;
  tickLblPos?: TickLabelPosition | null;
  majorTickMark?: TickMark | null;
  minorTickMark?: TickMark | null;
  majorGridlines?: boolean;
  minorGridlines?: boolean;
  delete?: boolean;
  position?: AxisPosition | null;
  crossesAt?: number | null;
  /** Font size for tick labels in hundredths of a point (1400 = 14pt). */
  fontSize?: number | null;
}

export interface ChartLegendData {
  position: LegendPosition;
  overlay?: boolean;
}

export interface DataLabelsData {
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  numFmt?: string | null;
  showLeaderLines?: boolean;
}

export interface ChartSeriesData {
  idx: number;
  order: number;
  name?: string | null;
  catRef?: string | null;
  valRef: string;
  xValRef?: string | null;
  bubbleSizeRef?: string | null;
  fillColor?: string | null;
  lineColor?: string | null;
  lineWidth?: number | null;
  marker?: MarkerStyleType | null;
  smooth?: boolean | null;
  explosion?: number | null;
  dataLabels?: DataLabelsData | null;
  trendline?: TrendlineData | null;
  errorBars?: ErrorBarsData | null;
}

export interface ManualLayoutData {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ChartDataModel {
  chartType: ChartType;
  title?: ChartTitleData | null;
  readonly series: ChartSeriesData[];
  catAxis?: ChartAxisData | null;
  valAxis?: ChartAxisData | null;
  legend?: ChartLegendData | null;
  dataLabels?: DataLabelsData | null;
  grouping?: ChartGrouping | null;
  scatterStyle?: ScatterStyle | null;
  radarStyle?: RadarStyle | null;
  holeSize?: number | null;
  barDirHorizontal?: boolean | null;
  styleId?: number | null;
  plotAreaLayout?: ManualLayoutData | null;
  secondaryChart?: ChartDataModel | null;
  secondaryValAxis?: ChartAxisData | null;
  showDataTable?: boolean;
  view3d?: View3DData | null;
}

export interface ChartAnchorData {
  fromCol: number;
  fromRow: number;
  fromColOff?: number;
  fromRowOff?: number;
  toCol: number;
  toRow: number;
  toColOff?: number;
  toRowOff?: number;
  /** Width in EMUs (for oneCellAnchor charts). When set, the anchor is a oneCellAnchor. */
  extCx?: number | null;
  /** Height in EMUs (for oneCellAnchor charts). When set, the anchor is a oneCellAnchor. */
  extCy?: number | null;
}

export interface WorksheetChartData {
  chart: ChartDataModel;
  anchor: ChartAnchorData;
}

// ---------------------------------------------------------------------------
// Pivot Tables
// ---------------------------------------------------------------------------

export type PivotAxis = 'axisRow' | 'axisCol' | 'axisPage' | 'axisValues';
export type SubtotalFunction =
  | 'sum'
  | 'count'
  | 'average'
  | 'max'
  | 'min'
  | 'product'
  | 'countNums'
  | 'stdDev'
  | 'stdDevP'
  | 'var'
  | 'varP';

export interface PivotLocation {
  ref: string;
  firstHeaderRow?: number;
  firstDataRow?: number;
  firstDataCol?: number;
}

export interface PivotItem {
  t?: string;
  x?: number;
}

export interface PivotFieldData {
  axis?: PivotAxis;
  name?: string;
  items: PivotItem[];
  subtotals: SubtotalFunction[];
  compact: boolean;
  outline: boolean;
}

export interface PivotDataFieldData {
  name?: string;
  fld: number;
  subtotal: SubtotalFunction;
  numFmtId?: number;
}

export interface PivotPageFieldData {
  fld: number;
  item?: number;
  name?: string;
}

export interface PivotFieldRef {
  x: number;
}

export interface PivotTableData {
  name: string;
  dataCaption?: string;
  location: PivotLocation;
  pivotFields: PivotFieldData[];
  rowFields: PivotFieldRef[];
  colFields: PivotFieldRef[];
  dataFields: PivotDataFieldData[];
  pageFields: PivotPageFieldData[];
  cacheId: number;
}

// ---------------------------------------------------------------------------
// Pivot Cache
// ---------------------------------------------------------------------------

export type CacheValueType =
  | 'number'
  | 'string'
  | 'boolean'
  | 'dateTime'
  | 'missing'
  | 'error'
  | 'index';

export interface CacheValueNumber {
  type: 'number';
  v: number;
}

export interface CacheValueString {
  type: 'string';
  v: string;
}

export interface CacheValueBoolean {
  type: 'boolean';
  v: boolean;
}

export interface CacheValueDateTime {
  type: 'dateTime';
  v: string;
}

export interface CacheValueMissing {
  type: 'missing';
}

export interface CacheValueError {
  type: 'error';
  v: string;
}

export interface CacheValueIndex {
  type: 'index';
  v: number;
}

export type CacheValue =
  | CacheValueNumber
  | CacheValueString
  | CacheValueBoolean
  | CacheValueDateTime
  | CacheValueMissing
  | CacheValueError
  | CacheValueIndex;

export interface CacheSource {
  ref: string;
  sheet: string;
}

export interface CacheFieldData {
  name: string;
  sharedItems?: CacheValue[];
}

export interface PivotCacheDefinitionData {
  source: CacheSource;
  fields?: CacheFieldData[];
  recordCount?: number;
}

export interface PivotCacheRecordsData {
  records: CacheValue[][];
}

// ---------------------------------------------------------------------------
// Slicers
// ---------------------------------------------------------------------------

export type SortOrder = 'ascending' | 'descending';

export interface SlicerItem {
  n: string;
  s: boolean;
}

export interface SlicerData {
  name: string;
  caption?: string;
  cacheName: string;
  columnName?: string;
  sortOrder?: SortOrder;
  startItem?: number;
}

export interface SlicerCacheData {
  name: string;
  sourceName?: string;
  items: SlicerItem[];
}

// ---------------------------------------------------------------------------
// Timelines
// ---------------------------------------------------------------------------

export type TimelineLevel = 'years' | 'quarters' | 'months' | 'days';

export interface TimelineData {
  name: string;
  caption?: string;
  cacheName: string;
  sourceName?: string;
  level?: TimelineLevel;
}

export interface TimelineCacheData {
  name: string;
  sourceName?: string;
  selectionStart?: string;
  selectionEnd?: string;
}

// ---------------------------------------------------------------------------
// WorksheetData
// ---------------------------------------------------------------------------

export interface CommentData {
  cellRef: string;
  author: string;
  text: string;
}

export interface ThreadedCommentData {
  id: string;
  refCell: string;
  personId: string;
  text: string;
  timestamp: string;
  parentId?: string;
}

export interface PersonData {
  id: string;
  displayName: string;
  providerId?: string;
}

export interface WorksheetData {
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
  tabColor?: string | null;
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
}

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

/** A custom number format definition. */
export interface NumFmt {
  /** The numeric format ID (built-in: 0-163, custom: 164+). */
  id: number;
  /** The Excel format code string (e.g., `'#,##0.00'`, `'yyyy-mm-dd'`). */
  formatCode: string;
}

/** Cell text alignment properties. */
export interface AlignmentData {
  /** Horizontal alignment. */
  horizontal?:
    | 'general'
    | 'left'
    | 'center'
    | 'right'
    | 'fill'
    | 'justify'
    | 'centerContinuous'
    | 'distributed'
    | null;
  /** Vertical alignment. */
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed' | null;
  /** Whether text wraps within the cell. */
  wrapText?: boolean;
  /** Text rotation in degrees (0-180, or 255 for vertical text). */
  textRotation?: number | null;
  /** Text indent level. */
  indent?: number | null;
  /** Whether text shrinks to fit the cell width. */
  shrinkToFit?: boolean;
}

/** Cell protection properties (effective only when sheet protection is enabled). */
export interface ProtectionData {
  /** Whether the cell is locked (prevents editing when sheet is protected). */
  locked: boolean;
  /** Whether the cell formula is hidden from the user. */
  hidden: boolean;
}

/** Font properties for cell text rendering. */
export interface FontData {
  /** Font family name (e.g., `'Aptos'`, `'Arial'`). */
  name: string | null;
  /** Font size in points. */
  size: number | null;
  /** Whether the text is bold. */
  bold: boolean;
  /** Whether the text is italic. */
  italic: boolean;
  /** Whether the text is underlined. */
  underline: boolean;
  /** Whether the text has strikethrough. */
  strike: boolean;
  /** Font color as hex RGB (e.g., `'FF0000'` for red). */
  color: string | null;
  /** Vertical alignment: baseline, superscript, or subscript. */
  vertAlign?: 'baseline' | 'superscript' | 'subscript' | null;
  /** Font family number (1=Roman, 2=Swiss, 3=Modern, etc.). */
  family?: number | null;
  /** Character set identifier. */
  charset?: number | null;
  /** Font scheme (e.g., `'major'`, `'minor'`). */
  scheme?: string | null;
  /** Condense font spacing. */
  condense?: boolean;
  /** Extend font spacing. */
  extend?: boolean;
}

export interface GradientStopData {
  position: number;
  color: string;
}

export interface GradientFillData {
  gradientType?: string | null;
  degree?: number | null;
  stops?: readonly GradientStopData[];
}

/** Cell background fill properties. */
export interface FillData {
  /** Fill pattern type (e.g., `'none'`, `'solid'`, `'gray125'`). */
  patternType: string;
  /** Foreground color as hex RGB. */
  fgColor: string | null;
  /** Background color as hex RGB (used with pattern fills). */
  bgColor: string | null;
  /** Gradient fill configuration (alternative to pattern fill). */
  gradientFill?: GradientFillData | null;
}

/** A single border edge (left, right, top, or bottom). */
export interface BorderSideData {
  /** The border line style (e.g., `'thin'`, `'medium'`, `'double'`). */
  style: BorderStyle;
  /** Border color as hex RGB, or `null` for automatic. */
  color: string | null;
}

/** Cell border configuration for all four sides plus optional diagonal. */
export interface BorderData {
  /** Left border, or `null` for none. */
  left: BorderSideData | null;
  /** Right border, or `null` for none. */
  right: BorderSideData | null;
  /** Top border, or `null` for none. */
  top: BorderSideData | null;
  /** Bottom border, or `null` for none. */
  bottom: BorderSideData | null;
  /** Diagonal border, or `null` for none. */
  diagonal?: BorderSideData | null;
  /** Whether the diagonal goes from bottom-left to top-right. */
  diagonalUp?: boolean;
  /** Whether the diagonal goes from top-left to bottom-right. */
  diagonalDown?: boolean;
}

/** A cell format (XF) record referencing font, fill, border, and number format by index. */
export interface CellXfData {
  /** Index into `StylesData.numFmts` (0 = General). */
  numFmtId: number;
  /** Index into `StylesData.fonts`. */
  fontId: number;
  /** Index into `StylesData.fills`. */
  fillId: number;
  /** Index into `StylesData.borders`. */
  borderId: number;
  /** Text alignment configuration. */
  alignment?: AlignmentData | null;
  /** Cell protection configuration. */
  protection?: ProtectionData | null;
  /** Whether the font from this XF should be applied. */
  applyFont?: boolean;
  /** Whether the fill from this XF should be applied. */
  applyFill?: boolean;
  /** Whether the border from this XF should be applied. */
  applyBorder?: boolean;
  /** Whether the number format from this XF should be applied. */
  applyNumberFormat?: boolean;
  /** Whether the alignment from this XF should be applied. */
  applyAlignment?: boolean;
  /** Whether the protection from this XF should be applied. */
  applyProtection?: boolean;
}

export interface DxfStyleData {
  font?: FontData | null;
  fill?: FillData | null;
  border?: BorderData | null;
  numFmt?: NumFmt | null;
}

export interface CellStyleData {
  name: string;
  xfId: number;
  builtinId?: number | null;
}

/** The shared style tables for the entire workbook. */
export interface StylesData {
  /** Custom number format definitions (IDs 164+). */
  numFmts: NumFmt[];
  /** Font definitions referenced by CellXfData.fontId. */
  fonts: FontData[];
  /** Fill definitions referenced by CellXfData.fillId. */
  fills: FillData[];
  /** Border definitions referenced by CellXfData.borderId. */
  borders: BorderData[];
  /** Cell format records -- each cell's `styleIndex` indexes into this array. */
  cellXfs: CellXfData[];
  /** Differential formatting styles (used by conditional formatting and tables). */
  dxfs?: DxfStyleData[];
  /** Named cell styles (e.g., `'Normal'`, `'Heading 1'`). */
  cellStyles?: CellStyleData[];
}

// ---------------------------------------------------------------------------
// Workbook Protection
// ---------------------------------------------------------------------------

export interface WorkbookProtectionData {
  lockStructure?: boolean;
  lockWindows?: boolean;
  lockRevision?: boolean;
  workbookAlgorithmName?: string;
  workbookHashValue?: string;
  workbookSaltValue?: string;
  workbookSpinCount?: number;
  revisionsAlgorithmName?: string;
  revisionsHashValue?: string;
  revisionsSaltValue?: string;
  revisionsSpinCount?: number;
  workbookPassword?: string;
  revisionsPassword?: string;
}

// ---------------------------------------------------------------------------
// Sheet & Workbook
// ---------------------------------------------------------------------------

export interface SheetData {
  name: string;
  /** Sheet visibility: 'visible' (default/omitted), 'hidden', or 'veryHidden'. */
  state?: SheetState | null;
  worksheet: WorksheetData;
}

/** A named range or defined name in the workbook. */
export interface DefinedNameData {
  /** The defined name (e.g., `'Revenue'`, `'_xlnm.Print_Area'`). */
  name: string;
  /** The range reference or formula value (e.g., `'Sheet1!$A$1:$A$10'`). */
  value: string;
  /** Zero-based sheet index for sheet-scoped names, or `null` for workbook-scoped. */
  sheetId: number | null;
}

/** A single text segment with optional formatting within a rich text cell. */
export interface RichTextRun {
  /** The text content for this segment. */
  text: string;
  /** Whether this segment is bold. */
  bold?: boolean;
  /** Whether this segment is italic. */
  italic?: boolean;
  /** Whether this segment is underlined. */
  underline?: boolean;
  /** Whether this segment has strikethrough. */
  strike?: boolean;
  /** Font family name for this segment. */
  fontName?: string;
  /** Font size in points for this segment. */
  fontSize?: number;
  /** Text color as hex RGB for this segment (e.g., `'FF0000'`). */
  color?: string;
}

export interface SharedStringsData {
  readonly strings: string[];
  readonly richRuns?: (RichTextRun[] | null)[];
}

// ---------------------------------------------------------------------------
// Document properties, theme, calc chain, workbook views
// ---------------------------------------------------------------------------

export interface DocPropertiesData {
  title?: string | null;
  subject?: string | null;
  creator?: string | null;
  keywords?: string | null;
  description?: string | null;
  lastModifiedBy?: string | null;
  created?: string | null;
  modified?: string | null;
  category?: string | null;
  contentStatus?: string | null;
  application?: string | null;
  company?: string | null;
  manager?: string | null;
  appVersion?: string | null;
  hyperlinkBase?: string | null;
  revision?: string | null;
}

export interface ThemeColorsData {
  dk1: string;
  lt1: string;
  dk2: string;
  lt2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  hlink: string;
  folHlink: string;
}

export interface CalcChainEntryData {
  cellRef: string;
  sheetId: number;
}

export interface WorkbookViewData {
  activeTab: number;
  firstSheet: number;
  showHorizontalScroll: boolean;
  showVerticalScroll: boolean;
  showSheetTabs: boolean;
  windowWidth?: number | null;
  windowHeight?: number | null;
  tabRatio?: number | null;
}

export interface WorkbookData {
  sheets: SheetData[];
  dateSystem: DateSystem;
  styles: StylesData;
  definedNames?: DefinedNameData[];
  sharedStrings?: SharedStringsData;
  docProperties?: DocPropertiesData | null;
  themeColors?: ThemeColorsData | null;
  calcChain?: readonly CalcChainEntryData[];
  workbookViews?: WorkbookViewData[];
  protection?: WorkbookProtectionData | null;
  /** Persons list for threaded comments (workbook-level). */
  persons?: PersonData[];
  /** Pivot cache definitions (workbook-level). */
  pivotCaches?: PivotCacheDefinitionData[];
  /** Pivot cache records (workbook-level, parallel to pivotCaches). */
  pivotCacheRecords?: PivotCacheRecordsData[];
  /** Slicer cache definitions (workbook-level). */
  slicerCaches?: SlicerCacheData[];
  /** Timeline cache definitions (workbook-level). */
  timelineCaches?: TimelineCacheData[];
  /** Opaque ZIP entries preserved through roundtrip (drawings, media, charts, etc.) */
  preservedEntries?: Record<string, number[]>;
}

// ---------------------------------------------------------------------------
// Validation & Compliance
// ---------------------------------------------------------------------------

export type Severity = 'info' | 'warning' | 'error';

export type IssueCategory =
  | 'cellReference'
  | 'styleIndex'
  | 'mergeCell'
  | 'sharedString'
  | 'sheetName'
  | 'definedName'
  | 'dataValidation'
  | 'conditionalFormatting'
  | 'theme'
  | 'metadata'
  | 'structure';

export interface ValidationIssue {
  severity: Severity;
  category: IssueCategory;
  message: string;
  location: string;
  suggestion: string;
  autoFixable: boolean;
}

export interface ValidationReport {
  readonly issues: ValidationIssue[];
  errorCount: number;
  warningCount: number;
  infoCount: number;
  isValid: boolean;
}

export interface RepairResult {
  workbook: WorkbookData;
  report: ValidationReport;
  repairCount: number;
}
