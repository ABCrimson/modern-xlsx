export type DateSystem = 'date1900' | 'date1904';

export type CellType = 'sharedString' | 'number' | 'boolean' | 'error' | 'formulaStr' | 'inlineStr';

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

export interface CellData {
  reference: string;
  cellType: CellType;
  styleIndex: number | null;
  value: string | null;
  formula: string | null;
  formulaType?: string | null;
  formulaRef?: string | null;
  sharedIndex?: number | null;
  inlineString?: string | null;
}

export interface RowData {
  index: number;
  cells: CellData[];
  height: number | null;
  hidden: boolean;
  outlineLevel?: number | null;
}

// ---------------------------------------------------------------------------
// Worksheet structures
// ---------------------------------------------------------------------------

export interface FrozenPane {
  rows: number;
  cols: number;
}

export interface ColumnInfo {
  min: number;
  max: number;
  width: number;
  hidden: boolean;
  customWidth: boolean;
  outlineLevel?: number | null;
}

export interface AutoFilterData {
  range: string;
  filterColumns?: FilterColumnData[];
}

export interface FilterColumnData {
  colId: number;
  filters: string[];
}

export interface HyperlinkData {
  cellRef: string;
  location?: string | null;
  display?: string | null;
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
  cfvos: CfvoData[];
  colors: string[];
}

export interface DataBarData {
  cfvos: CfvoData[];
  color: string;
}

export interface IconSetData {
  iconSetType?: string | null;
  cfvos: CfvoData[];
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
  rules: ConditionalFormattingRuleData[];
}

// ---------------------------------------------------------------------------
// WorksheetData
// ---------------------------------------------------------------------------

export interface CommentData {
  cellRef: string;
  author: string;
  text: string;
}

export interface WorksheetData {
  dimension: string | null;
  rows: RowData[];
  mergeCells: string[];
  autoFilter: AutoFilterData | null;
  frozenPane: FrozenPane | null;
  columns: ColumnInfo[];
  dataValidations?: DataValidationData[];
  conditionalFormatting?: ConditionalFormattingData[];
  hyperlinks?: HyperlinkData[];
  pageSetup?: PageSetupData | null;
  sheetProtection?: SheetProtectionData | null;
  comments?: CommentData[];
}

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

export interface NumFmt {
  id: number;
  formatCode: string;
}

export interface AlignmentData {
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
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed' | null;
  wrapText?: boolean;
  textRotation?: number | null;
  indent?: number | null;
  shrinkToFit?: boolean;
}

export interface ProtectionData {
  locked: boolean;
  hidden: boolean;
}

export interface FontData {
  name: string | null;
  size: number | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color: string | null;
  vertAlign?: 'baseline' | 'superscript' | 'subscript' | null;
  family?: number | null;
  charset?: number | null;
  scheme?: string | null;
  condense?: boolean;
  extend?: boolean;
}

export interface GradientStopData {
  position: number;
  color: string;
}

export interface GradientFillData {
  gradientType?: string | null;
  degree?: number | null;
  stops?: GradientStopData[];
}

export interface FillData {
  patternType: string;
  fgColor: string | null;
  bgColor: string | null;
  gradientFill?: GradientFillData | null;
}

export interface BorderSideData {
  style: BorderStyle;
  color: string | null;
}

export interface BorderData {
  left: BorderSideData | null;
  right: BorderSideData | null;
  top: BorderSideData | null;
  bottom: BorderSideData | null;
  diagonal?: BorderSideData | null;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
}

export interface CellXfData {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  alignment?: AlignmentData | null;
  protection?: ProtectionData | null;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyNumberFormat?: boolean;
  applyAlignment?: boolean;
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

export interface StylesData {
  numFmts: NumFmt[];
  fonts: FontData[];
  fills: FillData[];
  borders: BorderData[];
  cellXfs: CellXfData[];
  dxfs?: DxfStyleData[];
  cellStyles?: CellStyleData[];
}

// ---------------------------------------------------------------------------
// Sheet & Workbook
// ---------------------------------------------------------------------------

export interface SheetData {
  name: string;
  worksheet: WorksheetData;
}

export interface DefinedNameData {
  name: string;
  value: string;
  sheetId: number | null;
}

export interface RichTextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  fontName?: string;
  fontSize?: number;
  color?: string;
}

export interface SharedStringsData {
  strings: string[];
  richRuns?: (RichTextRun[] | null)[];
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
  calcChain?: CalcChainEntryData[];
  workbookViews?: WorkbookViewData[];
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
  issues: ValidationIssue[];
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
