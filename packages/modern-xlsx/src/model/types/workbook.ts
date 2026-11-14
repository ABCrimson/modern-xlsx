import type { ColorData, DateSystem } from './cells.js';
import type {
  PivotCacheDefinitionData,
  PivotCacheRecordsData,
  SlicerCacheData,
  TimelineCacheData,
} from './pivot.js';
import type { StylesData } from './styles.js';
import type { PersonData, SheetState, WorksheetData } from './worksheet.js';

// ---------------------------------------------------------------------------
// Workbook Protection
// ---------------------------------------------------------------------------

/**
 * Workbook protection, `<workbookProtection>`. The password fields hold the hashes a file
 * carries (the iterated `*HashValue` group or the legacy 16-bit `*Password` hash) and are
 * written back unchanged.
 */
export interface WorkbookProtectionData {
  /** Locks the workbook structure: adding, deleting, renaming, moving and copying sheets. */
  lockStructure?: boolean;
  /** Locks the window position and size. */
  lockWindows?: boolean;
  /** Locks the revision log. */
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

/** A sheet of the workbook: its name, visibility and worksheet content. */
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
  /** `comment` attribute. */
  comment?: string | null;
  /** `customMenu` attribute. */
  customMenu?: string | null;
  /** `description` attribute. */
  description?: string | null;
  /** `help` attribute. */
  help?: string | null;
  /** `statusBar` attribute. */
  statusBar?: string | null;
  /** Whether the name is hidden in the Name Manager. */
  hidden?: boolean;
  /** Whether the name is a function (macro) name. */
  function?: boolean;
  /** Whether the name is a Visual Basic procedure. */
  vbProcedure?: boolean;
  /** Whether the name is an Excel 4.0 macro name. */
  xlm?: boolean;
  /** `functionGroupId` attribute. */
  functionGroupId?: number | null;
  /** `shortcutKey` attribute. */
  shortcutKey?: string | null;
  /** Whether the name is published to the server. */
  publishToServer?: boolean;
  /** Whether the name is a workbook parameter. */
  workbookParameter?: boolean;
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
  /** Text color for this segment as hex RGB (RRGGBB) or ARGB (AARRGGBB), with or without a leading #; written as uppercase ARGB (FF alpha added for RRGGBB) and read back in that form (e.g., `'FF0000'`). */
  color?: string;
  /** The full run colour; see {@link ColorData} for the precedence rule against `color`. */
  colorSpec?: ColorData;
  /** Vertical alignment: `'baseline'`, `'superscript'` or `'subscript'`. */
  vertAlign?: string;
  /** Font family number. */
  family?: number;
  /** Character set identifier. */
  charset?: number;
  /** Font scheme: `'major'`, `'minor'` or `'none'`. */
  scheme?: string;
  /** The underline type when it is not single: `'double'`, `'singleAccounting'` or `'doubleAccounting'`. */
  underlineStyle?: string;
}

/**
 * The shared string table (`xl/sharedStrings.xml`) as read. The writer ignores it and
 * builds the table from the cell values.
 */
export interface SharedStringsData {
  /** The plain text of every entry; a rich-text entry's runs joined, phonetic runs left out. */
  readonly strings: string[];
  /** Parallel to `strings`: the runs of a rich-text entry, `null` for a plain one. */
  readonly richRuns?: (RichTextRun[] | null)[];
}

// ---------------------------------------------------------------------------
// Document properties, theme, calc chain, workbook views
// ---------------------------------------------------------------------------

/**
 * Document properties from `docProps/core.xml` and `docProps/app.xml`. Each value is the
 * element's raw text, white space included; properties not listed here are dropped.
 */
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
  lastPrinted?: string | null;
  language?: string | null;
  identifier?: string | null;
  version?: string | null;
  totalTime?: string | null;
  docSecurity?: string | null;
}

/**
 * The 12 colours of the theme's `<a:clrScheme>` (`xl/theme/theme1.xml`), each six hex
 * digits RRGGBB without `#` (e.g. `'FFFFFF'`); a system colour gives its `lastClr`.
 */
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

/** A `<c>` entry of the calculation chain (`xl/calcChain.xml`): a formula cell in calculation order. */
export interface CalcChainEntryData {
  cellRef: string;
  /** The 1-based position of the cell's sheet in `WorkbookData.sheets`. */
  sheetId: number;
  /** `s`: the cell is on a child chain. */
  s?: boolean;
  /** `l`: a new dependency level starts here. */
  l?: boolean;
  /** `t`: the cell is on a new thread. */
  t?: boolean;
  /** `a`: the cell is part of an array formula. */
  a?: boolean;
}

/** `<calcPr>` calculation settings from workbook.xml; every attribute is optional. */
export interface CalcPrData {
  calcId?: number | null;
  calcMode?: string | null;
  fullCalcOnLoad?: boolean | null;
  refMode?: string | null;
  iterate?: boolean | null;
  iterateCount?: number | null;
  iterateDelta?: number | null;
  fullPrecision?: boolean | null;
  calcCompleted?: boolean | null;
  calcOnSave?: boolean | null;
  concurrentCalc?: boolean | null;
  concurrentManualCount?: number | null;
  forceFullCalc?: boolean | null;
}

/** `<fileVersion>` application version from workbook.xml. */
export interface FileVersionData {
  appName?: string | null;
  lastEdited?: string | null;
  lowestEdited?: string | null;
  rupBuild?: string | null;
  codeName?: string | null;
}

/** `<workbookPr>` attributes other than `date1904` (see `WorkbookData.dateSystem`). */
export interface WorkbookPrData {
  showObjects?: string | null;
  showBorderUnselectedTables?: boolean | null;
  filterPrivacy?: boolean | null;
  promptedSolutions?: boolean | null;
  showInkAnnotation?: boolean | null;
  backupFile?: boolean | null;
  saveExternalLinkValues?: boolean | null;
  updateLinks?: string | null;
  codeName?: string | null;
  hidePivotFieldList?: boolean | null;
  showPivotChartFilter?: boolean | null;
  allowRefreshQuery?: boolean | null;
  publishItems?: boolean | null;
  checkCompatibility?: boolean | null;
  autoCompressPictures?: boolean | null;
  refreshAllConnections?: boolean | null;
  defaultThemeVersion?: number | null;
}

/** A `<workbookView>` of `<bookViews>`: the window, its scroll bars and sheet tabs. */
export interface WorkbookViewData {
  /** 0-based position in `WorkbookData.sheets` of the active sheet. */
  activeTab: number;
  /** 0-based position in `WorkbookData.sheets` of the first visible sheet tab. */
  firstSheet: number;
  showHorizontalScroll: boolean;
  showVerticalScroll: boolean;
  showSheetTabs: boolean;
  windowWidth?: number | null;
  windowHeight?: number | null;
  tabRatio?: number | null;
}

/**
 * The whole workbook data model: what the reader produces and the writer consumes, and
 * what a `Workbook` wraps.
 */
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
  /** Calculation settings from `<calcPr>`. */
  calcPr?: CalcPrData | null;
  /** Application version from `<fileVersion>`. */
  fileVersion?: FileVersionData | null;
  /** `<workbookPr>` attributes other than `date1904`. */
  workbookPr?: WorkbookPrData | null;
  /**
   * ZIP paths of the external link parts (`xl/externalLinks/externalLink{n}.xml`,
   * kept in `preservedEntries`) in workbook.xml `<externalReferences>` order,
   * which formulas use as `[n]`.
   */
  externalReferences?: string[];
  /** Persons list for threaded comments (workbook-level). */
  persons?: PersonData[];
  /**
   * Pivot cache definitions (workbook-level). A read lists the caches that have records first
   * and the caches saved without their data after them, each group in the order of the
   * workbook's relationships, so that `pivotCacheRecords[i]` belongs to `pivotCaches[i]`.
   * Find a pivot table's cache by `cacheId`, not by position.
   */
  pivotCaches?: PivotCacheDefinitionData[];
  /**
   * Pivot cache records (workbook-level): `pivotCacheRecords[i]` belongs to `pivotCaches[i]`,
   * so a read lists the caches with records first.
   */
  pivotCacheRecords?: PivotCacheRecordsData[];
  /** Slicer cache definitions (workbook-level). */
  slicerCaches?: SlicerCacheData[];
  /** Timeline cache definitions (workbook-level). */
  timelineCaches?: TimelineCacheData[];
  /**
   * ZIP entries the reader keeps as raw bytes rather than in the model: parts it does not
   * model (media, VML, custom XML, external links, printer settings, VBA and the like),
   * worksheet `.rels`, drawings that are not chart-only, a theme other than the one the
   * writer generates, and a comment, table, drawing, chart, pivot, threaded-comment, person,
   * slicer or timeline part that fails to parse. The writer copies each entry unchanged
   * except where it writes a part at the same path.
   */
  preservedEntries?: Record<string, number[]>;
  /**
   * The content type the source `[Content_Types].xml` declared for each `preservedEntries`
   * path but a `.rels`: its Override, else the Default for its extension. On write each
   * recorded type is declared again (as an Override, unless the Default for the extension
   * already gives it), so a part the writer has no rule for (`chartEx`, `ctrlProps`, ...)
   * keeps its type. Paths the source declared no type for are absent.
   */
  preservedContentTypes?: Record<string, string>;
}

// ---------------------------------------------------------------------------
// Validation & Compliance
// ---------------------------------------------------------------------------

/**
 * Severity of a validation issue: `'error'` for a broken name, reference or rule,
 * `'warning'` for a problem a reader can usually work around, `'info'` for a finding that
 * needs no action.
 */
export type Severity = 'info' | 'warning' | 'error';

/** The part of the workbook a validation issue concerns. */
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
  | 'structure'
  | 'sheetProtection';

/** One finding of `Workbook.validate()`, with a suggested fix. */
export interface ValidationIssue {
  severity: Severity;
  category: IssueCategory;
  message: string;
  /** Where the issue is, such as `'Sheet1!A1'` or `'styles.cellXfs[3]'`. */
  location: string;
  suggestion: string;
  /** Whether `Workbook.repair()` fixes this issue. */
  autoFixable: boolean;
}

/**
 * The result of `Workbook.validate()`: the issues found and the count per severity.
 *
 * The report is bounded: `issues` holds the first issues found, at most 10,000 of them and
 * at most 16 MiB of `message` and `location` text (a sheet or table name can repeat in
 * every issue). The counts and `isValid` cover every issue found, stored or not, and
 * `omittedIssueCount` says how many were left out.
 */
export interface ValidationReport {
  readonly issues: ValidationIssue[];
  errorCount: number;
  warningCount: number;
  infoCount: number;
  /** `true` when there are no `'error'` issues. */
  isValid: boolean;
  /**
   * The issues found but left out of `issues` because the report reached its bound, so
   * `issues.length + (omittedIssueCount ?? 0)` is `errorCount + warningCount + infoCount`.
   * Absent when every issue is stored.
   */
  omittedIssueCount?: number;
}

/**
 * The repair result from the WASM layer: the repaired workbook data, the validation report
 * of the repaired workbook, and the number of repairs applied.
 */
export interface RepairResult {
  workbook: WorkbookData;
  report: ValidationReport;
  repairCount: number;
}
