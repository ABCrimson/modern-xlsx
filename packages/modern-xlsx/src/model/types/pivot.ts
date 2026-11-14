import type { CellXfData } from './styles.js';

// ---------------------------------------------------------------------------
// Pivot Tables
// ---------------------------------------------------------------------------

/** The `axis` of a `<pivotField>`: the row, column, page (filter) or values area. */
export type PivotAxis = 'axisRow' | 'axisCol' | 'axisPage' | 'axisValues';
/**
 * An aggregation function: the `subtotal` of a `<dataField>`, or one of the
 * `*Subtotal` flags of a `<pivotField>`.
 */
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

/**
 * A pivot table's `<location>`: its range on the sheet and the offsets of its header and
 * data. An unset offset is written with the value Excel gives a table of this shape:
 * `firstHeaderRow` 1; `firstDataRow` 2 with column fields, else 1; `firstDataCol` 1 with
 * row fields, else 0.
 */
export interface PivotLocation {
  ref: string;
  firstHeaderRow?: number;
  firstDataRow?: number;
  firstDataCol?: number;
}

/** An `<item>` of a `<pivotField>`, with its `t` and `x` attributes. */
export interface PivotItem {
  t?: string;
  x?: number;
}

/**
 * A `<pivotField>`: a field of the pivot table and the area (`axis`) it is placed in.
 * Its `showAll`, `sortType` and `numFmtId` attributes are dropped on read.
 */
export interface PivotFieldData {
  axis?: PivotAxis;
  /** `true` when the field is used as a data (values) field, with or without an axis. */
  dataField?: boolean;
  name?: string;
  /**
   * The field's `<item>`s. Excel saves every row, column and page field with items; an
   * empty list on such a field marks a field built in code, and the writer then writes the
   * items Excel needs to refresh the table: one `{ x }` per shared item of the cache field at
   * the same index, then the subtotal items (`{ t: 'default' }`, or one per function in
   * `subtotals`, none when `defaultSubtotal` is `false`). The model is not changed. A
   * non-empty list is written as given. The write throws (SIZE_LIMIT) when a field would
   * get more than 1,048,576 such items, or the workbook's tables more than 4,194,304 items
   * and empty fields in all; explicit items are not counted.
   */
  items: PivotItem[];
  /** The functions whose `*Subtotal="1"` flag is set. */
  subtotals: SubtotalFunction[];
  compact: boolean;
  outline: boolean;
  /**
   * `defaultSubtotal`. `false` shows no subtotal for the field; when unset
   * Excel's automatic subtotal applies, also when `subtotals` is empty.
   */
  defaultSubtotal?: boolean | null;
}

/** A `<dataField>`: a field summarised in the pivot table's values area. */
export interface PivotDataFieldData {
  name?: string;
  fld: number;
  subtotal: SubtotalFunction;
  /** Number-format ID of the values; see {@link CellXfData.numFmtId}. */
  numFmtId?: number;
}

/** A `<pageField>`: a field placed in the pivot table's page (filter) area. */
export interface PivotPageFieldData {
  fld: number;
  item?: number;
  name?: string;
}

/**
 * A `<field>` of `<rowFields>` or `<colFields>`, naming a pivot field by its `x` index.
 * `x: -2` is the values field, which holds the data fields; with two or more data fields
 * and no `-2` in either list, the writer appends it to `<colFields>` (Excel refuses the
 * table without it).
 */
export interface PivotFieldRef {
  x: number;
}

/**
 * A pivot table, `xl/pivotTables/pivotTable{n}.xml`. The `rowItems` and `colItems`
 * layout caches and `pivotTableStyleInfo` are dropped on read.
 *
 * The writer does not compute the table's layout or its cell values. For a table built
 * in code (a row, column or page field with empty `items`), it completes the items (see
 * {@link PivotFieldData.items}) and writes `refreshOnLoad="1"` on the cache the table
 * uses, so Excel lays the table out and shows its values when it opens the file (and
 * marks the workbook changed). Read back, the table has those items, so its cache is
 * written without `refreshOnLoad`: the table shows the values in its cells, which Excel
 * saves with the file, and a file no Excel has saved shows them after a refresh.
 *
 * On write, `pivotFields` gets empty fields up to the number of fields in the cache that
 * `cacheId` names, because Excel refuses a table with fewer; the model is not changed.
 */
export interface PivotTableData {
  name: string;
  /** `dataCaption`; unset is written as `'Values'`. */
  dataCaption?: string;
  location: PivotLocation;
  pivotFields: PivotFieldData[];
  rowFields: PivotFieldRef[];
  colFields: PivotFieldRef[];
  dataFields: PivotDataFieldData[];
  pageFields: PivotPageFieldData[];
  /**
   * The cache this table summarises: a {@link PivotCacheDefinitionData.cacheId}, or the
   * cache's 0-based index in `WorkbookData.pivotCaches` when that is unset.
   */
  cacheId: number;
}

// ---------------------------------------------------------------------------
// Pivot Cache
// ---------------------------------------------------------------------------

/** The `type` tag of a {@link CacheValue}. */
export type CacheValueType =
  | 'number'
  | 'string'
  | 'boolean'
  | 'dateTime'
  | 'missing'
  | 'error'
  | 'index';

/** A numeric cache value, `<n v>`. */
export interface CacheValueNumber {
  type: 'number';
  v: number;
}

/** A text cache value, `<s v>`. */
export interface CacheValueString {
  type: 'string';
  v: string;
}

/** A boolean cache value, `<b v>`. */
export interface CacheValueBoolean {
  type: 'boolean';
  v: boolean;
}

/** A date-time cache value, `<d v>`, kept as the attribute's text. */
export interface CacheValueDateTime {
  type: 'dateTime';
  v: string;
}

/** An empty cache value, `<m/>`. */
export interface CacheValueMissing {
  type: 'missing';
}

/** An error cache value, `<e v>`, such as `'#N/A'`. */
export interface CacheValueError {
  type: 'error';
  v: string;
}

/** A cache record value that refers to the field's shared item at index `v`, `<x v>`. */
export interface CacheValueIndex {
  type: 'index';
  v: number;
}

/** A value in a pivot cache field's shared items or in a cache record, tagged by `type`. */
export type CacheValue =
  | CacheValueNumber
  | CacheValueString
  | CacheValueBoolean
  | CacheValueDateTime
  | CacheValueMissing
  | CacheValueError
  | CacheValueIndex;

/**
 * The source of a pivot cache: the attributes of its `<worksheetSource>`, or, for a source
 * that is not a worksheet, the whole `<cacheSource>` element in `rawXml`. The writer emits
 * each `<worksheetSource>` attribute only when it has a value, never as `""`. An `r:id` in
 * the source (Excel writes one for a source in another workbook) names one of
 * {@link PivotCacheDefinitionData.relationships}; the writer drops one that names no kept
 * relationship, because Excel refuses such a file.
 */
export interface CacheSource {
  /** The source range, such as `A1:C5`; `''` when the file has none (a source given by `name`). */
  ref: string;
  /** The sheet of `ref`, or of a sheet-scoped `name`; `''` when the file has none. */
  sheet: string;
  /**
   * A table or defined name. Excel writes it, with no `ref` or `sheet`, for a pivot table
   * made from a table or a named range (`<worksheetSource name="Sales"/>`).
   */
  name?: string;
  /** `r:id`: the relationship naming the workbook that holds the source, for a source in another workbook. */
  rId?: string;
  /**
   * A source that is not a worksheet (`<cacheSource>` `type` `consolidation`, `external` or
   * `scenario`; Excel writes `consolidation` for a pivot table over multiple consolidation
   * ranges): the whole `<cacheSource>` element as read. Opaque: it is written as read in
   * place of the fields above, except that an `r:id` naming no kept relationship is dropped;
   * delete it to write those instead.
   */
  rawXml?: string;
}

/** Attributes of a cache field's `<sharedItems>` element, kept so the type flags survive a round trip. */
export interface SharedItemsAttrs {
  containsSemiMixedTypes?: boolean | null;
  containsNonDate?: boolean | null;
  containsDate?: boolean | null;
  containsString?: boolean | null;
  containsBlank?: boolean | null;
  containsMixedTypes?: boolean | null;
  containsNumber?: boolean | null;
  containsInteger?: boolean | null;
  minValue?: number | null;
  maxValue?: number | null;
  minDate?: string | null;
  maxDate?: string | null;
  longText?: boolean | null;
}

/** A `<cacheField>` of a pivot cache: a source column's name and shared items. */
export interface CacheFieldData {
  name: string;
  /** Number-format ID of the cache field; see {@link CellXfData.numFmtId}. */
  numFmtId?: number | null;
  sharedItems?: CacheValue[];
  /** The `<sharedItems>` attributes. */
  sharedItemsAttrs?: SharedItemsAttrs | null;
}

/**
 * A pivot cache definition, `xl/pivotCache/pivotCacheDefinition{n}.xml`. Its
 * `refreshOnLoad` attribute is not modelled: the reader drops it, and the writer sets it
 * on a cache that a pivot table built in code uses (see {@link PivotTableData}).
 */
export interface PivotCacheDefinitionData {
  /** The `cacheId` workbook.xml gives this cache; pivot tables refer to it through `PivotTableData.cacheId`. Unset writes the cache's 0-based index; duplicate ids make the write throw (INVALID_WORKBOOK). */
  cacheId?: number | null;
  source: CacheSource;
  fields?: CacheFieldData[];
  recordCount?: number;
  /**
   * The relationships of the definition's `.rels` other than the one to its records, such as
   * the `externalLinkPath` Excel writes for a source in another workbook. Opaque: written back
   * with their ids; the records relationship takes an id none of them uses. Absent when there
   * are none. Skipped with a warning: one of the `pivotCacheRecords` type, one whose `id` is
   * empty, does not start with an ASCII letter or `_`, or holds a character other than an ASCII
   * letter, digit, `_`, `-` or `.` (a source `r:id` naming such an `id` is dropped), and a
   * second one with the same `id` (the first is kept).
   */
  relationships?: {
    /** `Id`, which the source names by `r:id`. */
    id: string;
    /** `Type`, the relationship type URI. */
    type: string;
    /** `Target`: a file path or URI when external, else a part name relative to the definition. */
    target: string;
    /**
     * `TargetMode`: `'External'`, or absent for a part inside the package. The writer also
     * reads it in another ASCII letter case (JSON from elsewhere), and writes `External`.
     */
    targetMode?: 'External';
  }[];
}

/** Pivot cache records, `xl/pivotCache/pivotCacheRecords{n}.xml`: one value list per record. */
export interface PivotCacheRecordsData {
  records: CacheValue[][];
}

// ---------------------------------------------------------------------------
// Slicers
// ---------------------------------------------------------------------------

/** Sort order of slicer items. */
export type SortOrder = 'ascending' | 'descending';

/** An item of a slicer cache, with its selection state. */
export interface SlicerItem {
  /** Item name. */
  n: string;
  /** Whether the item is selected. */
  s: boolean;
  /** Index of the item in the pivot field's shared items, when the item carries `x`. */
  x?: number | null;
  /** The item has no data (`nd`). */
  nd?: boolean;
}

/** A pivot table a slicer cache is connected to. */
export interface SlicerPivotTable {
  /** Index of the sheet holding the pivot table (`tabId`). */
  tabId: number;
  /** Pivot table name. */
  name: string;
}

/**
 * A slicer of a worksheet, `xl/slicers/slicer{n}.xml`. The writer emits its own slicer
 * dialect rather than the MS-XLSX x14 form, and does not write the links Excel needs to
 * load a slicer; the reader accepts both forms.
 */
export interface SlicerData {
  name: string;
  caption?: string;
  /** The `name` of the {@link SlicerCacheData} the slicer shows. */
  cacheName: string;
  columnName?: string;
  sortOrder?: SortOrder;
  /** Index of the first visible item. */
  startItem?: number;
}

/** A slicer cache, `xl/slicerCaches/slicerCache{n}.xml`: the items a slicer filters by. */
export interface SlicerCacheData {
  name: string;
  sourceName?: string;
  items: SlicerItem[];
  /** `pivotCacheId` of the pivot cache the slicer filters. Read from an Excel (x14) file only; the writer does not emit it. */
  pivotCacheId?: number | null;
  /** Sort order of the slicer items. Read from an Excel (x14) file only; the writer does not emit it. */
  sortOrder?: SortOrder | null;
  /** How items without data are treated (`crossFilter`), e.g. `'showItemsWithDataAtTop'`. Read from an Excel (x14) file only; the writer does not emit it. */
  crossFilter?: string | null;
  /** Pivot tables the slicer cache is connected to. Read from an Excel (x14) file only; the writer does not emit it. */
  pivotTables?: SlicerPivotTable[];
}

// ---------------------------------------------------------------------------
// Timelines
// ---------------------------------------------------------------------------

/**
 * Date granularity of a timeline. The writer emits the name; the reader also accepts the
 * MS-XLSX x15 numbers `0` (years) to `3` (days).
 */
export type TimelineLevel = 'years' | 'quarters' | 'months' | 'days';

/**
 * A timeline of a worksheet, `xl/timelines/timeline{n}.xml`. The writer emits its own
 * timeline dialect rather than the MS-XLSX x15 form, and does not write the links Excel
 * needs to load a timeline; the reader accepts both forms.
 */
export interface TimelineData {
  name: string;
  caption?: string;
  /** The `name` of the {@link TimelineCacheData} the timeline shows. */
  cacheName: string;
  sourceName?: string;
  level?: TimelineLevel;
  /** The level the current selection is made at. Read from an Excel (x15) file only; the writer does not emit it. */
  selectionLevel?: TimelineLevel | null;
}

/** A timeline cache, `xl/timelineCaches/timelineCache{n}.xml`: the source and selected date range. */
export interface TimelineCacheData {
  name: string;
  sourceName?: string;
  /** Selection start date (ISO 8601). */
  selectionStart?: string;
  /** Selection end date (ISO 8601). */
  selectionEnd?: string;
  /** `pivotCacheId` of the pivot cache the timeline filters. Read from an Excel (x15) file only; the writer does not emit it. */
  pivotCacheId?: number | null;
  /** Filter type of the timeline's state (`filterType`). Read from an Excel (x15) file only; the writer does not emit it. */
  filterType?: string | null;
  /** Start of the timeline's bounds. Read from an Excel (x15) file only; the writer does not emit it. */
  boundsStart?: string | null;
  /** End of the timeline's bounds. Read from an Excel (x15) file only; the writer does not emit it. */
  boundsEnd?: string | null;
}
