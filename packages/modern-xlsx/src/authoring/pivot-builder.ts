import type {
  PivotAxis,
  PivotDataFieldData,
  PivotFieldData,
  PivotFieldRef,
  PivotLocation,
  PivotPageFieldData,
  PivotTableData,
  SubtotalFunction,
} from '../model/types.js';

/**
 * Options for adding a row, column, or page field via the builder.
 */
export interface PivotFieldOptions {
  /** Source field index (0-based, position in the source data columns). */
  fieldIndex: number;
  /** Display name override. */
  name?: string;
  /** Subtotal functions, written as the field's `*Subtotal` flags. Empty (the default) writes no flag, so Excel shows its automatic subtotal. */
  subtotals?: SubtotalFunction[];
  /** Compact layout (default `true`). */
  compact?: boolean;
  /** Outline layout (default `true`). */
  outline?: boolean;
}

/**
 * Options for adding a data (values) field via the builder.
 */
export interface PivotDataFieldOptions {
  /** Source field index (0-based). */
  fieldIndex: number;
  /** Aggregation function (default `'sum'`). */
  subtotal?: SubtotalFunction;
  /** Caption of the data field (e.g. "Sum of Revenue"). Written on the data field only, never on the pivot field. */
  name?: string;
  /** Number format ID. */
  numFmtId?: number;
}

/**
 * Options for adding a page (filter) field via the builder.
 */
export interface PivotPageFieldOptions {
  /** Source field index (0-based). */
  fieldIndex: number;
  /** Selected item index. */
  item?: number;
  /** Display name override. */
  name?: string;
}

/**
 * Fluent builder for constructing `PivotTableData` objects.
 *
 * Every pivot field is built with empty `items`. On write, the row, column and page fields
 * get the items Excel needs from the cache that `cacheId` names, and that cache refreshes
 * when Excel opens the file, which lays the table out and shows its values (see
 * `PivotTableData`). With two or more data fields the writer also puts the values field on
 * the columns.
 *
 * The builder makes pivot fields only up to the highest `fieldIndex` used. On write,
 * `pivotFields` gets empty fields up to the number of fields in the cache that `cacheId`
 * names, because Excel refuses a table with fewer.
 *
 * @example
 * ```ts
 * ws.addPivotTableFromBuilder((b) => {
 *   b.name('SalesPivot')
 *    .cacheId(0)
 *    .location('A3:D20')
 *    .addRowField({ fieldIndex: 0, name: 'Region' })
 *    .addColField({ fieldIndex: 1, name: 'Product' })
 *    .addDataField({ fieldIndex: 2, subtotal: 'sum', name: 'Total Revenue' })
 *    .addPageField({ fieldIndex: 3, name: 'Year' });
 * });
 * ```
 */
export class PivotTableBuilder {
  #name = 'PivotTable1';
  #cacheId = 0;
  #location: PivotLocation = { ref: 'A3' };
  #dataCaption?: string;
  #fields: Map<number, PivotFieldData> = new Map();
  #rowFields: PivotFieldRef[] = [];
  #colFields: PivotFieldRef[] = [];
  #dataFields: PivotDataFieldData[] = [];
  #pageFields: PivotPageFieldData[] = [];

  /**
   * Set the pivot table name.
   *
   * @param value - The display name for the pivot table.
   * @returns `this` for chaining.
   */
  name(value: string): this {
    this.#name = value;
    return this;
  }

  /**
   * Set the pivot cache ID: the `cacheId` of a cache in the workbook's pivotCaches,
   * or that cache's 0-based index when its cacheId is unset. No cache is created
   * here; register one with `addPivotCache`.
   *
   * @param value - The cache's cacheId (or index when unset).
   * @returns `this` for chaining.
   */
  cacheId(value: number): this {
    this.#cacheId = value;
    return this;
  }

  /**
   * Set the location reference (e.g., `'A3:D20'` or `'A3'`).
   *
   * @param ref - The A1-style range or cell reference.
   * @param options - Optional first header row, data row, and data column offsets.
   * @returns `this` for chaining.
   */
  location(
    ref: string,
    options?: { firstHeaderRow?: number; firstDataRow?: number; firstDataCol?: number },
  ): this {
    this.#location = { ref, ...options };
    return this;
  }

  /**
   * Set the data caption text.
   *
   * @param value - The caption label for the data area.
   * @returns `this` for chaining.
   */
  dataCaption(value: string): this {
    this.#dataCaption = value;
    return this;
  }

  /**
   * Add a row field.
   *
   * @param opts - Field options including the source field index and display name.
   * @returns `this` for chaining.
   * @throws {RangeError} When the index is not an integer from 0 to 16383 or the field is already on an axis.
   */
  addRowField(opts: PivotFieldOptions): this {
    this.#ensureField(opts, 'axisRow');
    this.#rowFields.push({ x: opts.fieldIndex });
    return this;
  }

  /**
   * Add a column field.
   *
   * @param opts - Field options including the source field index and display name.
   * @returns `this` for chaining.
   * @throws {RangeError} When the index is not an integer from 0 to 16383 or the field is already on an axis.
   */
  addColField(opts: PivotFieldOptions): this {
    this.#ensureField(opts, 'axisCol');
    this.#colFields.push({ x: opts.fieldIndex });
    return this;
  }

  /**
   * Add a data (values) field.
   *
   * The same source field may be added again with another aggregation, and may
   * also sit on a row, column or page axis. The display name goes on the data
   * field only, never on the pivot field.
   *
   * @param opts - Data field options including field index, aggregation function, and display name.
   * @returns `this` for chaining.
   * @throws {RangeError} When the index is not an integer from 0 to 16383.
   */
  addDataField(opts: PivotDataFieldOptions): this {
    assertFieldIndex(opts.fieldIndex);
    this.#getOrCreateField(opts.fieldIndex).dataField = true;
    const entry: PivotDataFieldData = {
      fld: opts.fieldIndex,
      subtotal: opts.subtotal ?? 'sum',
    };
    if (opts.name !== undefined) entry.name = opts.name;
    if (opts.numFmtId !== undefined) entry.numFmtId = opts.numFmtId;
    this.#dataFields.push(entry);
    return this;
  }

  /**
   * Add a page (report filter) field.
   *
   * @param opts - Page field options including field index and optional selected item.
   * @returns `this` for chaining.
   * @throws {RangeError} When the index is not an integer from 0 to 16383 or the field is already on an axis.
   */
  addPageField(opts: PivotPageFieldOptions): this {
    this.#ensureField(
      { fieldIndex: opts.fieldIndex, ...(opts.name !== undefined && { name: opts.name }) },
      'axisPage',
    );
    const entry: PivotPageFieldData = { fld: opts.fieldIndex };
    if (opts.item !== undefined) entry.item = opts.item;
    if (opts.name !== undefined) entry.name = opts.name;
    this.#pageFields.push(entry);
    return this;
  }

  /**
   * Build the final `PivotTableData` object.
   *
   * Returns an independent deep copy; later builder calls do not affect it. Every pivot
   * field has empty `items`, which the writer completes from the cache (see
   * `PivotFieldData.items`).
   *
   * @returns The complete pivot table data ready for insertion via `Worksheet.addPivotTable()`.
   */
  build(): PivotTableData {
    // Build pivotFields array sorted by field index
    const maxIndex = Math.max(0, ...this.#fields.keys());
    const pivotFields: PivotFieldData[] = [];
    for (let i = 0; i <= maxIndex; i++) {
      pivotFields.push(
        this.#fields.get(i) ?? {
          items: [],
          subtotals: [],
          compact: true,
          outline: true,
        },
      );
    }

    const result: PivotTableData = {
      name: this.#name,
      location: this.#location,
      pivotFields,
      rowFields: this.#rowFields,
      colFields: this.#colFields,
      dataFields: this.#dataFields,
      pageFields: this.#pageFields,
      cacheId: this.#cacheId,
    };
    if (this.#dataCaption !== undefined) result.dataCaption = this.#dataCaption;
    return structuredClone(result);
  }

  #getOrCreateField(index: number): PivotFieldData {
    let field = this.#fields.get(index);
    if (!field) {
      field = { items: [], subtotals: [], compact: true, outline: true };
      this.#fields.set(index, field);
    }
    return field;
  }

  /** Put a field on an axis (a data-only placeholder gets its axis here); throws when the field already has one. Name, subtotals, compact and outline are set when given. */
  #ensureField(opts: PivotFieldOptions, axis: PivotAxis): void {
    assertFieldIndex(opts.fieldIndex);
    const field = this.#getOrCreateField(opts.fieldIndex);
    if (field.axis !== undefined) {
      throw new RangeError(
        `pivot field ${opts.fieldIndex} is already on axis '${field.axis}'; cannot add it to '${axis}'`,
      );
    }
    field.axis = axis;
    if (opts.name !== undefined) field.name = opts.name;
    if (opts.subtotals !== undefined) field.subtotals = [...opts.subtotals];
    if (opts.compact !== undefined) field.compact = opts.compact;
    if (opts.outline !== undefined) field.outline = opts.outline;
  }
}

function assertFieldIndex(index: number): void {
  if (!Number.isInteger(index) || index < 0 || index > 16383) {
    throw new RangeError(`pivot fieldIndex must be an integer from 0 to 16383, got ${index}`);
  }
}
