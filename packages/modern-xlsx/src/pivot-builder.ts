import type {
  PivotAxis,
  PivotDataFieldData,
  PivotFieldData,
  PivotFieldRef,
  PivotLocation,
  PivotPageFieldData,
  PivotTableData,
  SubtotalFunction,
} from './types.js';

/**
 * Options for adding a row, column, or page field via the builder.
 */
export interface PivotFieldOptions {
  /** Source field index (0-based, position in the source data columns). */
  fieldIndex: number;
  /** Display name override. */
  name?: string;
  /** Subtotal functions (defaults to none). */
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
  /** Display name (e.g. "Sum of Revenue"). */
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
   * Set the pivot cache ID (references workbook-level pivotCaches array index).
   *
   * @param value - The zero-based cache index.
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
   */
  addColField(opts: PivotFieldOptions): this {
    this.#ensureField(opts, 'axisCol');
    this.#colFields.push({ x: opts.fieldIndex });
    return this;
  }

  /**
   * Add a data (values) field.
   *
   * @param opts - Data field options including field index, aggregation function, and display name.
   * @returns `this` for chaining.
   */
  addDataField(opts: PivotDataFieldOptions): this {
    this.#ensureField(
      { fieldIndex: opts.fieldIndex, ...(opts.name !== undefined && { name: opts.name }) },
      'axisValues',
    );
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
    return result;
  }

  #ensureField(opts: PivotFieldOptions, axis: PivotAxis): void {
    if (!this.#fields.has(opts.fieldIndex)) {
      const field: PivotFieldData = {
        axis,
        items: [],
        subtotals: opts.subtotals ?? [],
        compact: opts.compact ?? true,
        outline: opts.outline ?? true,
      };
      if (opts.name !== undefined) field.name = opts.name;
      this.#fields.set(opts.fieldIndex, field);
    }
  }
}
