# Table Layout Engine Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add a high-level Table Layout Engine that generates styled XLSX tables from simple declarative options — headers, rows, styling, merging, zebra striping, auto-width, and more.

**Architecture:** A pure TypeScript module (`packages/modern-xlsx/src/table.ts`) that composes existing primitives (`Worksheet.cell()`, `StyleBuilder`, `addMergeCell`, `setColumnWidth`, `setRowHeight`, `frozenPane`). No Rust/WASM changes needed. The engine pre-builds a style palette (header, body-even, body-odd, total-row variants x border positions), then writes cells in a single pass over the grid.

**Tech Stack:** TypeScript 6.0, Vitest 4.1, Biome 2.4. Builds on existing `Workbook`, `Worksheet`, `Cell`, `StyleBuilder`, `columnToLetter`, `encodeCellRef`, `encodeRange` from `modern-xlsx`.

---

## Task 1: Core `drawTable` — basic table with headers and rows

**Files:**
- Create: `packages/modern-xlsx/src/table.ts`
- Modify: `packages/modern-xlsx/src/index.ts`
- Test: `packages/modern-xlsx/__tests__/table.test.ts`

**Step 1: Write the failing test**

```typescript
// packages/modern-xlsx/__tests__/table.test.ts
import { describe, expect, it } from 'vitest';
import { Workbook, drawTable } from '../src/index.js';

describe('drawTable', () => {
  it('creates a basic table with headers and rows', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Name', 'Age', 'City'],
      rows: [
        ['Alice', 30, 'NYC'],
        ['Bob', 25, 'LA'],
      ],
    });

    // Header row (row 1)
    expect(ws.cell('A1').value).toBe('Name');
    expect(ws.cell('B1').value).toBe('Age');
    expect(ws.cell('C1').value).toBe('City');

    // Data rows
    expect(ws.cell('A2').value).toBe('Alice');
    expect(ws.cell('B2').value).toBe(30);
    expect(ws.cell('C2').value).toBe('NYC');
    expect(ws.cell('A3').value).toBe('Bob');
    expect(ws.cell('B3').value).toBe(25);
    expect(ws.cell('C3').value).toBe('LA');
  });

  it('applies header styling (bold, background, center)', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Name', 'Age'],
      rows: [['Alice', 30]],
    });

    // Headers should have a styleIndex assigned
    const headerCell = ws.cell('A1');
    expect(headerCell.styleIndex).not.toBeNull();
    expect(headerCell.styleIndex).toBeGreaterThan(0);

    // Verify the style is bold with fill
    const xf = wb.styles.cellXfs[headerCell.styleIndex!];
    expect(wb.styles.fonts[xf.fontId].bold).toBe(true);
    expect(wb.styles.fills[xf.fillId].patternType).toBe('solid');
    expect(xf.alignment?.horizontal).toBe('center');
  });

  it('applies thin borders to all cells', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['A', 'B'],
      rows: [['1', '2']],
    });

    const cell = ws.cell('A1');
    const xf = wb.styles.cellXfs[cell.styleIndex!];
    const border = wb.styles.borders[xf.borderId];
    expect(border.top?.style).toBe('thin');
    expect(border.bottom?.style).toBe('thin');
    expect(border.left?.style).toBe('thin');
    expect(border.right?.style).toBe('thin');
  });

  it('supports origin option to place table at offset', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['X', 'Y'],
      rows: [['1', '2']],
      origin: 'C5',
    });

    expect(ws.cell('C5').value).toBe('X');
    expect(ws.cell('D5').value).toBe('Y');
    expect(ws.cell('C6').value).toBe('1');
    expect(ws.cell('D6').value).toBe('2');
  });

  it('returns TableResult with range and dimensions', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    const result = drawTable(wb, ws, {
      headers: ['Name', 'Age'],
      rows: [['Alice', 30], ['Bob', 25]],
    });

    expect(result.range).toBe('A1:B3');
    expect(result.rowCount).toBe(3); // 1 header + 2 data
    expect(result.colCount).toBe(2);
  });
});
```

**Step 2: Run test to verify it fails**

Run: `pnpm -C packages/modern-xlsx test -- --run table.test.ts`
Expected: FAIL — `drawTable` is not exported

**Step 3: Write minimal implementation**

```typescript
// packages/modern-xlsx/src/table.ts
import { columnToLetter, decodeCellRef, encodeRange, encodeCellRef } from './cell-ref.js';
import { StyleBuilder } from './style-builder.js';
import type { AlignmentData, BorderData, BorderStyle, FontData, StylesData } from './types.js';
import type { Workbook, Worksheet } from './workbook.js';

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** Result returned by drawTable with layout metadata. */
export interface TableResult {
  /** A1-style range covering the entire table (e.g. "A1:D10"). */
  range: string;
  /** Total number of rows (header + data + optional total). */
  rowCount: number;
  /** Number of columns. */
  colCount: number;
  /** 0-based row index of the first data row in the worksheet. */
  firstDataRow: number;
  /** 0-based row index of the last data row in the worksheet. */
  lastDataRow: number;
}

/** Per-cell style override. */
export interface CellStyle {
  font?: Partial<FontData>;
  fill?: { pattern?: string; fgColor?: string };
  border?: Partial<BorderData>;
  alignment?: Partial<AlignmentData>;
  numberFormat?: string;
}

/** Column definition for drawTable. */
export interface TableColumn {
  /** Header label. If omitted, uses the corresponding headers[] entry. */
  header?: string;
  /** Fixed width in Excel character units. */
  width?: number;
  /** Horizontal alignment for data cells in this column. */
  align?: 'left' | 'center' | 'right';
  /** Number format code for data cells (e.g. '#,##0.00'). */
  numberFormat?: string;
}

/** Options for drawTable. */
export interface DrawTableOptions {
  /** Header labels. */
  headers: string[];
  /** Data rows — each row is an array of cell values. */
  rows: (string | number | boolean | null)[][];
  /** Column definitions for width, alignment, and number format. */
  columns?: TableColumn[];
  /** Fixed column widths array (shorthand for columns[].width). */
  columnWidths?: number[];
  /** Origin cell reference (default: "A1"). */
  origin?: string;

  // --- Styling ---
  /** Header row font. Default: bold white. */
  headerFont?: Partial<FontData>;
  /** Header row background color (ARGB hex). Default: '4472C4'. */
  headerColor?: string;
  /** Body font. Default: Aptos 11. */
  bodyFont?: Partial<FontData>;
  /** Border style for all cells. Default: 'thin'. Set to null for no borders. */
  borderStyle?: BorderStyle | null;
  /** Border color (ARGB hex). Default: '000000'. */
  borderColor?: string;

  // --- Zebra striping ---
  /** Alternating row background color (ARGB hex). Null = no striping. */
  alternateRowColor?: string | null;

  // --- Alignment ---
  /** Default horizontal alignment for header cells. Default: 'center'. */
  headerAlign?: 'left' | 'center' | 'right';
  /** Default horizontal alignment for body cells. Default: undefined (general). */
  bodyAlign?: 'left' | 'center' | 'right';
  /** Default vertical alignment for all cells. Default: undefined. */
  verticalAlign?: 'top' | 'center' | 'bottom';

  // --- Content ---
  /** Enable text wrapping in body cells. Default: false. */
  wrapText?: boolean;
  /** Auto-calculate column widths from content. Default: false. */
  autoWidth?: boolean;
  /** Freeze the header row. Default: false. */
  freezeHeader?: boolean;
  /** Add auto-filter dropdowns to headers. Default: false. */
  autoFilter?: boolean;

  // --- Merge cells ---
  /** Cells to merge, as array of {row, col, rowSpan, colSpan}. Row/col are 0-based relative to data area. */
  merges?: { row: number; col: number; rowSpan?: number; colSpan?: number }[];

  // --- Per-cell styling ---
  /** Per-cell style overrides. Key is "row,col" (0-based, data area). */
  cellStyles?: Record<string, CellStyle>;
}

// ---------------------------------------------------------------------------
// Style palette — pre-built style indices for all table regions
// ---------------------------------------------------------------------------

interface StylePalette {
  header: number;
  bodyEven: number;
  bodyOdd: number;
  /** Per-column body style overrides (for alignment/numberFormat). */
  colBodyEven: (number | null)[];
  colBodyOdd: (number | null)[];
}

function buildPalette(wb: Workbook, opts: DrawTableOptions, colCount: number): StylePalette {
  const styles = wb.styles;
  const bs = opts.borderStyle === undefined ? 'thin' : opts.borderStyle;
  const bc = opts.borderColor ?? '000000';
  const borderObj = bs
    ? {
        left: { style: bs, color: bc },
        right: { style: bs, color: bc },
        top: { style: bs, color: bc },
        bottom: { style: bs, color: bc },
      }
    : undefined;

  // Header style
  const header = new StyleBuilder()
    .font({
      bold: true,
      color: 'FFFFFF',
      ...(opts.headerFont ?? {}),
    })
    .fill({ pattern: 'solid', fgColor: opts.headerColor ?? '4472C4' });
  if (borderObj) header.border(borderObj);
  header.alignment({
    horizontal: opts.headerAlign ?? 'center',
    vertical: opts.verticalAlign ?? null,
  });
  const headerIdx = header.build(styles);

  // Body even style
  const bodyEvenIdx = buildBodyStyle(styles, opts, borderObj, false);

  // Body odd style (with alternate row color)
  const bodyOddIdx = opts.alternateRowColor
    ? buildBodyStyle(styles, opts, borderObj, true)
    : bodyEvenIdx;

  // Per-column styles (if columns have custom alignment or numberFormat)
  const colBodyEven: (number | null)[] = Array.from({ length: colCount }, () => null);
  const colBodyOdd: (number | null)[] = Array.from({ length: colCount }, () => null);

  if (opts.columns) {
    for (let c = 0; c < colCount; c++) {
      const col = opts.columns[c];
      if (!col) continue;
      if (col.align || col.numberFormat) {
        colBodyEven[c] = buildColStyle(styles, opts, borderObj, false, col);
        colBodyOdd[c] = opts.alternateRowColor
          ? buildColStyle(styles, opts, borderObj, true, col)
          : colBodyEven[c];
      }
    }
  }

  return { header: headerIdx, bodyEven: bodyEvenIdx, bodyOdd: bodyOddIdx, colBodyEven, colBodyOdd };
}

function buildBodyStyle(
  styles: StylesData,
  opts: DrawTableOptions,
  borderObj: Parameters<StyleBuilder['border']>[0] | undefined,
  isOdd: boolean,
): number {
  const sb = new StyleBuilder();
  if (opts.bodyFont) sb.font(opts.bodyFont);
  if (isOdd && opts.alternateRowColor) {
    sb.fill({ pattern: 'solid', fgColor: opts.alternateRowColor });
  }
  if (borderObj) sb.border(borderObj);
  const align: Partial<AlignmentData> = {};
  if (opts.bodyAlign) align.horizontal = opts.bodyAlign;
  if (opts.verticalAlign) align.vertical = opts.verticalAlign;
  if (opts.wrapText) align.wrapText = true;
  if (Object.keys(align).length > 0) sb.alignment(align);
  return sb.build(styles);
}

function buildColStyle(
  styles: StylesData,
  opts: DrawTableOptions,
  borderObj: Parameters<StyleBuilder['border']>[0] | undefined,
  isOdd: boolean,
  col: TableColumn,
): number {
  const sb = new StyleBuilder();
  if (opts.bodyFont) sb.font(opts.bodyFont);
  if (isOdd && opts.alternateRowColor) {
    sb.fill({ pattern: 'solid', fgColor: opts.alternateRowColor });
  }
  if (borderObj) sb.border(borderObj);
  const align: Partial<AlignmentData> = {};
  if (col.align) align.horizontal = col.align;
  else if (opts.bodyAlign) align.horizontal = opts.bodyAlign;
  if (opts.verticalAlign) align.vertical = opts.verticalAlign;
  if (opts.wrapText) align.wrapText = true;
  if (Object.keys(align).length > 0) sb.alignment(align);
  if (col.numberFormat) sb.numberFormat(col.numberFormat);
  return sb.build(styles);
}

// ---------------------------------------------------------------------------
// Auto-width calculation
// ---------------------------------------------------------------------------

/** Estimate column width in Excel character units from content. */
function estimateWidth(value: string | number | boolean | null | undefined): number {
  if (value == null) return 8;
  const str = String(value);
  // Count characters — CJK/emoji count as 2
  let len = 0;
  for (const ch of str) {
    const code = ch.codePointAt(0)!;
    len += code > 0x2e7f ? 2 : 1;
  }
  return Math.max(8, Math.ceil(len * 1.2) + 2);
}

function computeAutoWidths(headers: string[], rows: (string | number | boolean | null)[][]): number[] {
  const widths = headers.map((h) => estimateWidth(h));
  for (const row of rows) {
    for (let c = 0; c < widths.length; c++) {
      const w = estimateWidth(row[c]);
      if (w > widths[c]!) widths[c] = w;
    }
  }
  return widths;
}

// ---------------------------------------------------------------------------
// Per-cell style override
// ---------------------------------------------------------------------------

function buildCellOverrideStyle(
  wb: Workbook,
  baseIdx: number,
  override: CellStyle,
): number {
  const styles = wb.styles;
  const base = styles.cellXfs[baseIdx];
  const sb = new StyleBuilder();

  // Merge base font with override
  const baseFont = styles.fonts[base?.fontId ?? 0];
  sb.font({ ...baseFont, ...(override.font ?? {}) });

  // Merge fill
  const baseFill = styles.fills[base?.fillId ?? 0];
  if (override.fill) {
    sb.fill({
      pattern: (override.fill.pattern ?? baseFill?.patternType ?? 'none') as any,
      fgColor: override.fill.fgColor ?? baseFill?.fgColor,
    });
  } else if (baseFill && baseFill.patternType !== 'none') {
    sb.fill({ pattern: baseFill.patternType as any, fgColor: baseFill.fgColor });
  }

  // Merge border
  const baseBorder = styles.borders[base?.borderId ?? 0];
  const mergedBorder = { ...baseBorder, ...(override.border ?? {}) };
  if (mergedBorder.left || mergedBorder.right || mergedBorder.top || mergedBorder.bottom) {
    sb.border(mergedBorder as any);
  }

  // Merge alignment
  const baseAlign = base?.alignment ?? {};
  if (override.alignment || Object.keys(baseAlign).length > 0) {
    sb.alignment({ ...baseAlign, ...(override.alignment ?? {}) });
  }

  // Number format
  if (override.numberFormat) {
    sb.numberFormat(override.numberFormat);
  } else if (base?.numFmtId && base.numFmtId > 0) {
    const fmt = styles.numFmts.find((f) => f.id === base.numFmtId);
    if (fmt) sb.numberFormat(fmt.formatCode);
  }

  return sb.build(styles);
}

// ---------------------------------------------------------------------------
// drawTable — main entry point
// ---------------------------------------------------------------------------

/**
 * Draw a styled table on a worksheet.
 *
 * Creates header row, data rows, applies styles, borders, column widths,
 * merge cells, zebra striping, frozen panes, and auto-filter.
 */
export function drawTable(wb: Workbook, ws: Worksheet, opts: DrawTableOptions): TableResult {
  const { headers, rows } = opts;
  const colCount = headers.length;
  const origin = opts.origin ? decodeCellRef(opts.origin) : { row: 0, col: 0 };

  // Build style palette once
  const palette = buildPalette(wb, opts, colCount);

  // --- Column widths ---
  const widths = opts.autoWidth
    ? computeAutoWidths(headers, rows)
    : opts.columnWidths ?? opts.columns?.map((c) => c.width ?? 0) ?? [];
  for (let c = 0; c < colCount; c++) {
    const w = widths[c];
    if (w && w > 0) ws.setColumnWidth(origin.col + c + 1, w);
  }

  // --- Header row ---
  const headerRowIdx = origin.row + 1; // 1-based worksheet row
  for (let c = 0; c < colCount; c++) {
    const ref = `${columnToLetter(origin.col + c)}${headerRowIdx}`;
    const cell = ws.cell(ref);
    cell.value = headers[c] ?? '';
    cell.styleIndex = palette.header;
  }

  // --- Data rows ---
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]!;
    const wsRowIdx = origin.row + r + 2; // 1-based, after header
    const isOdd = r % 2 === 1;

    for (let c = 0; c < colCount; c++) {
      const ref = `${columnToLetter(origin.col + c)}${wsRowIdx}`;
      const cell = ws.cell(ref);
      const val = row[c];
      if (val !== null && val !== undefined) {
        cell.value = val;
      }

      // Determine style: per-column override > zebra stripe > base body
      const colStyle = isOdd ? palette.colBodyOdd[c] : palette.colBodyEven[c];
      let styleIdx = colStyle ?? (isOdd ? palette.bodyOdd : palette.bodyEven);

      // Per-cell style override
      const overrideKey = `${r},${c}`;
      if (opts.cellStyles?.[overrideKey]) {
        styleIdx = buildCellOverrideStyle(wb, styleIdx, opts.cellStyles[overrideKey]!);
      }

      cell.styleIndex = styleIdx;
    }
  }

  // --- Merge cells ---
  if (opts.merges) {
    for (const merge of opts.merges) {
      const startRow = origin.row + merge.row + 1; // +1 for header
      const startCol = origin.col + merge.col;
      const endRow = startRow + (merge.rowSpan ?? 1) - 1;
      const endCol = startCol + (merge.colSpan ?? 1) - 1;
      const rangeStr = encodeRange(
        { row: startRow, col: startCol },
        { row: endRow, col: endCol },
      );
      ws.addMergeCell(rangeStr);
    }
  }

  // --- Frozen header ---
  if (opts.freezeHeader) {
    ws.frozenPane = {
      row: 1,
      column: 0,
      topLeftCell: `A${headerRowIdx + 1}`,
    };
  }

  // --- Auto filter ---
  if (opts.autoFilter) {
    const filterRange = encodeRange(
      { row: origin.row, col: origin.col },
      { row: origin.row, col: origin.col + colCount - 1 },
    );
    ws.autoFilter = filterRange;
  }

  // --- Result ---
  const totalRows = 1 + rows.length;
  const rangeStr = encodeRange(
    { row: origin.row, col: origin.col },
    { row: origin.row + totalRows - 1, col: origin.col + colCount - 1 },
  );

  return {
    range: rangeStr,
    rowCount: totalRows,
    colCount,
    firstDataRow: origin.row + 1,
    lastDataRow: origin.row + rows.length,
  };
}

// ---------------------------------------------------------------------------
// drawTableFromData — create table from JSON array
// ---------------------------------------------------------------------------

/** Options for drawTableFromData. */
export interface DrawTableFromDataOptions extends Omit<DrawTableOptions, 'headers' | 'rows'> {
  /** Explicit header order/labels. If omitted, uses Object.keys of first item. */
  headers?: string[];
  /** Map object keys to display headers. */
  headerMap?: Record<string, string>;
}

/**
 * Create a styled table from a JSON array.
 *
 * Automatically extracts headers from object keys and maps values to rows.
 */
export function drawTableFromData(
  wb: Workbook,
  ws: Worksheet,
  data: Record<string, unknown>[],
  opts?: DrawTableFromDataOptions,
): TableResult {
  if (data.length === 0) {
    return drawTable(wb, ws, { headers: [], rows: [], ...opts });
  }

  const keys = opts?.headers ?? Object.keys(data[0]!);
  const headers = keys.map((k) => opts?.headerMap?.[k] ?? k);

  const rows: (string | number | boolean | null)[][] = data.map((item) =>
    keys.map((key) => {
      const val = item[key];
      if (val === null || val === undefined) return null;
      if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return val;
      return String(val);
    }),
  );

  return drawTable(wb, ws, { ...opts, headers, rows });
}
```

**Step 4: Add exports to index.ts**

Add to `packages/modern-xlsx/src/index.ts`:
```typescript
// Table layout engine
export type {
  CellStyle,
  DrawTableFromDataOptions,
  DrawTableOptions,
  TableColumn,
  TableResult,
} from './table.js';
export { drawTable, drawTableFromData } from './table.js';
```

**Step 5: Run test to verify it passes**

Run: `pnpm -C packages/modern-xlsx test -- --run table.test.ts`
Expected: PASS (all 5 tests)

**Step 6: Commit**

```bash
git add packages/modern-xlsx/src/table.ts packages/modern-xlsx/src/index.ts packages/modern-xlsx/__tests__/table.test.ts
git commit -m "feat: table layout engine — drawTable, drawTableFromData with styled headers, borders, origin"
```

---

## Task 2: Auto-width calculation from content

**Files:**
- Modify: `packages/modern-xlsx/src/table.ts` (already has `computeAutoWidths`)
- Test: `packages/modern-xlsx/__tests__/table.test.ts`

**Step 1: Write the failing test**

```typescript
describe('auto-width', () => {
  it('calculates column widths from content', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['ID', 'Full Name', 'Description'],
      rows: [
        [1, 'Alice Wonderland', 'A very long description that should make the column wider'],
        [2, 'Bob', 'Short'],
      ],
      autoWidth: true,
    });

    // Column C should be widest (long description)
    const cols = ws.columns;
    expect(cols.length).toBeGreaterThanOrEqual(3);
    const colC = cols.find((c) => c.min === 3);
    expect(colC).toBeDefined();
    expect(colC!.width).toBeGreaterThan(20);

    // Column A should be narrowest (just "ID" and numbers)
    const colA = cols.find((c) => c.min === 1);
    expect(colA).toBeDefined();
    expect(colA!.width).toBeLessThan(colC!.width);
  });

  it('handles CJK characters as double-width', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Name'],
      rows: [['Hello'], ['你好世界测试']],
      autoWidth: true,
    });

    const col = ws.columns.find((c) => c.min === 1);
    expect(col).toBeDefined();
    // 6 CJK chars = 12 effective chars, so width > 14
    expect(col!.width).toBeGreaterThan(14);
  });
});
```

**Step 2: Run test to verify it passes**

These tests should pass with the implementation from Task 1 (auto-width is already implemented in `computeAutoWidths`).

Run: `pnpm -C packages/modern-xlsx test -- --run table.test.ts`
Expected: PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/__tests__/table.test.ts
git commit -m "test: auto-width column calculation with CJK support"
```

---

## Task 3: Per-cell styling — borders, background, font overrides

**Files:**
- Test: `packages/modern-xlsx/__tests__/table.test.ts`

**Step 1: Write the failing test**

```typescript
describe('per-cell styling', () => {
  it('applies cellStyles overrides to specific cells', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Item', 'Amount'],
      rows: [
        ['Widget', 100],
        ['Gadget', -50],
      ],
      cellStyles: {
        '1,1': { font: { color: 'FF0000', bold: true } }, // Gadget amount = red bold
      },
    });

    // Row 1 (0-based data), Col 1 = B3
    const cell = ws.cell('B3');
    expect(cell.styleIndex).not.toBeNull();
    const xf = wb.styles.cellXfs[cell.styleIndex!];
    const font = wb.styles.fonts[xf.fontId];
    expect(font.color).toBe('FF0000');
    expect(font.bold).toBe(true);
  });

  it('merges cell override with base body style (preserves borders)', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['A'],
      rows: [['val']],
      cellStyles: {
        '0,0': { fill: { pattern: 'solid', fgColor: 'FFFF00' } },
      },
    });

    const cell = ws.cell('A2');
    const xf = wb.styles.cellXfs[cell.styleIndex!];
    // Should still have borders from the base style
    const border = wb.styles.borders[xf.borderId];
    expect(border.top?.style).toBe('thin');
    // Should have yellow fill
    const fill = wb.styles.fills[xf.fillId];
    expect(fill.fgColor).toBe('FFFF00');
  });
});
```

**Step 2: Run test to verify it passes**

These should pass with the `cellStyles` + `buildCellOverrideStyle` implementation from Task 1.

Run: `pnpm -C packages/modern-xlsx test -- --run table.test.ts`
Expected: PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/__tests__/table.test.ts
git commit -m "test: per-cell styling overrides with font, fill, border merging"
```

---

## Task 4: Row and column spanning (merge cells)

**Files:**
- Test: `packages/modern-xlsx/__tests__/table.test.ts`

**Step 1: Write the failing test**

```typescript
describe('merge cells', () => {
  it('spans columns for a header-like merge', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Product', 'Q1', 'Q2', 'Q3'],
      rows: [
        ['Widget', 100, 200, 150],
        ['Subtotal', null, null, 450],
      ],
      merges: [
        { row: 1, col: 0, colSpan: 3 }, // "Subtotal" spans 3 columns
      ],
    });

    expect(ws.mergeCells).toContain('A3:C3');
    expect(ws.cell('A3').value).toBe('Subtotal');
  });

  it('spans rows for a category label', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Category', 'Item', 'Price'],
      rows: [
        ['Electronics', 'Phone', 699],
        [null, 'Laptop', 1299],
        [null, 'Tablet', 499],
      ],
      merges: [
        { row: 0, col: 0, rowSpan: 3 }, // "Electronics" spans 3 rows
      ],
    });

    expect(ws.mergeCells).toContain('A2:A4');
    expect(ws.cell('A2').value).toBe('Electronics');
  });
});
```

**Step 2: Run test to verify it passes**

Run: `pnpm -C packages/modern-xlsx test -- --run table.test.ts`
Expected: PASS (merge logic already implemented in Task 1)

**Step 3: Commit**

```bash
git add packages/modern-xlsx/__tests__/table.test.ts
git commit -m "test: row and column spanning with merge cells"
```

---

## Task 5: Alternating row colors (zebra striping)

**Files:**
- Test: `packages/modern-xlsx/__tests__/table.test.ts`

**Step 1: Write the failing test**

```typescript
describe('zebra striping', () => {
  it('applies alternating row background', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Name', 'Value'],
      rows: [
        ['A', 1],
        ['B', 2],
        ['C', 3],
        ['D', 4],
      ],
      alternateRowColor: 'E8EDF3',
    });

    // Even rows (0, 2) should NOT have fill (default body style)
    const evenXf = wb.styles.cellXfs[ws.cell('A2').styleIndex!];
    const evenFill = wb.styles.fills[evenXf.fillId];
    expect(evenFill.patternType).toBe('none');

    // Odd rows (1, 3) should have the alternate color
    const oddXf = wb.styles.cellXfs[ws.cell('A3').styleIndex!];
    const oddFill = wb.styles.fills[oddXf.fillId];
    expect(oddFill.patternType).toBe('solid');
    expect(oddFill.fgColor).toBe('E8EDF3');

    // Row 4 (index 2, even) = no fill
    const even2Xf = wb.styles.cellXfs[ws.cell('A4').styleIndex!];
    const even2Fill = wb.styles.fills[even2Xf.fillId];
    expect(even2Fill.patternType).toBe('none');

    // Row 5 (index 3, odd) = fill
    const odd2Xf = wb.styles.cellXfs[ws.cell('A5').styleIndex!];
    const odd2Fill = wb.styles.fills[odd2Xf.fillId];
    expect(odd2Fill.fgColor).toBe('E8EDF3');
  });
});
```

**Step 2: Run test to verify it passes**

Run: `pnpm -C packages/modern-xlsx test -- --run table.test.ts`
Expected: PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/__tests__/table.test.ts
git commit -m "test: zebra striping with alternateRowColor"
```

---

## Task 6: Cell alignment — horizontal + vertical

**Files:**
- Test: `packages/modern-xlsx/__tests__/table.test.ts`

**Step 1: Write the failing test**

```typescript
describe('cell alignment', () => {
  it('applies horizontal alignment to header and body', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Left', 'Center', 'Right'],
      rows: [['a', 'b', 'c']],
      columns: [
        { align: 'left' },
        { align: 'center' },
        { align: 'right' },
      ],
    });

    // Body cells should respect per-column alignment
    const leftXf = wb.styles.cellXfs[ws.cell('A2').styleIndex!];
    expect(leftXf.alignment?.horizontal).toBe('left');

    const centerXf = wb.styles.cellXfs[ws.cell('B2').styleIndex!];
    expect(centerXf.alignment?.horizontal).toBe('center');

    const rightXf = wb.styles.cellXfs[ws.cell('C2').styleIndex!];
    expect(rightXf.alignment?.horizontal).toBe('right');
  });

  it('applies vertical alignment to all cells', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['A'],
      rows: [['val']],
      verticalAlign: 'center',
    });

    const headerXf = wb.styles.cellXfs[ws.cell('A1').styleIndex!];
    expect(headerXf.alignment?.vertical).toBe('center');

    const bodyXf = wb.styles.cellXfs[ws.cell('A2').styleIndex!];
    expect(bodyXf.alignment?.vertical).toBe('center');
  });
});
```

**Step 2: Run test to verify it passes**

Run: `pnpm -C packages/modern-xlsx test -- --run table.test.ts`
Expected: PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/__tests__/table.test.ts
git commit -m "test: horizontal and vertical cell alignment"
```

---

## Task 7: Cell content wrapping

**Files:**
- Test: `packages/modern-xlsx/__tests__/table.test.ts`

**Step 1: Write the failing test**

```typescript
describe('content wrapping', () => {
  it('enables wrapText on body cells', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Description'],
      rows: [['A very long description that should wrap within the cell']],
      wrapText: true,
    });

    const bodyXf = wb.styles.cellXfs[ws.cell('A2').styleIndex!];
    expect(bodyXf.alignment?.wrapText).toBe(true);
  });
});
```

**Step 2: Run test to verify it passes**

Run: `pnpm -C packages/modern-xlsx test -- --run table.test.ts`
Expected: PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/__tests__/table.test.ts
git commit -m "test: cell content wrapping with wrapText option"
```

---

## Task 8: Nested tables — table as cell content

**Files:**
- Modify: `packages/modern-xlsx/src/table.ts`
- Test: `packages/modern-xlsx/__tests__/table.test.ts`

**Step 1: Write the failing test**

```typescript
describe('nested tables', () => {
  it('draws a nested table below the parent table', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    // Parent table
    const result = drawTable(wb, ws, {
      headers: ['Order', 'Total'],
      rows: [['ORD-001', 500]],
    });

    // Nested table placed after parent
    const nested = drawTable(wb, ws, {
      headers: ['Item', 'Qty', 'Price'],
      rows: [
        ['Widget', 2, 100],
        ['Gadget', 3, 100],
      ],
      origin: `A${result.rowCount + 2}`, // leave a gap row
    });

    // Parent table cells
    expect(ws.cell('A1').value).toBe('Order');
    expect(ws.cell('A2').value).toBe('ORD-001');

    // Nested table cells (starts at row 4, with 1-row gap)
    expect(ws.cell('A4').value).toBe('Item');
    expect(ws.cell('A5').value).toBe('Widget');
    expect(ws.cell('A6').value).toBe('Gadget');

    expect(nested.range).toBe('A4:C6');
  });

  it('draws side-by-side tables using origin offset', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Left'],
      rows: [['L1'], ['L2']],
    });

    drawTable(wb, ws, {
      headers: ['Right'],
      rows: [['R1'], ['R2']],
      origin: 'C1',
    });

    expect(ws.cell('A1').value).toBe('Left');
    expect(ws.cell('C1').value).toBe('Right');
    expect(ws.cell('C2').value).toBe('R1');
  });
});
```

**Step 2: Run test to verify it passes**

This works via the existing `origin` option — nested tables are just additional `drawTable` calls with offset origins. No new implementation needed.

Run: `pnpm -C packages/modern-xlsx test -- --run table.test.ts`
Expected: PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/__tests__/table.test.ts
git commit -m "test: nested and side-by-side tables using origin offset"
```

---

## Task 9: drawTableFromData — table from JSON/CSV data

**Files:**
- Test: `packages/modern-xlsx/__tests__/table.test.ts`

**Step 1: Write the failing test**

```typescript
describe('drawTableFromData', () => {
  it('creates table from JSON array', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    const result = drawTableFromData(wb, ws, [
      { name: 'Alice', age: 30, city: 'NYC' },
      { name: 'Bob', age: 25, city: 'LA' },
    ]);

    expect(ws.cell('A1').value).toBe('name');
    expect(ws.cell('B1').value).toBe('age');
    expect(ws.cell('C1').value).toBe('city');
    expect(ws.cell('A2').value).toBe('Alice');
    expect(ws.cell('B2').value).toBe(30);
    expect(result.rowCount).toBe(3);
    expect(result.colCount).toBe(3);
  });

  it('uses headerMap to rename columns', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTableFromData(wb, ws, [{ first_name: 'Alice', last_name: 'Smith' }], {
      headerMap: { first_name: 'First Name', last_name: 'Last Name' },
    });

    expect(ws.cell('A1').value).toBe('First Name');
    expect(ws.cell('B1').value).toBe('Last Name');
    expect(ws.cell('A2').value).toBe('Alice');
  });

  it('applies all drawTable options (zebra, autoWidth, etc.)', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTableFromData(
      wb,
      ws,
      [
        { id: 1, name: 'Widget' },
        { id: 2, name: 'Gadget' },
        { id: 3, name: 'Doohickey' },
      ],
      {
        alternateRowColor: 'F0F0F0',
        autoWidth: true,
        freezeHeader: true,
      },
    );

    // Verify zebra on row 3 (index 1, odd)
    const oddXf = wb.styles.cellXfs[ws.cell('A3').styleIndex!];
    const oddFill = wb.styles.fills[oddXf.fillId];
    expect(oddFill.fgColor).toBe('F0F0F0');

    // Verify frozen pane
    expect(ws.frozenPane).toBeDefined();
    expect(ws.frozenPane!.row).toBe(1);

    // Verify auto-width was applied
    expect(ws.columns.length).toBeGreaterThan(0);
  });

  it('handles empty data array', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    const result = drawTableFromData(wb, ws, []);
    expect(result.rowCount).toBe(1); // just header row (empty)
    expect(result.colCount).toBe(0);
  });
});
```

**Step 2: Run test to verify it passes**

Run: `pnpm -C packages/modern-xlsx test -- --run table.test.ts`
Expected: PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/__tests__/table.test.ts
git commit -m "test: drawTableFromData — JSON data, headerMap, options passthrough"
```

---

## Task 10: Freeze header + auto-filter integration

**Files:**
- Test: `packages/modern-xlsx/__tests__/table.test.ts`

**Step 1: Write the failing test**

```typescript
describe('freeze header & auto-filter', () => {
  it('freezes the header row', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['A', 'B'],
      rows: [['1', '2']],
      freezeHeader: true,
    });

    expect(ws.frozenPane).toEqual({
      row: 1,
      column: 0,
      topLeftCell: 'A2',
    });
  });

  it('adds auto-filter to header row', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Name', 'Age', 'City'],
      rows: [['Alice', 30, 'NYC']],
      autoFilter: true,
    });

    expect(ws.autoFilter).toBe('A1:C1');
  });
});
```

**Step 2: Run test to verify it passes**

Run: `pnpm -C packages/modern-xlsx test -- --run table.test.ts`
Expected: PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/__tests__/table.test.ts
git commit -m "test: freeze header and auto-filter integration"
```

---

## Task 11: Round-trip test — write table then read back

**Files:**
- Test: `packages/modern-xlsx/__tests__/table.test.ts`

**Step 1: Write the failing test**

```typescript
describe('round-trip', () => {
  it('writes a styled table and reads it back', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Invoice');

    drawTable(wb, ws, {
      headers: ['Item', 'Qty', 'Price', 'Total'],
      rows: [
        ['Widget', 10, 25.5, 255],
        ['Gadget', 5, 42.0, 210],
        ['Doohickey', 2, 99.99, 199.98],
      ],
      columnWidths: [20, 8, 12, 12],
      alternateRowColor: 'E8EDF3',
      freezeHeader: true,
      autoFilter: true,
      columns: [
        { align: 'left' },
        { align: 'center' },
        { align: 'right', numberFormat: '#,##0.00' },
        { align: 'right', numberFormat: '#,##0.00' },
      ],
    });

    // Write to buffer and read back
    const buffer = await wb.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Invoice');
    expect(ws2).toBeDefined();

    // Cell values preserved
    expect(ws2!.cell('A1').value).toBe('Item');
    expect(ws2!.cell('B2').value).toBe('10');
    expect(ws2!.cell('C3').value).toBe('42');

    // Merge cells, frozen pane, auto-filter preserved
    expect(ws2!.frozenPane).toBeDefined();
  });
});
```

Add `readBuffer` to the import at the top of the test file:
```typescript
import { Workbook, drawTable, drawTableFromData, readBuffer } from '../src/index.js';
```

**Step 2: Run test to verify it passes**

Run: `pnpm -C packages/modern-xlsx test -- --run table.test.ts`
Expected: PASS

**Step 3: Commit**

```bash
git add packages/modern-xlsx/__tests__/table.test.ts
git commit -m "test: table round-trip — write styled table, read back, verify"
```

---

## Task 12: Table layout guide documentation

**Files:**
- Create: `docs/guide/tables.md`

**Step 1: Write the documentation**

```markdown
# Table Layout Guide

Generate styled XLSX tables with a single function call — no manual cell coordinate math.

## Quick Start

```typescript
import { initWasm, Workbook, drawTable } from 'modern-xlsx';

await initWasm();
const wb = new Workbook();
const ws = wb.addSheet('Report');

drawTable(wb, ws, {
  headers: ['Product', 'Qty', 'Price', 'Total'],
  rows: [
    ['Widget', 10, 25.50, 255.00],
    ['Gadget', 5, 42.00, 210.00],
  ],
  alternateRowColor: 'E8EDF3',
  freezeHeader: true,
  autoFilter: true,
  autoWidth: true,
});

await wb.toFile('report.xlsx');
```

## Invoice Example

```typescript
const wb = new Workbook();
const ws = wb.addSheet('Invoice');

// Company header
ws.cell('A1').value = 'ACME Corporation';
ws.cell('A2').value = 'Invoice #INV-2026-001';
ws.cell('A3').value = 'Date: 2026-03-02';

// Line items table starting at row 5
drawTable(wb, ws, {
  headers: ['Item', 'Description', 'Qty', 'Unit Price', 'Total'],
  rows: [
    ['WDG-001', 'Premium Widget', 10, 25.50, 255.00],
    ['GDG-002', 'Standard Gadget', 5, 42.00, 210.00],
    ['DHK-003', 'Deluxe Doohickey', 2, 99.99, 199.98],
  ],
  origin: 'A5',
  headerColor: '1F4E79',
  alternateRowColor: 'D6E4F0',
  autoWidth: true,
  columns: [
    { align: 'left' },
    { align: 'left' },
    { align: 'center' },
    { align: 'right', numberFormat: '$#,##0.00' },
    { align: 'right', numberFormat: '$#,##0.00' },
  ],
});
```

## Financial Report with Zebra Striping

```typescript
drawTable(wb, ws, {
  headers: ['Account', 'Q1', 'Q2', 'Q3', 'Q4', 'Total'],
  rows: [
    ['Revenue', 120000, 135000, 142000, 158000, 555000],
    ['COGS', -48000, -54000, -57000, -63000, -222000],
    ['Gross Profit', 72000, 81000, 85000, 95000, 333000],
    ['Operating Exp', -30000, -32000, -33000, -35000, -130000],
    ['Net Income', 42000, 49000, 52000, 60000, 203000],
  ],
  alternateRowColor: 'F5F5FA',
  headerColor: '2D3748',
  columns: [
    { width: 20, align: 'left' },
    { width: 14, align: 'right', numberFormat: '$#,##0' },
    { width: 14, align: 'right', numberFormat: '$#,##0' },
    { width: 14, align: 'right', numberFormat: '$#,##0' },
    { width: 14, align: 'right', numberFormat: '$#,##0' },
    { width: 16, align: 'right', numberFormat: '$#,##0' },
  ],
  cellStyles: {
    '4,0': { font: { bold: true } },
    '4,5': { font: { bold: true } },
  },
  freezeHeader: true,
});
```

## Table from JSON Data

```typescript
import { drawTableFromData } from 'modern-xlsx';

const apiData = [
  { id: 1, name: 'Alice', department: 'Engineering', salary: 95000 },
  { id: 2, name: 'Bob', department: 'Marketing', salary: 72000 },
  { id: 3, name: 'Carol', department: 'Engineering', salary: 105000 },
];

drawTableFromData(wb, ws, apiData, {
  headerMap: {
    id: 'ID',
    name: 'Full Name',
    department: 'Dept',
    salary: 'Annual Salary',
  },
  columns: [
    { width: 6, align: 'center' },
    { width: 20, align: 'left' },
    { width: 16, align: 'left' },
    { width: 16, align: 'right', numberFormat: '$#,##0' },
  ],
  alternateRowColor: 'F0F4FF',
  autoFilter: true,
  freezeHeader: true,
});
```

## Merge Cells (Row & Column Spanning)

```typescript
drawTable(wb, ws, {
  headers: ['Category', 'Product', 'Q1', 'Q2'],
  rows: [
    ['Electronics', 'Phone', 500, 620],
    [null, 'Laptop', 300, 410],
    [null, 'Tablet', 200, 180],
    ['Clothing', 'Shirts', 1200, 1100],
    [null, 'Pants', 800, 750],
  ],
  merges: [
    { row: 0, col: 0, rowSpan: 3 },  // "Electronics" spans rows 0-2
    { row: 3, col: 0, rowSpan: 2 },  // "Clothing" spans rows 3-4
  ],
});
```

## Multiple Tables on One Sheet

```typescript
const result1 = drawTable(wb, ws, {
  headers: ['Summary'],
  rows: [['Overview data']],
});

// Place second table below the first with a gap
drawTable(wb, ws, {
  headers: ['Detail', 'Value'],
  rows: [['Item A', 100], ['Item B', 200]],
  origin: `A${result1.rowCount + 2}`,
});

// Side-by-side tables
drawTable(wb, ws, {
  headers: ['Left Table'],
  rows: [['L1'], ['L2']],
  origin: 'E1',
});
```

## API Reference

### `drawTable(wb, ws, options): TableResult`

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `headers` | `string[]` | required | Column header labels |
| `rows` | `(string\|number\|boolean\|null)[][]` | required | Data rows |
| `origin` | `string` | `'A1'` | Top-left cell reference |
| `columns` | `TableColumn[]` | — | Per-column width, align, format |
| `columnWidths` | `number[]` | — | Shorthand for fixed widths |
| `autoWidth` | `boolean` | `false` | Auto-calculate widths |
| `headerFont` | `Partial<FontData>` | bold white | Header font |
| `headerColor` | `string` | `'4472C4'` | Header background (hex) |
| `headerAlign` | `string` | `'center'` | Header alignment |
| `bodyFont` | `Partial<FontData>` | default | Body font |
| `bodyAlign` | `string` | — | Body alignment |
| `verticalAlign` | `string` | — | Vertical alignment |
| `borderStyle` | `BorderStyle\|null` | `'thin'` | Border style |
| `borderColor` | `string` | `'000000'` | Border color |
| `alternateRowColor` | `string\|null` | — | Zebra stripe color |
| `wrapText` | `boolean` | `false` | Enable text wrapping |
| `freezeHeader` | `boolean` | `false` | Freeze header row |
| `autoFilter` | `boolean` | `false` | Add filter dropdowns |
| `merges` | `Merge[]` | — | Cell merge definitions |
| `cellStyles` | `Record<string, CellStyle>` | — | Per-cell overrides |

### `drawTableFromData(wb, ws, data, options?): TableResult`

Same options as `drawTable` plus:

| Option | Type | Description |
|--------|------|-------------|
| `headers` | `string[]` | Explicit column order (default: Object.keys) |
| `headerMap` | `Record<string, string>` | Rename keys to display headers |

### `TableResult`

| Property | Type | Description |
|----------|------|-------------|
| `range` | `string` | A1-style range of entire table |
| `rowCount` | `number` | Total rows (header + data) |
| `colCount` | `number` | Number of columns |
| `firstDataRow` | `number` | 0-based first data row index |
| `lastDataRow` | `number` | 0-based last data row index |
```

**Step 2: Commit**

```bash
git add docs/guide/tables.md
git commit -m "docs: table layout guide with invoice, report, and JSON examples"
```

---

## Task 13: Lint, typecheck, full test suite

**Step 1: Run lint**

Run: `pnpm -C packages/modern-xlsx lint`
Expected: PASS (no errors)

**Step 2: Run typecheck**

Run: `pnpm -C packages/modern-xlsx typecheck`
Expected: PASS

**Step 3: Run full test suite**

Run: `pnpm -C packages/modern-xlsx test`
Expected: All tests pass (280 existing + ~25 new table tests)

**Step 4: Fix any issues found**

If lint/typecheck/tests fail, fix the issues in `table.ts` or `table.test.ts`.

**Step 5: Final commit**

```bash
git add -A
git commit -m "feat: table layout engine v0.4.0 — drawTable, drawTableFromData, styling, merges, zebra, auto-width"
```
