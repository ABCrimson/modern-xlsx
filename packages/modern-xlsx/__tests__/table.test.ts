import { assert, describe, expect, it } from 'vitest';
import type { TableColumnData, TableDefinitionData } from '../src/index.js';
import {
  drawTable,
  drawTableFromData,
  ModernXlsxError,
  readBuffer,
  Workbook,
} from '../src/index.js';

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

    expect(ws.cell('A1').value).toBe('Name');
    expect(ws.cell('B1').value).toBe('Age');
    expect(ws.cell('C1').value).toBe('City');

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

    const headerCell = ws.cell('A1');
    expect(headerCell.styleIndex).not.toBeNull();
    expect(headerCell.styleIndex).toBeGreaterThan(0);

    const headerStyleIndex = headerCell.styleIndex;
    if (headerStyleIndex == null) throw new Error('headerCell styleIndex not found');
    const xf = wb.styles.cellXfs[headerStyleIndex];
    assert.isDefined(xf);
    expect(wb.styles.fonts[xf.fontId]?.bold).toBe(true);
    expect(wb.styles.fills[xf.fillId]?.patternType).toBe('solid');
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
    const cellStyleIndex = cell.styleIndex;
    if (cellStyleIndex == null) throw new Error('cell styleIndex not found');
    const xf = wb.styles.cellXfs[cellStyleIndex];
    assert.isDefined(xf);
    const border = wb.styles.borders[xf.borderId];
    assert.isDefined(border);
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
      rows: [
        ['Alice', 30],
        ['Bob', 25],
      ],
    });

    expect(result.range).toBe('A1:B3');
    expect(result.rowCount).toBe(3);
    expect(result.colCount).toBe(2);
  });
});

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

    const cols = ws.columns;
    expect(cols.length).toBeGreaterThanOrEqual(3);
    const colC = cols.find((c) => c.min === 3);
    assert.isDefined(colC);
    expect(colC.width).toBeGreaterThan(20);

    const colA = cols.find((c) => c.min === 1);
    expect(colA).toBeDefined();
    expect(colA?.width).toBeLessThan(colC.width);
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
    expect(col?.width).toBeGreaterThan(14);
  });
});

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
        '1,1': { font: { color: 'FF0000', bold: true } },
      },
    });

    const cell = ws.cell('B3');
    expect(cell.styleIndex).not.toBeNull();
    const cellStyleIndex = cell.styleIndex;
    if (cellStyleIndex == null) throw new Error('cell styleIndex not found');
    const xf = wb.styles.cellXfs[cellStyleIndex];
    assert.isDefined(xf);
    const font = wb.styles.fonts[xf.fontId];
    assert.isDefined(font);
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
    const cellStyleIndex = cell.styleIndex;
    if (cellStyleIndex == null) throw new Error('cell styleIndex not found');
    const xf = wb.styles.cellXfs[cellStyleIndex];
    assert.isDefined(xf);
    const border = wb.styles.borders[xf.borderId];
    assert.isDefined(border);
    expect(border.top?.style).toBe('thin');
    const fill = wb.styles.fills[xf.fillId];
    assert.isDefined(fill);
    expect(fill.fgColor).toBe('FFFF00');
  });
});

describe('merge cells', () => {
  it('spans columns for a subtotal merge', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Product', 'Q1', 'Q2', 'Q3'],
      rows: [
        ['Widget', 100, 200, 150],
        ['Subtotal', null, null, 450],
      ],
      merges: [{ row: 1, col: 0, colSpan: 3 }],
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
      merges: [{ row: 0, col: 0, rowSpan: 3 }],
    });

    expect(ws.mergeCells).toContain('A2:A4');
    expect(ws.cell('A2').value).toBe('Electronics');
  });
});

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

    // Even rows (0, 2) should NOT have fill
    const evenStyleIndex = ws.cell('A2').styleIndex;
    if (evenStyleIndex == null) throw new Error('A2 styleIndex not found');
    const evenXf = wb.styles.cellXfs[evenStyleIndex];
    assert.isDefined(evenXf);
    const evenFill = wb.styles.fills[evenXf.fillId];
    assert.isDefined(evenFill);
    expect(evenFill.patternType).toBe('none');

    // Odd rows (1, 3) should have the alternate color
    const oddStyleIndex = ws.cell('A3').styleIndex;
    if (oddStyleIndex == null) throw new Error('A3 styleIndex not found');
    const oddXf = wb.styles.cellXfs[oddStyleIndex];
    assert.isDefined(oddXf);
    const oddFill = wb.styles.fills[oddXf.fillId];
    assert.isDefined(oddFill);
    expect(oddFill.patternType).toBe('solid');
    expect(oddFill.fgColor).toBe('E8EDF3');

    const even2StyleIndex = ws.cell('A4').styleIndex;
    if (even2StyleIndex == null) throw new Error('A4 styleIndex not found');
    const even2Xf = wb.styles.cellXfs[even2StyleIndex];
    assert.isDefined(even2Xf);
    const even2Fill = wb.styles.fills[even2Xf.fillId];
    assert.isDefined(even2Fill);
    expect(even2Fill.patternType).toBe('none');

    const odd2StyleIndex = ws.cell('A5').styleIndex;
    if (odd2StyleIndex == null) throw new Error('A5 styleIndex not found');
    const odd2Xf = wb.styles.cellXfs[odd2StyleIndex];
    assert.isDefined(odd2Xf);
    const odd2Fill = wb.styles.fills[odd2Xf.fillId];
    assert.isDefined(odd2Fill);
    expect(odd2Fill.fgColor).toBe('E8EDF3');
  });
});

describe('cell alignment', () => {
  it('applies per-column horizontal alignment', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Left', 'Center', 'Right'],
      rows: [['a', 'b', 'c']],
      columns: [{ align: 'left' }, { align: 'center' }, { align: 'right' }],
    });

    const leftStyleIndex = ws.cell('A2').styleIndex;
    if (leftStyleIndex == null) throw new Error('A2 styleIndex not found');
    const leftXf = wb.styles.cellXfs[leftStyleIndex];
    assert.isDefined(leftXf);
    expect(leftXf.alignment?.horizontal).toBe('left');

    const centerStyleIndex = ws.cell('B2').styleIndex;
    if (centerStyleIndex == null) throw new Error('B2 styleIndex not found');
    const centerXf = wb.styles.cellXfs[centerStyleIndex];
    assert.isDefined(centerXf);
    expect(centerXf.alignment?.horizontal).toBe('center');

    const rightStyleIndex = ws.cell('C2').styleIndex;
    if (rightStyleIndex == null) throw new Error('C2 styleIndex not found');
    const rightXf = wb.styles.cellXfs[rightStyleIndex];
    assert.isDefined(rightXf);
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

    const headerStyleIndex = ws.cell('A1').styleIndex;
    if (headerStyleIndex == null) throw new Error('A1 styleIndex not found');
    const headerXf = wb.styles.cellXfs[headerStyleIndex];
    assert.isDefined(headerXf);
    expect(headerXf.alignment?.vertical).toBe('center');

    const bodyStyleIndex = ws.cell('A2').styleIndex;
    if (bodyStyleIndex == null) throw new Error('A2 styleIndex not found');
    const bodyXf = wb.styles.cellXfs[bodyStyleIndex];
    assert.isDefined(bodyXf);
    expect(bodyXf.alignment?.vertical).toBe('center');
  });
});

describe('content wrapping', () => {
  it('enables wrapText on body cells', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    drawTable(wb, ws, {
      headers: ['Description'],
      rows: [['A very long description that should wrap within the cell']],
      wrapText: true,
    });

    const wrapStyleIndex = ws.cell('A2').styleIndex;
    if (wrapStyleIndex == null) throw new Error('A2 styleIndex not found');
    const bodyXf = wb.styles.cellXfs[wrapStyleIndex];
    assert.isDefined(bodyXf);
    expect(bodyXf.alignment?.wrapText).toBe(true);
  });
});

describe('nested tables', () => {
  it('draws a nested table below the parent table', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    const result = drawTable(wb, ws, {
      headers: ['Order', 'Total'],
      rows: [['ORD-001', 500]],
    });

    const nested = drawTable(wb, ws, {
      headers: ['Item', 'Qty', 'Price'],
      rows: [
        ['Widget', 2, 100],
        ['Gadget', 3, 100],
      ],
      origin: `A${result.rowCount + 2}`,
    });

    expect(ws.cell('A1').value).toBe('Order');
    expect(ws.cell('A2').value).toBe('ORD-001');
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
      rows: 1,
      cols: 0,
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

    expect(ws.autoFilter).toEqual({ range: 'A1:C2' });
  });
});

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

  it('applies all drawTable options (zebra, autoWidth, freezeHeader)', () => {
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

    const oddStyleIdx = ws.cell('A3').styleIndex;
    if (oddStyleIdx == null) throw new Error('A3 styleIndex not found');
    const oddXf = wb.styles.cellXfs[oddStyleIdx];
    assert.isDefined(oddXf);
    const oddFill = wb.styles.fills[oddXf.fillId];
    assert.isDefined(oddFill);
    expect(oddFill.fgColor).toBe('F0F0F0');

    expect(ws.frozenPane).toEqual({ rows: 1, cols: 0 });

    expect(ws.columns.length).toBeGreaterThan(0);
  });

  it('handles empty data array', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    const result = drawTableFromData(wb, ws, []);
    expect(result).toEqual({
      range: '',
      rowCount: 0,
      colCount: 0,
      firstDataRow: 1,
      lastDataRow: 0,
    });
    expect(ws.rows).toHaveLength(0);
  });

  it('writes only the header row for an empty array with headers and headerMap', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const result = drawTableFromData(wb, ws, [], { headers: ['a'], headerMap: { a: 'A!' } });
    expect(ws.cell('A1').value).toBe('A!');
    expect(result.rowCount).toBe(1);
    expect(result.range).toBe('A1:A1');
  });

  it('maps Date and bigint cells', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    drawTableFromData(wb, ws, [
      { d: new Date(Date.UTC(2026, 0, 2)), n: 12n, big: 2n ** 70n, bad: new Date(Number.NaN) },
    ]);
    expect(ws.cell('A2').value).toBe(46024);
    expect(ws.cell('B2').value).toBe(12);
    expect(ws.cell('C2').value).toBe(String(2n ** 70n));
    expect(ws.cell('D2').value).toBeNull();
    const xf = wb.styles.cellXfs[ws.cell('A2').styleIndex ?? 0];
    expect(wb.styles.numFmts.find((f) => f.id === xf?.numFmtId)?.formatCode).toBe('yyyy-mm-dd');
  });
});

describe('drawTable layout contract', () => {
  it('throws RangeError for empty headers', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(() => drawTable(wb, ws, { headers: [], rows: [] })).toThrow(RangeError);
    expect(() => drawTableFromData(wb, ws, [])).not.toThrow();
  });

  it('honours a header override and sizes the column from it', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    drawTable(wb, ws, {
      headers: ['n'],
      rows: [['x']],
      columns: [{ header: 'Full Name Of The Person' }],
      autoWidth: true,
    });
    expect(ws.cell('A1').value).toBe('Full Name Of The Person');
    const w = ws.columns.find((c) => c.min === 1)?.width ?? 0;
    expect(w).toBeGreaterThanOrEqual(Math.ceil('Full Name Of The Person'.length * 1.2) + 2);
  });

  it('freezes below the header of an offset origin and spans the autoFilter over the body', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const r = drawTable(wb, ws, {
      headers: ['a', 'b'],
      rows: [['1', '2']],
      origin: 'B3',
      freezeHeader: true,
      autoFilter: true,
    });
    expect(r.range).toBe('B3:C4');
    expect(ws.frozenPane).toEqual({ rows: 3, cols: 0 });
    expect(ws.autoFilter).toEqual({ range: 'B3:C4' });
    const ws2 = wb.addSheet('S2');
    drawTable(wb, ws2, { headers: ['a'], rows: [], origin: 'C5', freezeHeader: true });
    expect(ws2.frozenPane).toEqual({ rows: 5, cols: 0 });
  });

  it('lets an explicit width win over autoWidth', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    drawTable(wb, ws, {
      headers: ['a', 'b'],
      rows: [['long text value here', 'y']],
      autoWidth: true,
      columnWidths: [40, 0],
      columns: [{}, { width: 25 }],
    });
    const widthOf = (col: number) => ws.columns.find((c) => c.min === col)?.width;
    expect(widthOf(1)).toBe(40);
    expect(widthOf(2)).toBe(25);
  });
});

describe('round-trip', () => {
  it('writes a basic table and reads it back', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Invoice');

    drawTable(wb, ws, {
      headers: ['Item', 'Qty'],
      rows: [
        ['Widget', 10],
        ['Gadget', 5],
      ],
    });

    const buffer = await wb.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Invoice');
    expect(ws2).toBeDefined();
    expect(ws2?.cell('A1').value).toBe('Item');
  });

  it('writes a fully styled table and reads it back', async () => {
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
      columns: [
        { align: 'left' },
        { align: 'center' },
        { align: 'right', numberFormat: '#,##0.00' },
        { align: 'right', numberFormat: '#,##0.00' },
      ],
    });

    const buffer = await wb.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Invoice');
    expect(ws2).toBeDefined();

    expect(ws2?.cell('A1').value).toBe('Item');
    expect(Number(ws2?.cell('B2').value)).toBe(10);
    expect(ws2?.frozenPane).toEqual({ rows: 1, cols: 0 });
  });

  it('drawTableFromData ignores inherited Object.prototype keys', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    drawTableFromData(wb, ws, [{ constructor: 1, a: 2 }, { a: 3 }], { headerMap: { a: 'A' } });
    expect(ws.cell('A1').value).toBe('constructor');
    expect(ws.cell('B1').value).toBe('A');
    expect(ws.cell('A2').value).toBe(1);
    expect(ws.cell('B2').value).toBe(2);
    expect(ws.cell('A3').value).toBeNull();
    expect(ws.cell('B3').value).toBe(3);
  });

  it('drawTableFromData reads a getter on the record class prototype', () => {
    class Row {
      qty = 2;
      price = 5;
      get total(): number {
        return this.qty * this.price;
      }
    }
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    drawTableFromData(wb, ws, [new Row() as unknown as Record<string, unknown>], {
      headers: ['qty', 'price', 'total'],
    });
    expect(ws.cell('C2').value).toBe(10);
  });
});

/*
 * The `@` this-row shorthand in a table definition, as Excel 365 (Windows,
 * build 20527) proved it on 2026-11-12: one workbook per form, written by
 * modern-xlsx with the form stored as given and the column cells in the file
 * form, opened through COM with alerts off. Excel refused every form with
 * `[@` in `<calculatedColumnFormula>` and in `<totalsRowFormula>`, and opened
 * the others. Excel itself writes `CalcT[[#This Row],[Qty]]` for `[@Qty]` and
 * `CalcT[#This Row]` for `[@]`.
 */

/** [formula, first shorthand reference, the file form Excel writes for it] Excel refused as a calculatedColumnFormula. */
const CALC_REFUSED: [string, string, string][] = [
  ['[@Qty]*[@Price]', '[@Qty]', 'CalcT[[#This Row],[Qty]]'],
  ['[@[Unit Cost]]*2', '[@[Unit Cost]]', 'CalcT[[#This Row],[Unit Cost]]'],
  ['CalcT[@Qty]*2', 'CalcT[@Qty]', 'CalcT[[#This Row],[Qty]]'],
  ['CalcT[@[Unit Cost]]*2', 'CalcT[@[Unit Cost]]', 'CalcT[[#This Row],[Unit Cost]]'],
  ['COLUMNS([@])', '[@]', 'CalcT[#This Row]'],
  ['COLUMNS(CalcT[@])', 'CalcT[@]', 'CalcT[#This Row]'],
  ['[@qty]*2', '[@qty]', 'CalcT[[#This Row],[qty]]'],
  ['"[@x]"&[@Qty]', '[@Qty]', 'CalcT[[#This Row],[Qty]]'],
];

/** calculatedColumnFormula forms Excel opened. */
const CALC_ACCEPTED: string[] = [
  'SUM([Qty])',
  '[[#This Row],[Qty]]*2',
  'CalcT[[#This Row],[Qty]]*2',
  'CalcT[[#This Row],[Qty]]*CalcT[[#This Row],[Price]]',
  'SUM(CalcT[Qty])',
  '"[@Qty]"&CalcT[[#This Row],[Qty]]',
  'calct[[#this row],[qty]]*2',
  'COLUMNS([#This Row])',
  'COLUMNS(CalcT[#This Row])',
  '=CalcT[[#This Row],[Qty]]*CalcT[[#This Row],[Price]]',
];

/** [formula, first shorthand reference, the file form] Excel refused as a totalsRowFormula. */
const TOTALS_REFUSED: [string, string, string][] = [
  ['SUM([@Qty])', '[@Qty]', 'CalcT[[#This Row],[Qty]]'],
  ['SUM(CalcT[@Qty])', 'CalcT[@Qty]', 'CalcT[[#This Row],[Qty]]'],
];

/** totalsRowFormula forms Excel opened. */
const TOTALS_ACCEPTED: string[] = [
  'SUM([Qty])',
  'SUM(CalcT[Qty])',
  'SUM([[#This Row],[Qty]])',
  'CalcT[[#Totals],[Price]]*2',
  '[[#Totals],[Price]]*2',
  'SUM([[#Data],[Qty]])',
];

/** Table CalcT over A1:D4 (Qty, Price, Unit Cost, Total) with a totals row; `qty` and `total` extend those columns. */
function calcTable(
  qty: Partial<TableColumnData>,
  total: Partial<TableColumnData>,
): TableDefinitionData {
  return {
    id: 1,
    displayName: 'CalcT',
    name: 'CalcT',
    ref: 'A1:D4',
    headerRowCount: 1,
    totalsRowCount: 1,
    totalsRowShown: true,
    autoFilterRef: 'A1:D3',
    columns: [
      { id: 1, name: 'Qty', ...qty },
      { id: 2, name: 'Price' },
      { id: 3, name: 'Unit Cost' },
      { id: 4, name: 'Total', ...total },
    ],
  };
}

/** A sheet `S` with CalcT's header and two data rows. */
function calcSheet(wb: Workbook) {
  const ws = wb.addSheet('S');
  ['Qty', 'Price', 'Unit Cost', 'Total'].forEach((h, i) => {
    ws.cell(`${'ABCD'[i]}1`).value = h;
  });
  ws.cell('A2').value = 2;
  ws.cell('A3').value = 4;
  return ws;
}

/** The error issues of a raw model holding `table` on sheet `S`, as read from a file. */
function rawTableErrors(table: TableDefinitionData) {
  const wb = new Workbook();
  calcSheet(wb);
  const sheet = wb.toJSON().sheets[0];
  if (!sheet) throw new Error('missing sheet');
  sheet.worksheet.tables = [table];
  return wb.validate().issues.filter((issue) => issue.severity === 'error');
}

describe('addTable and the @ this-row shorthand', () => {
  it.each(CALC_REFUSED)(
    'refuses the calculatedColumnFormula %j and names %s and %s',
    (formula, ref, fileForm) => {
      const wb = new Workbook();
      const ws = calcSheet(wb);
      let error: unknown;
      try {
        ws.addTable(calcTable({}, { calculatedColumnFormula: formula }));
      } catch (e) {
        error = e;
      }
      assert(error instanceof ModernXlsxError, 'addTable did not throw');
      expect(error.code).toBe('INVALID_ARGUMENT');
      expect(error.message).toContain('calculatedColumnFormula of column "Total" in table "CalcT"');
      expect(error.message).toContain(`@ shorthand ${ref}`);
      expect(error.message).toContain(`write ${fileForm} instead`);
      expect(ws.tables).toHaveLength(0);
    },
  );

  it.each(CALC_REFUSED)(
    'validate() reports the calculatedColumnFormula %j in a raw model',
    (formula, ref, fileForm) => {
      const errors = rawTableErrors(calcTable({}, { calculatedColumnFormula: formula }));
      expect(errors).toHaveLength(1);
      expect(errors[0]?.category).toBe('structure');
      expect(errors[0]?.location).toBe('S!table(CalcT)');
      expect(errors[0]?.message).toContain(`@ shorthand ${ref}`);
      expect(errors[0]?.message).toContain(`write ${fileForm} instead`);
    },
  );

  it.each(CALC_ACCEPTED)('takes the calculatedColumnFormula %j as given', (formula) => {
    const wb = new Workbook();
    const ws = calcSheet(wb);
    ws.addTable(calcTable({}, { calculatedColumnFormula: formula }));
    expect(ws.tables[0]?.columns[3]?.calculatedColumnFormula).toBe(formula);
    expect(wb.validate().errorCount).toBe(0);
    expect(rawTableErrors(calcTable({}, { calculatedColumnFormula: formula }))).toEqual([]);
  });

  it.each(TOTALS_REFUSED)(
    'refuses the totalsRowFormula %j and writes no totals cell',
    (formula, ref, fileForm) => {
      const wb = new Workbook();
      const ws = calcSheet(wb);
      const column = { totalsRowFunction: 'custom' as const, totalsRowFormula: formula };
      expect(() => ws.addTable(calcTable(column, {}))).toThrow(
        `totalsRowFormula of column "Qty" in table "CalcT": Excel refuses the @ shorthand ${ref} in a table definition; write ${fileForm} instead`,
      );
      expect(ws.tables).toHaveLength(0);
      expect(ws.cell('A4').formula).toBeNull();
      // The raw model has no totals cell either, which validate() also reports.
      const errors = rawTableErrors(calcTable(column, {})).filter((e) =>
        e.message.includes('@ shorthand'),
      );
      expect(errors.map((e) => e.message)).toEqual([
        expect.stringContaining(`write ${fileForm} instead`),
      ]);
    },
  );

  it.each(TOTALS_ACCEPTED)('takes the totalsRowFormula %j', (formula) => {
    const wb = new Workbook();
    const ws = calcSheet(wb);
    ws.addTable(calcTable({ totalsRowFunction: 'custom', totalsRowFormula: formula }, {}));
    expect(ws.cell('A4').formula).toBe(formula);
    expect(wb.validate().errorCount).toBe(0);
  });

  it('scans a long formula of string literals in one pass', () => {
    const formula = `${'"[@x]"&'.repeat(20_000)}CalcT[[#This Row],[Qty]]`;
    const wb = new Workbook();
    const ws = calcSheet(wb);
    ws.addTable(calcTable({}, { calculatedColumnFormula: formula }));
    expect(wb.validate().errorCount).toBe(0);
  });

  /** Table AtT over F1:G4 (`@Home`, `Share`) with the formulas Excel 365 saved for `Share`. */
  function atHomeTable(share: Partial<TableColumnData> = {}): TableDefinitionData {
    return {
      id: 2,
      displayName: 'AtT',
      name: 'AtT',
      ref: 'F1:G4',
      headerRowCount: 1,
      totalsRowCount: 1,
      totalsRowShown: true,
      autoFilterRef: 'F1:G3',
      columns: [
        { id: 1, name: '@Home' },
        {
          id: 2,
          name: 'Share',
          calculatedColumnFormula: 'AtT[[#This Row],[@Home]]/SUM(AtT[@Home])',
          totalsRowFunction: 'custom',
          totalsRowFormula: 'SUM(AtT[@Home])*2',
          ...share,
        },
      ],
    };
  }

  /** Sheet `S` with CalcT's cells and AtT's header and data cells. */
  function atHomeSheet(wb: Workbook) {
    const ws = calcSheet(wb);
    ws.cell('F1').value = '@Home';
    ws.cell('G1').value = 'Share';
    ws.cell('F2').value = 1;
    ws.cell('F3').value = 3;
    return ws;
  }

  it('takes AtT[@Home] as the column @Home, as Excel writes it', () => {
    const wb = new Workbook();
    const ws = atHomeSheet(wb);
    ws.addTable(atHomeTable());
    expect(ws.tables[0]?.columns[1]?.calculatedColumnFormula).toBe(
      'AtT[[#This Row],[@Home]]/SUM(AtT[@Home])',
    );
    expect(ws.cell('G4').formula).toBe('SUM(AtT[@Home])*2');
    expect(wb.validate().errorCount).toBe(0);
    // Without the table name, in another case, and from another table.
    const other = new Workbook();
    const ws2 = atHomeSheet(other);
    ws2.addTable(atHomeTable({ calculatedColumnFormula: '[@home]*2' }));
    ws2.addTable(
      calcTable({}, { calculatedColumnFormula: 'CalcT[[#This Row],[Qty]]*SUM(att[@HOME])' }),
    );
    expect(ws2.tables).toHaveLength(2);
    expect(other.validate().errorCount).toBe(0);
  });

  it.each([
    ['[@Home]*2', '[@Home]', 'CalcT[[#This Row],[Home]]'],
    ['CalcT[@Home]*2', 'CalcT[@Home]', 'CalcT[[#This Row],[Home]]'],
    ['Nope[@Home]*2', 'Nope[@Home]', 'Nope[[#This Row],[Home]]'],
  ])('still refuses %j in a table with no @Home column', (formula, ref, fileForm) => {
    const wb = new Workbook();
    const ws = atHomeSheet(wb);
    ws.addTable(atHomeTable());
    expect(() => ws.addTable(calcTable({}, { calculatedColumnFormula: formula }))).toThrow(
      `calculatedColumnFormula of column "Total" in table "CalcT": Excel refuses the @ shorthand ${ref} in a table definition; write ${fileForm} instead`,
    );
    expect(ws.tables).toHaveLength(1);
  });

  it('keeps a table read from a file as it is', async () => {
    const formula = '[@Qty]*2';
    const wb = new Workbook();
    calcSheet(wb);
    const sheet = wb.toJSON().sheets[0];
    if (!sheet) throw new Error('missing sheet');
    sheet.worksheet.tables = [calcTable({}, { calculatedColumnFormula: formula })];
    const back = await readBuffer(await wb.toBuffer());
    expect(back.getSheet('S')?.getTable('CalcT')?.columns[3]?.calculatedColumnFormula).toBe(
      formula,
    );
  });
});
