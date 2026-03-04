import { describe, expect, it } from 'vitest';
import { drawTable, drawTableFromData, readBuffer, Workbook } from '../src/index.js';

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
    const cellStyleIndex = cell.styleIndex;
    if (cellStyleIndex == null) throw new Error('cell styleIndex not found');
    const xf = wb.styles.cellXfs[cellStyleIndex];
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
    expect(colC).toBeDefined();
    expect(colC?.width).toBeGreaterThan(20);

    const colA = cols.find((c) => c.min === 1);
    expect(colA).toBeDefined();
    expect(colA?.width).toBeLessThan(colC?.width);
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
    const cellStyleIndex = cell.styleIndex;
    if (cellStyleIndex == null) throw new Error('cell styleIndex not found');
    const xf = wb.styles.cellXfs[cellStyleIndex];
    const border = wb.styles.borders[xf.borderId];
    expect(border.top?.style).toBe('thin');
    const fill = wb.styles.fills[xf.fillId];
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
    const evenFill = wb.styles.fills[evenXf.fillId];
    expect(evenFill.patternType).toBe('none');

    // Odd rows (1, 3) should have the alternate color
    const oddStyleIndex = ws.cell('A3').styleIndex;
    if (oddStyleIndex == null) throw new Error('A3 styleIndex not found');
    const oddXf = wb.styles.cellXfs[oddStyleIndex];
    const oddFill = wb.styles.fills[oddXf.fillId];
    expect(oddFill.patternType).toBe('solid');
    expect(oddFill.fgColor).toBe('E8EDF3');

    const even2StyleIndex = ws.cell('A4').styleIndex;
    if (even2StyleIndex == null) throw new Error('A4 styleIndex not found');
    const even2Xf = wb.styles.cellXfs[even2StyleIndex];
    const even2Fill = wb.styles.fills[even2Xf.fillId];
    expect(even2Fill.patternType).toBe('none');

    const odd2StyleIndex = ws.cell('A5').styleIndex;
    if (odd2StyleIndex == null) throw new Error('A5 styleIndex not found');
    const odd2Xf = wb.styles.cellXfs[odd2StyleIndex];
    const odd2Fill = wb.styles.fills[odd2Xf.fillId];
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
    expect(leftXf.alignment?.horizontal).toBe('left');

    const centerStyleIndex = ws.cell('B2').styleIndex;
    if (centerStyleIndex == null) throw new Error('B2 styleIndex not found');
    const centerXf = wb.styles.cellXfs[centerStyleIndex];
    expect(centerXf.alignment?.horizontal).toBe('center');

    const rightStyleIndex = ws.cell('C2').styleIndex;
    if (rightStyleIndex == null) throw new Error('C2 styleIndex not found');
    const rightXf = wb.styles.cellXfs[rightStyleIndex];
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
    expect(headerXf.alignment?.vertical).toBe('center');

    const bodyStyleIndex = ws.cell('A2').styleIndex;
    if (bodyStyleIndex == null) throw new Error('A2 styleIndex not found');
    const bodyXf = wb.styles.cellXfs[bodyStyleIndex];
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

    expect(ws.autoFilter).toEqual({ range: 'A1:C1' });
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
    const oddFill = wb.styles.fills[oddXf.fillId];
    expect(oddFill.fgColor).toBe('F0F0F0');

    expect(ws.frozenPane).toBeDefined();
    expect(ws.frozenPane?.rows).toBe(1);

    expect(ws.columns.length).toBeGreaterThan(0);
  });

  it('handles empty data array', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    const result = drawTableFromData(wb, ws, []);
    expect(result.rowCount).toBe(1);
    expect(result.colCount).toBe(0);
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
    expect(ws2?.frozenPane).toBeDefined();
  });
});
