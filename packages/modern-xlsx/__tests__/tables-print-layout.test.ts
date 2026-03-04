import { describe, expect, it } from 'vitest';
import type { HeaderFooterData, OutlinePropertiesData, TableDefinitionData } from '../src/index.js';
import {
  HeaderFooterBuilder,
  readBuffer,
  TABLE_STYLES,
  VALID_TABLE_STYLES,
  Workbook,
} from '../src/index.js';

// ---------------------------------------------------------------------------
// Tables (Excel ListObjects)
// ---------------------------------------------------------------------------

describe('Excel Tables (ListObjects)', () => {
  it('adds and retrieves a table on a worksheet', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'Name';
    ws.cell('B1').value = 'Age';
    ws.cell('A2').value = 'Alice';
    ws.cell('B2').value = 30;

    const table: TableDefinitionData = {
      id: 1,
      displayName: 'People',
      ref: 'A1:B2',
      headerRowCount: 1,
      totalsRowCount: 0,
      totalsRowShown: true,
      columns: [
        { id: 1, name: 'Name' },
        { id: 2, name: 'Age' },
      ],
      styleInfo: {
        name: 'TableStyleMedium2',
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    };

    ws.addTable(table);
    expect(ws.tables).toHaveLength(1);
    expect(ws.getTable('People')).toBeDefined();
    expect(ws.getTable('People')?.ref).toBe('A1:B2');
  });

  it('removes a table by display name', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.addTable({
      id: 1,
      displayName: 'T1',
      ref: 'A1:B2',
      headerRowCount: 1,
      totalsRowCount: 0,
      totalsRowShown: true,
      columns: [
        { id: 1, name: 'Col1' },
        { id: 2, name: 'Col2' },
      ],
    });

    expect(ws.removeTable('T1')).toBe(true);
    expect(ws.tables).toHaveLength(0);
    expect(ws.removeTable('NonExistent')).toBe(false);
  });

  it('returns undefined for non-existent table', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    expect(ws.getTable('DoesNotExist')).toBeUndefined();
  });

  it('supports table with totals row and calculated columns', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sales');
    const table: TableDefinitionData = {
      id: 2,
      name: 'Sales',
      displayName: 'Sales',
      ref: 'A1:D6',
      headerRowCount: 1,
      totalsRowCount: 1,
      totalsRowShown: true,
      columns: [
        { id: 1, name: 'Product' },
        { id: 2, name: 'Qty', totalsRowFunction: 'sum' },
        { id: 3, name: 'Price', totalsRowFunction: 'average' },
        {
          id: 4,
          name: 'Total',
          totalsRowFunction: 'sum',
          calculatedColumnFormula: 'Sales[@Qty]*Sales[@Price]',
        },
      ],
      autoFilterRef: 'A1:D5',
    };

    ws.addTable(table);
    expect(ws.getTable('Sales')?.columns[3]?.calculatedColumnFormula).toBe(
      'Sales[@Qty]*Sales[@Price]',
    );
    expect(ws.getTable('Sales')?.totalsRowCount).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// Table Styles
// ---------------------------------------------------------------------------

describe('Table Styles', () => {
  it('has 60 built-in styles', () => {
    expect(TABLE_STYLES.light).toHaveLength(21);
    expect(TABLE_STYLES.medium).toHaveLength(28);
    expect(TABLE_STYLES.dark).toHaveLength(11);
    expect(VALID_TABLE_STYLES.size).toBe(60);
  });

  it('includes expected style names', () => {
    expect(VALID_TABLE_STYLES.has('TableStyleLight1')).toBe(true);
    expect(VALID_TABLE_STYLES.has('TableStyleMedium14')).toBe(true);
    expect(VALID_TABLE_STYLES.has('TableStyleDark11')).toBe(true);
    expect(VALID_TABLE_STYLES.has('TableStyleFake1')).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Headers / Footers
// ---------------------------------------------------------------------------

describe('Header/Footer', () => {
  it('sets and gets header/footer data', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    const hf: HeaderFooterData = {
      oddHeader: '&CConfidential',
      oddFooter: '&LPage &P&R&D',
    };

    ws.headerFooter = hf;
    expect(ws.headerFooter).toEqual(hf);
  });

  it('clears header/footer with null', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.headerFooter = { oddHeader: '&CTest' };
    ws.headerFooter = null;
    expect(ws.headerFooter).toBeNull();
  });
});

describe('HeaderFooterBuilder', () => {
  it('builds a basic header string', () => {
    const result = new HeaderFooterBuilder()
      .left('Left Text')
      .center('Center Text')
      .right('Right Text')
      .build();

    expect(result).toBe('&LLeft Text&CCenter Text&RRight Text');
  });

  it('uses static code helpers', () => {
    expect(HeaderFooterBuilder.pageNumber()).toBe('&P');
    expect(HeaderFooterBuilder.totalPages()).toBe('&N');
    expect(HeaderFooterBuilder.date()).toBe('&D');
    expect(HeaderFooterBuilder.time()).toBe('&T');
    expect(HeaderFooterBuilder.fileName()).toBe('&F');
    expect(HeaderFooterBuilder.sheetName()).toBe('&A');
    expect(HeaderFooterBuilder.filePath()).toBe('&Z');
  });

  it('uses font formatting helpers', () => {
    expect(HeaderFooterBuilder.bold('test')).toBe('&Btest&B');
    expect(HeaderFooterBuilder.italic('test')).toBe('&Itest&I');
    expect(HeaderFooterBuilder.underline('test')).toBe('&Utest&U');
    expect(HeaderFooterBuilder.strikethrough('test')).toBe('&Stest&S');
    expect(HeaderFooterBuilder.fontSize(12)).toBe('&12');
    expect(HeaderFooterBuilder.fontName('Arial')).toBe('&"Arial"');
    expect(HeaderFooterBuilder.color('FF0000')).toBe('&KFF0000');
  });

  it('composes a complex header/footer', () => {
    const hf = new HeaderFooterBuilder()
      .left(`Printed: ${HeaderFooterBuilder.date()}`)
      .center(HeaderFooterBuilder.bold('Confidential'))
      .right(`Page ${HeaderFooterBuilder.pageNumber()} of ${HeaderFooterBuilder.totalPages()}`)
      .build();

    expect(hf).toBe('&LPrinted: &D&C&BConfidential&B&RPage &P of &N');
  });
});

// ---------------------------------------------------------------------------
// Outline Properties
// ---------------------------------------------------------------------------

describe('Outline Properties', () => {
  it('sets and gets outline properties', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    const props: OutlinePropertiesData = {
      summaryBelow: false,
      summaryRight: false,
    };

    ws.outlineProperties = props;
    expect(ws.outlineProperties).toEqual(props);
  });

  it('defaults to null', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.outlineProperties).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// Row / Column Grouping
// ---------------------------------------------------------------------------

describe('Row Grouping', () => {
  it('groups rows with outline level', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Header';
    ws.cell('A2').value = 'Detail 1';
    ws.cell('A3').value = 'Detail 2';

    ws.groupRows(2, 3, 1);

    const rows = ws.rows;
    const row2 = rows.find((r) => r.index === 2);
    const row3 = rows.find((r) => r.index === 3);
    expect(row2?.outlineLevel).toBe(1);
    expect(row3?.outlineLevel).toBe(1);
  });

  it('ungroups rows', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.groupRows(2, 3, 2);
    ws.ungroupRows(2, 3);

    const rows = ws.rows;
    const row2 = rows.find((r) => r.index === 2);
    expect(row2?.outlineLevel).toBeNull();
    expect(row2?.collapsed).toBe(false);
  });

  it('collapses and expands rows', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Header';
    ws.cell('A2').value = 'Detail 1';
    ws.cell('A3').value = 'Detail 2';
    ws.cell('A4').value = 'Summary';

    ws.groupRows(2, 3, 1);
    ws.collapseRows(2, 3);

    const rows = ws.rows;
    expect(rows.find((r) => r.index === 2)?.hidden).toBe(true);
    expect(rows.find((r) => r.index === 3)?.hidden).toBe(true);
    expect(rows.find((r) => r.index === 4)?.collapsed).toBe(true);

    ws.expandRows(2, 3);
    expect(rows.find((r) => r.index === 2)?.hidden).toBe(false);
    expect(rows.find((r) => r.index === 3)?.hidden).toBe(false);
  });
});

describe('Column Grouping', () => {
  it('groups columns with outline level', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.groupColumns(2, 3, 1);

    const cols = ws.columns;
    const col2 = cols.find((c) => c.min === 2);
    const col3 = cols.find((c) => c.min === 3);
    expect(col2?.outlineLevel).toBe(1);
    expect(col3?.outlineLevel).toBe(1);
  });

  it('ungroups columns', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.groupColumns(2, 3, 2);
    ws.ungroupColumns(2, 3);

    const cols = ws.columns;
    const col2 = cols.find((c) => c.min === 2);
    expect(col2?.outlineLevel).toBeNull();
    expect(col2?.collapsed).toBe(false);
  });

  it('updates existing column definitions', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.setColumnWidth(2, 20);
    ws.groupColumns(2, 2, 1);

    const cols = ws.columns;
    const col2 = cols.find((c) => c.min === 2);
    expect(col2?.width).toBe(20);
    expect(col2?.outlineLevel).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// Print Titles & Print Areas
// ---------------------------------------------------------------------------

describe('Print Titles', () => {
  it('sets and gets print titles', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    wb.setPrintTitles(0, 'Sheet1!$1:$2');
    expect(wb.getPrintTitles(0)).toBe('Sheet1!$1:$2');
  });

  it('clears print titles', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    wb.setPrintTitles(0, 'Sheet1!$1:$2');
    wb.setPrintTitles(0, null);
    expect(wb.getPrintTitles(0)).toBeNull();
  });

  it('updates existing print titles', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    wb.setPrintTitles(0, 'Sheet1!$1:$2');
    wb.setPrintTitles(0, 'Sheet1!$1:$3');
    expect(wb.getPrintTitles(0)).toBe('Sheet1!$1:$3');
    // Should not have duplicates
    expect(wb.namedRanges.filter((d) => d.name === '_xlnm.Print_Titles')).toHaveLength(1);
  });
});

describe('Print Area', () => {
  it('sets and gets print area', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    wb.setPrintArea(0, 'Sheet1!$A$1:$D$50');
    expect(wb.getPrintArea(0)).toBe('Sheet1!$A$1:$D$50');
  });

  it('clears print area', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    wb.setPrintArea(0, 'Sheet1!$A$1:$D$50');
    wb.setPrintArea(0, null);
    expect(wb.getPrintArea(0)).toBeNull();
  });

  it('handles multiple sheets independently', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.addSheet('Sheet2');

    wb.setPrintArea(0, 'Sheet1!$A$1:$D$50');
    wb.setPrintArea(1, 'Sheet2!$A$1:$F$100');

    expect(wb.getPrintArea(0)).toBe('Sheet1!$A$1:$D$50');
    expect(wb.getPrintArea(1)).toBe('Sheet2!$A$1:$F$100');
  });
});

// ---------------------------------------------------------------------------
// Roundtrip tests (require WASM)
// ---------------------------------------------------------------------------

describe('Roundtrip: Header/Footer', () => {
  it('roundtrips header/footer through WASM', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.headerFooter = {
      oddHeader: '&CConfidential',
      oddFooter: '&LPage &P&RDate: &D',
      differentOddEven: false,
      differentFirst: false,
    };

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Sheet1');

    expect(ws2?.headerFooter?.oddHeader).toBe('&CConfidential');
    expect(ws2?.headerFooter?.oddFooter).toBe('&LPage &P&RDate: &D');
  });
});

describe('Roundtrip: Outline Properties', () => {
  it('roundtrips outline properties through WASM', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.outlineProperties = { summaryBelow: false, summaryRight: true };

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Sheet1');

    expect(ws2?.outlineProperties?.summaryBelow).toBe(false);
    expect(ws2?.outlineProperties?.summaryRight).toBe(true);
  });
});

describe('Roundtrip: Row/Column Grouping', () => {
  it('roundtrips row outline levels through WASM', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Header';
    ws.cell('A2').value = 'Detail 1';
    ws.cell('A3').value = 'Detail 2';

    ws.groupRows(2, 3, 2);

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Sheet1');
    const rows = ws2?.rows ?? [];
    const row2 = rows.find((r) => r.index === 2);
    expect(row2?.outlineLevel).toBe(2);
  });

  it('roundtrips column outline levels through WASM', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';

    ws.groupColumns(2, 3, 1);

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Sheet1');
    const cols = ws2?.columns ?? [];
    const col2 = cols.find((c) => c.min === 2);
    expect(col2?.outlineLevel).toBe(1);
  });
});
