import { unlink } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook, readFile as xlsxReadFile } from '../src/index.js';

describe('Workbook enhancements', () => {
  it('sheetCount returns the number of sheets', () => {
    const wb = new Workbook();
    expect(wb.sheetCount).toBe(0);
    wb.addSheet('Sheet1');
    wb.addSheet('Sheet2');
    expect(wb.sheetCount).toBe(2);
  });

  it('removeSheet by name', () => {
    const wb = new Workbook();
    wb.addSheet('Alpha');
    wb.addSheet('Beta');
    wb.addSheet('Gamma');

    expect(wb.removeSheet('Beta')).toBe(true);
    expect(wb.sheetNames).toEqual(['Alpha', 'Gamma']);
    expect(wb.sheetCount).toBe(2);
  });

  it('removeSheet by index', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    wb.addSheet('C');

    expect(wb.removeSheet(0)).toBe(true);
    expect(wb.sheetNames).toEqual(['B', 'C']);
  });

  it('removeSheet returns false for non-existent', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    expect(wb.removeSheet('ZZZ')).toBe(false);
    expect(wb.removeSheet(99)).toBe(false);
  });

  it('removeNamedRange works', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.addNamedRange('Range1', 'Sheet1!$A$1');
    wb.addNamedRange('Range2', 'Sheet1!$B$1');

    expect(wb.removeNamedRange('Range1')).toBe(true);
    expect(wb.namedRanges).toHaveLength(1);
    expect(wb.namedRanges[0]?.name).toBe('Range2');
  });

  it('removeNamedRange returns false for missing', () => {
    const wb = new Workbook();
    expect(wb.removeNamedRange('nope')).toBe(false);
  });
});

describe('Worksheet.name setter', () => {
  it('renames a sheet', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('OldName');
    ws.name = 'NewName';
    expect(ws.name).toBe('NewName');
    expect(wb.sheetNames).toEqual(['NewName']);
  });
});

describe('Worksheet.columns', () => {
  it('set and get columns', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.columns = [{ min: 1, max: 1, width: 20, hidden: false, customWidth: true }];
    expect(ws.columns).toHaveLength(1);
    expect(ws.columns[0]?.width).toBe(20);
  });

  it('setColumnWidth creates column info', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.setColumnWidth(1, 15);
    expect(ws.columns).toHaveLength(1);
    expect(ws.columns[0]?.width).toBe(15);
  });

  it('setColumnWidth updates existing column', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.setColumnWidth(1, 15);
    ws.setColumnWidth(1, 25);
    expect(ws.columns).toHaveLength(1);
    expect(ws.columns[0]?.width).toBe(25);
  });

  it('columns survive round-trip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.setColumnWidth(1, 20);

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2?.columns).toHaveLength(1);
    expect(ws2?.columns[0]?.width).toBe(20);
  });
});

describe('Worksheet.mergeCells', () => {
  it('addMergeCell adds a merge range', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addMergeCell('A1:B2');
    expect(ws.mergeCells).toEqual(['A1:B2']);
  });

  it('addMergeCell prevents duplicates', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addMergeCell('A1:B2');
    ws.addMergeCell('A1:B2');
    expect(ws.mergeCells).toHaveLength(1);
  });

  it('removeMergeCell removes a range', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addMergeCell('A1:B2');
    ws.addMergeCell('C1:D2');
    expect(ws.removeMergeCell('A1:B2')).toBe(true);
    expect(ws.mergeCells).toEqual(['C1:D2']);
  });

  it('removeMergeCell returns false for missing', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.removeMergeCell('X1:X2')).toBe(false);
  });

  it('merge cells survive round-trip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'merged';
    ws.addMergeCell('A1:C1');

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2?.mergeCells).toContain('A1:C1');
  });
});

describe('Worksheet.autoFilter', () => {
  it('set and get auto filter with string shorthand', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.autoFilter = 'A1:D10';
    expect(ws.autoFilter).toEqual({ range: 'A1:D10' });
  });

  it('set and get auto filter with object', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.autoFilter = { range: 'A1:D10', filterColumns: [{ colId: 0, filters: ['Yes'] }] };
    expect(ws.autoFilter?.range).toBe('A1:D10');
    expect(ws.autoFilter?.filterColumns).toHaveLength(1);
  });

  it('clear auto filter', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.autoFilter = 'A1:D10';
    ws.autoFilter = null;
    expect(ws.autoFilter).toBeNull();
  });
});

describe('Worksheet.frozenPane', () => {
  it('set and get frozen pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.frozenPane = { rows: 1, cols: 0 };
    expect(ws.frozenPane).toEqual({ rows: 1, cols: 0 });
  });

  it('clear frozen pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.frozenPane = { rows: 1, cols: 1 };
    ws.frozenPane = null;
    expect(ws.frozenPane).toBeNull();
  });

  it('frozen pane survives round-trip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'header';
    ws.frozenPane = { rows: 1, cols: 0 };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2?.frozenPane).toEqual({ rows: 1, cols: 0 });
  });
});

describe('Worksheet.validations', () => {
  it('removeValidation removes by ref', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addValidation('A1:A10', {
      validationType: 'list',
      operator: null,
      formula1: '"Yes,No"',
      formula2: null,
      allowBlank: true,
      showErrorMessage: true,
      errorTitle: null,
      errorMessage: null,
    });
    expect(ws.removeValidation('A1:A10')).toBe(true);
    expect(ws.validations).toHaveLength(0);
  });
});

describe('Worksheet row utilities', () => {
  it('setRowHeight sets height on existing row', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.setRowHeight(1, 30);
    expect(ws.rows[0]?.height).toBe(30);
  });

  it('setRowHidden hides a row', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.setRowHidden(1, true);
    expect(ws.rows[0]?.hidden).toBe(true);
  });
});

describe('Worksheet.hyperlinks', () => {
  it('addHyperlink and get hyperlinks', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addHyperlink('A1', 'Sheet2!A1', { display: 'Go to Sheet2', tooltip: 'Click me' });
    expect(ws.hyperlinks).toHaveLength(1);
    expect(ws.hyperlinks[0]?.cellRef).toBe('A1');
    expect(ws.hyperlinks[0]?.location).toBe('Sheet2!A1');
    expect(ws.hyperlinks[0]?.display).toBe('Go to Sheet2');
  });

  it('removeHyperlink removes by cellRef', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addHyperlink('A1', '#Sheet2!A1');
    ws.addHyperlink('B1', '#Sheet3!A1');
    expect(ws.removeHyperlink('A1')).toBe(true);
    expect(ws.hyperlinks).toHaveLength(1);
    expect(ws.hyperlinks[0]?.cellRef).toBe('B1');
  });

  it('removeHyperlink returns false for missing', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.removeHyperlink('Z99')).toBe(false);
  });
});

describe('Worksheet.pageSetup', () => {
  it('set and get page setup', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.pageSetup = { orientation: 'landscape', paperSize: 9 };
    expect(ws.pageSetup?.orientation).toBe('landscape');
    expect(ws.pageSetup?.paperSize).toBe(9);
  });

  it('clear page setup', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.pageSetup = { orientation: 'portrait' };
    ws.pageSetup = null;
    expect(ws.pageSetup).toBeNull();
  });
});

describe('Worksheet.sheetProtection', () => {
  it('set and get sheet protection', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.sheetProtection = {
      sheet: true,
      objects: false,
      scenarios: false,
      formatCells: false,
      formatColumns: false,
      formatRows: false,
      insertColumns: false,
      insertRows: false,
      deleteColumns: false,
      deleteRows: false,
      sort: false,
      autoFilter: false,
    };
    expect(ws.sheetProtection?.sheet).toBe(true);
  });

  it('clear sheet protection', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.sheetProtection = {
      sheet: true,
      objects: false,
      scenarios: false,
      formatCells: false,
      formatColumns: false,
      formatRows: false,
      insertColumns: false,
      insertRows: false,
      deleteColumns: false,
      deleteRows: false,
      sort: false,
      autoFilter: false,
    };
    ws.sheetProtection = null;
    expect(ws.sheetProtection).toBeNull();
  });
});

describe('Workbook.docProperties', () => {
  it('set and get doc properties', () => {
    const wb = new Workbook();
    wb.docProperties = { title: 'Test', creator: 'modern-xlsx' };
    expect(wb.docProperties?.title).toBe('Test');
    expect(wb.docProperties?.creator).toBe('modern-xlsx');
  });

  it('clear doc properties', () => {
    const wb = new Workbook();
    wb.docProperties = { title: 'Test' };
    wb.docProperties = null;
    expect(wb.docProperties).toBeNull();
  });
});

describe('Workbook.workbookViews', () => {
  it('set and get workbook views', () => {
    const wb = new Workbook();
    wb.workbookViews = [
      {
        activeTab: 0,
        firstSheet: 0,
        showHorizontalScroll: true,
        showVerticalScroll: true,
        showSheetTabs: true,
      },
    ];
    expect(wb.workbookViews).toHaveLength(1);
    expect(wb.workbookViews[0]?.activeTab).toBe(0);
  });
});

describe('File I/O', () => {
  it('toFile and readFile round-trip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'file test';

    const tmpPath = new URL('../__tests__/tmp-test.xlsx', import.meta.url).pathname;
    // Clean Windows path if needed
    const cleanPath =
      tmpPath.startsWith('/') && process.platform === 'win32' ? tmpPath.slice(1) : tmpPath;

    await wb.toFile(cleanPath);

    const wb2 = await xlsxReadFile(cleanPath);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2?.cell('A1').value).toBe('file test');

    // Cleanup
    await unlink(cleanPath);
  });
});
