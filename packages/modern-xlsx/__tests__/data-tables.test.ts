import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('Data Tables', () => {
  it('data table formula attributes roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    const data = wb.toJSON();
    const sheet = data.sheets[0];
    if (!sheet) throw new Error('missing sheet');
    sheet.worksheet.rows = [
      {
        index: 1,
        cells: [
          {
            reference: 'B1',
            cellType: 'number',
            styleIndex: null,
            value: '10',
            formula: null,
          },
        ],
        height: null,
        hidden: false,
      },
      {
        index: 2,
        cells: [
          {
            reference: 'A2',
            cellType: 'number',
            styleIndex: null,
            value: '20',
            formula: null,
          },
          {
            reference: 'B2',
            cellType: 'formulaStr',
            styleIndex: null,
            value: '42',
            formula: '',
            formulaType: 'dataTable',
            formulaR1: 'B1',
            formulaR2: 'A2',
          },
        ],
        height: null,
        hidden: false,
      },
    ];

    const wb2 = new Workbook(data);
    const buffer = await wb2.toBuffer();
    const wb3 = await readBuffer(buffer);

    const readData = wb3.toJSON();
    const row2 = readData.sheets[0]?.worksheet.rows.find((r) => r.index === 2);
    expect(row2).toBeDefined();

    const cell = row2?.cells.find((c) => c.reference === 'B2');
    expect(cell).toBeDefined();
    expect(cell?.formulaType).toBe('dataTable');
    expect(cell?.formulaR1).toBe('B1');
    expect(cell?.formulaR2).toBe('A2');
  });

  it('2D data table roundtrip with dtr flags', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    const data = wb.toJSON();
    const sheet = data.sheets[0];
    if (!sheet) throw new Error('missing sheet');
    sheet.worksheet.rows = [
      {
        index: 2,
        cells: [
          {
            reference: 'B2',
            cellType: 'formulaStr',
            styleIndex: null,
            value: '99',
            formula: '',
            formulaType: 'dataTable',
            formulaR1: 'B1',
            formulaR2: 'A2',
            formulaDt2d: true,
            formulaDtr1: true,
          },
        ],
        height: null,
        hidden: false,
      },
    ];

    const wb2 = new Workbook(data);
    const buffer = await wb2.toBuffer();
    const wb3 = await readBuffer(buffer);

    const readData = wb3.toJSON();
    const row2 = readData.sheets[0]?.worksheet.rows.find((r) => r.index === 2);
    const cell = row2?.cells.find((c) => c.reference === 'B2');
    expect(cell).toBeDefined();
    expect(cell?.formulaDt2d).toBe(true);
    expect(cell?.formulaDtr1).toBe(true);
    // formulaDtr2 was not set — should be absent or falsy
    expect(cell?.formulaDtr2).toBeFalsy();
  });

  it('normal formula has no data table attributes', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').formula = 'SUM(B1:B10)';
    ws.cell('A1').value = '55';

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    const readData = wb2.toJSON();
    const row1 = readData.sheets[0]?.worksheet.rows.find((r) => r.index === 1);
    const a1 = row1?.cells.find((c) => c.reference === 'A1');
    expect(a1).toBeDefined();
    expect(a1?.formula).toBe('SUM(B1:B10)');
    // Normal formulas should not have data table attributes
    expect(a1?.formulaR1).toBeFalsy();
    expect(a1?.formulaR2).toBeFalsy();
    expect(a1?.formulaDt2d).toBeFalsy();
  });

  it('data table with only r1 (1D table)', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    const data = wb.toJSON();
    const sheet = data.sheets[0];
    if (!sheet) throw new Error('missing sheet');
    sheet.worksheet.rows = [
      {
        index: 1,
        cells: [
          {
            reference: 'A1',
            cellType: 'formulaStr',
            styleIndex: null,
            value: '7',
            formula: '',
            formulaType: 'dataTable',
            formulaR1: 'C1',
          },
        ],
        height: null,
        hidden: false,
      },
    ];

    const wb2 = new Workbook(data);
    const buffer = await wb2.toBuffer();
    const wb3 = await readBuffer(buffer);

    const readData = wb3.toJSON();
    const row1 = readData.sheets[0]?.worksheet.rows.find((r) => r.index === 1);
    const cell = row1?.cells.find((c) => c.reference === 'A1');
    expect(cell).toBeDefined();
    expect(cell?.formulaType).toBe('dataTable');
    expect(cell?.formulaR1).toBe('C1');
    // r2 was not set
    expect(cell?.formulaR2).toBeFalsy();
    expect(cell?.formulaDt2d).toBeFalsy();
  });

  it('preserved entries survive roundtrip (external links)', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'data';

    const data = wb.toJSON();
    data.preservedEntries = {
      'xl/externalLinks/externalLink1.xml': Array.from(new TextEncoder().encode('<externalLink/>')),
    };

    const wb2 = new Workbook(data);
    const buffer = await wb2.toBuffer();
    const wb3 = await readBuffer(buffer);

    const readData = wb3.toJSON();
    expect(readData.preservedEntries).toBeDefined();
    const entry = readData.preservedEntries?.['xl/externalLinks/externalLink1.xml'];
    expect(entry).toBeDefined();
    if (!entry) throw new Error('missing entry');
    const decoded = new TextDecoder().decode(new Uint8Array(entry));
    expect(decoded).toBe('<externalLink/>');
  });

  it('preserved entries survive roundtrip (custom XML)', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'data';

    const data = wb.toJSON();
    const customXml = '<?xml version="1.0"?><customData><field>value</field></customData>';
    data.preservedEntries = {
      'customXml/item1.xml': Array.from(new TextEncoder().encode(customXml)),
    };

    const wb2 = new Workbook(data);
    const buffer = await wb2.toBuffer();
    const wb3 = await readBuffer(buffer);

    const readData = wb3.toJSON();
    expect(readData.preservedEntries).toBeDefined();
    const entry = readData.preservedEntries?.['customXml/item1.xml'];
    expect(entry).toBeDefined();
    if (!entry) throw new Error('missing entry');
    const decoded = new TextDecoder().decode(new Uint8Array(entry));
    expect(decoded).toBe(customXml);
  });

  it('multiple preserved entry types coexist', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'test';

    const data = wb.toJSON();
    data.preservedEntries = {
      'xl/externalLinks/externalLink1.xml': Array.from(new TextEncoder().encode('<extLink/>')),
      'customXml/item1.xml': Array.from(new TextEncoder().encode('<custom/>')),
      'xl/media/image1.png': [0x89, 0x50, 0x4e, 0x47],
    };

    const wb2 = new Workbook(data);
    const buffer = await wb2.toBuffer();
    const wb3 = await readBuffer(buffer);

    const readData = wb3.toJSON();
    expect(readData.preservedEntries).toBeDefined();
    expect(readData.preservedEntries?.['xl/externalLinks/externalLink1.xml']).toBeDefined();
    expect(readData.preservedEntries?.['customXml/item1.xml']).toBeDefined();
    expect(readData.preservedEntries?.['xl/media/image1.png']).toBeDefined();
  });
});
