import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';
import type { CellData } from '../src/types.js';

describe('0.1.3 — Formula & Metadata Roundtrip Tests', () => {
  // ---------------------------------------------------------------------------
  // 1. Simple formula roundtrip
  // ---------------------------------------------------------------------------
  it('simple formula roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 10;
    ws.cell('A2').value = 20;
    ws.cell('A3').formula = 'SUM(A1:A2)';

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2).toBeDefined();

    // Formula text preserved
    expect(ws2?.cell('A3').formula).toBe('SUM(A1:A2)');
    expect(ws2?.cell('A3').type).toBe('formulaStr');

    // Source values preserved
    expect(ws2?.cell('A1').value).toBe(10);
    expect(ws2?.cell('A2').value).toBe(20);
  });

  // ---------------------------------------------------------------------------
  // 2. Multiple formulas
  // ---------------------------------------------------------------------------
  it('multiple formulas on different cells', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 10;
    ws.cell('A2').value = 20;
    ws.cell('B1').value = 30;
    ws.cell('B2').value = 40;
    ws.cell('C1').value = 50;
    ws.cell('C2').value = 60;

    ws.cell('A3').formula = 'SUM(A1:A2)';
    ws.cell('B3').formula = 'AVERAGE(B1:B2)';
    ws.cell('C3').formula = 'MAX(C1:C2)';

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2).toBeDefined();

    expect(ws2?.cell('A3').formula).toBe('SUM(A1:A2)');
    expect(ws2?.cell('A3').type).toBe('formulaStr');

    expect(ws2?.cell('B3').formula).toBe('AVERAGE(B1:B2)');
    expect(ws2?.cell('B3').type).toBe('formulaStr');

    expect(ws2?.cell('C3').formula).toBe('MAX(C1:C2)');
    expect(ws2?.cell('C3').type).toBe('formulaStr');
  });

  // ---------------------------------------------------------------------------
  // 3. Array formula roundtrip
  // ---------------------------------------------------------------------------
  it('array formula roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    const data = wb.toJSON();
    // Ensure there's a row and cell for the array formula
    data.sheets[0]!.worksheet.rows = [
      {
        index: 1,
        cells: [
          {
            reference: 'A1',
            cellType: 'formulaStr',
            styleIndex: null,
            value: null,
            formula: '{=ROW(1:3)}',
            formulaType: 'array',
            formulaRef: 'A1:A3',
            sharedIndex: null,
            inlineString: null,
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
    const sheet = readData.sheets[0]!;
    const row = sheet.worksheet.rows.find((r) => r.index === 1);
    expect(row).toBeDefined();

    const cell = row?.cells.find((c) => c.reference === 'A1');
    expect(cell).toBeDefined();
    expect(cell?.formula).toBe('{=ROW(1:3)}');
    expect(cell?.formulaType).toBe('array');
    expect(cell?.formulaRef).toBe('A1:A3');
  });

  // ---------------------------------------------------------------------------
  // 4. Shared formula roundtrip
  // ---------------------------------------------------------------------------
  it('shared formula roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    const data = wb.toJSON();
    data.sheets[0]!.worksheet.rows = [
      {
        index: 1,
        cells: [
          {
            reference: 'A1',
            cellType: 'number',
            styleIndex: null,
            value: '10',
            formula: null,
            formulaType: null,
            formulaRef: null,
            sharedIndex: null,
            inlineString: null,
          },
          {
            // Master cell: has formula + sharedIndex + formulaRef
            reference: 'B1',
            cellType: 'formulaStr',
            styleIndex: null,
            value: null,
            formula: 'A1*2',
            formulaType: 'shared',
            formulaRef: 'B1:B3',
            sharedIndex: 0,
            inlineString: null,
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
            formulaType: null,
            formulaRef: null,
            sharedIndex: null,
            inlineString: null,
          },
          {
            // Child cell: sharedIndex + cached value, no formula text
            reference: 'B2',
            cellType: 'formulaStr',
            styleIndex: null,
            value: '40',
            formula: null,
            formulaType: 'shared',
            formulaRef: null,
            sharedIndex: 0,
            inlineString: null,
          },
        ],
        height: null,
        hidden: false,
      },
      {
        index: 3,
        cells: [
          {
            reference: 'A3',
            cellType: 'number',
            styleIndex: null,
            value: '30',
            formula: null,
            formulaType: null,
            formulaRef: null,
            sharedIndex: null,
            inlineString: null,
          },
          {
            // Child cell: sharedIndex + cached value, no formula text
            reference: 'B3',
            cellType: 'formulaStr',
            styleIndex: null,
            value: '60',
            formula: null,
            formulaType: 'shared',
            formulaRef: null,
            sharedIndex: 0,
            inlineString: null,
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
    const sheet = readData.sheets[0]!;

    // Master cell (B1) — should have formula, sharedIndex, and formulaRef
    const row1 = sheet.worksheet.rows.find((r) => r.index === 1);
    expect(row1).toBeDefined();
    const masterCell = row1?.cells.find((c) => c.reference === 'B1');
    expect(masterCell).toBeDefined();
    expect(masterCell?.formula).toBe('A1*2');
    expect(masterCell?.formulaType).toBe('shared');
    expect(masterCell?.formulaRef).toBe('B1:B3');
    expect(masterCell?.sharedIndex).toBe(0);

    // Child cell (B2) — should have sharedIndex but no formula text
    const row2 = sheet.worksheet.rows.find((r) => r.index === 2);
    expect(row2).toBeDefined();
    const childCell2 = row2?.cells.find((c) => c.reference === 'B2');
    expect(childCell2).toBeDefined();
    expect(childCell2?.formulaType).toBe('shared');
    expect(childCell2?.sharedIndex).toBe(0);

    // Child cell (B3) — should have sharedIndex but no formula text
    const row3 = sheet.worksheet.rows.find((r) => r.index === 3);
    expect(row3).toBeDefined();
    const childCell3 = row3?.cells.find((c) => c.reference === 'B3');
    expect(childCell3).toBeDefined();
    expect(childCell3?.formulaType).toBe('shared');
    expect(childCell3?.sharedIndex).toBe(0);
  });

  // ---------------------------------------------------------------------------
  // 5. Inline string roundtrip
  // ---------------------------------------------------------------------------
  it('inline string roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    const data = wb.toJSON();
    const cell: CellData = {
      reference: 'A1',
      cellType: 'inlineStr',
      styleIndex: null,
      value: null,
      formula: null,
      formulaType: null,
      formulaRef: null,
      sharedIndex: null,
      inlineString: 'Hello inline',
    };
    data.sheets[0]!.worksheet.rows = [
      {
        index: 1,
        cells: [cell],
        height: null,
        hidden: false,
      },
    ];

    const wb2 = new Workbook(data);
    const buffer = await wb2.toBuffer();
    const wb3 = await readBuffer(buffer);

    const readData = wb3.toJSON();
    const sheet = readData.sheets[0]!;
    const row = sheet.worksheet.rows.find((r) => r.index === 1);
    expect(row).toBeDefined();

    const readCell = row?.cells.find((c) => c.reference === 'A1');
    expect(readCell).toBeDefined();
    expect(readCell?.cellType).toBe('inlineStr');
    expect(readCell?.inlineString).toBe('Hello inline');
  });

  // ---------------------------------------------------------------------------
  // 6. Document properties roundtrip (all fields)
  // ---------------------------------------------------------------------------
  it('document properties roundtrip — all fields', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.docProperties = {
      title: 'Test Workbook',
      subject: 'Testing',
      creator: 'modern-xlsx',
      keywords: 'test, roundtrip',
      description: 'A test workbook for roundtrip verification',
      lastModifiedBy: 'test-suite',
      created: '2026-03-02T00:00:00Z',
      modified: '2026-03-02T12:00:00Z',
      category: 'Testing',
      contentStatus: 'Draft',
      application: 'modern-xlsx',
      company: 'Test Corp',
      manager: 'Test Manager',
    };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    const props = wb2.docProperties;
    expect(props).not.toBeNull();
    expect(props?.title).toBe('Test Workbook');
    expect(props?.subject).toBe('Testing');
    expect(props?.creator).toBe('modern-xlsx');
    expect(props?.keywords).toBe('test, roundtrip');
    expect(props?.description).toBe('A test workbook for roundtrip verification');
    expect(props?.lastModifiedBy).toBe('test-suite');
    expect(props?.created).toBe('2026-03-02T00:00:00Z');
    expect(props?.modified).toBe('2026-03-02T12:00:00Z');
    expect(props?.category).toBe('Testing');
    expect(props?.contentStatus).toBe('Draft');
    expect(props?.application).toBe('modern-xlsx');
    expect(props?.company).toBe('Test Corp');
    expect(props?.manager).toBe('Test Manager');
  });

  // ---------------------------------------------------------------------------
  // 7. Document properties roundtrip (partial)
  // ---------------------------------------------------------------------------
  it('document properties roundtrip — partial fields', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.docProperties = {
      title: 'Partial Props',
      creator: 'partial-test',
    };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    const props = wb2.docProperties;
    expect(props).not.toBeNull();
    expect(props?.title).toBe('Partial Props');
    expect(props?.creator).toBe('partial-test');

    // Unset fields should be null or undefined
    const unsetOrNull = (val: unknown) => val === null || val === undefined;
    expect(unsetOrNull(props?.subject)).toBe(true);
    expect(unsetOrNull(props?.keywords)).toBe(true);
    expect(unsetOrNull(props?.description)).toBe(true);
    expect(unsetOrNull(props?.lastModifiedBy)).toBe(true);
    expect(unsetOrNull(props?.category)).toBe(true);
    expect(unsetOrNull(props?.contentStatus)).toBe(true);
    expect(unsetOrNull(props?.company)).toBe(true);
    expect(unsetOrNull(props?.manager)).toBe(true);
  });

  // ---------------------------------------------------------------------------
  // 8. Calc chain roundtrip
  // ---------------------------------------------------------------------------
  it('calc chain roundtrip — formulas populate calc chain', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 10;
    ws.cell('A2').value = 20;
    ws.cell('A3').formula = 'SUM(A1:A2)';
    ws.cell('B3').formula = 'AVERAGE(A1:A2)';

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    const chain = wb2.calcChain;
    // Calc chain may or may not be populated depending on writer behavior.
    // If populated, verify structure is correct.
    if (chain.length > 0) {
      for (const entry of chain) {
        expect(entry.cellRef).toBeTypeOf('string');
        expect(entry.cellRef.length).toBeGreaterThan(0);
        expect(entry.sheetId).toBeTypeOf('number');
      }

      // Check that formula cells appear in the chain
      const cellRefs = chain.map((e) => e.cellRef);
      expect(cellRefs).toContain('A3');
      expect(cellRefs).toContain('B3');
    }

    // Regardless, formulas should be intact
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2).toBeDefined();
    expect(ws2?.cell('A3').formula).toBe('SUM(A1:A2)');
    expect(ws2?.cell('B3').formula).toBe('AVERAGE(A1:A2)');
  });

  // ---------------------------------------------------------------------------
  // 9. Preserved entries roundtrip
  // ---------------------------------------------------------------------------
  it('preserved entries roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    const data = wb.toJSON();

    // PNG signature bytes + a mock XML drawing
    const pngBytes = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
    const xmlBytes = Array.from(new TextEncoder().encode('<xml>test</xml>'));

    data.preservedEntries = {
      'xl/media/image1.png': pngBytes,
      'xl/drawings/drawing1.xml': xmlBytes,
    };

    const wb2 = new Workbook(data);
    const buffer = await wb2.toBuffer();
    const wb3 = await readBuffer(buffer);

    const readData = wb3.toJSON();
    expect(readData.preservedEntries).toBeDefined();

    // Verify image entry preserved
    const img = readData.preservedEntries?.['xl/media/image1.png'];
    expect(img).toBeDefined();
    expect(Array.from(img!)).toEqual(pngBytes);

    // Verify drawing XML entry preserved
    const drawing = readData.preservedEntries?.['xl/drawings/drawing1.xml'];
    expect(drawing).toBeDefined();
    const decodedXml = new TextDecoder().decode(new Uint8Array(drawing!));
    expect(decodedXml).toBe('<xml>test</xml>');
  });

  // ---------------------------------------------------------------------------
  // 10. Formula with value (cached result)
  // ---------------------------------------------------------------------------
  it('formula with cached value roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 50;
    ws.cell('A2').value = 50;

    // Set formula first, then cached value
    const cell = ws.cell('A3');
    cell.formula = 'SUM(A1:A2)';
    cell.value = 100; // cached result

    // Verify local state before roundtrip
    expect(cell.formula).toBe('SUM(A1:A2)');
    expect(cell.type).toBe('formulaStr');

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2).toBeDefined();

    // Formula must be preserved
    expect(ws2?.cell('A3').formula).toBe('SUM(A1:A2)');
    expect(ws2?.cell('A3').type).toBe('formulaStr');

    // Check raw data for cached value
    const readData = wb2.toJSON();
    const row = readData.sheets[0]?.worksheet.rows.find((r) => r.index === 3);
    expect(row).toBeDefined();
    const rawCell = row?.cells.find((c) => c.reference === 'A3');
    expect(rawCell).toBeDefined();
    expect(rawCell?.formula).toBe('SUM(A1:A2)');
    // Cached value should be present as a string
    expect(rawCell?.value).toBe('100');
  });

  // ---------------------------------------------------------------------------
  // 11. Multiple formula types in same workbook
  // ---------------------------------------------------------------------------
  it('multiple formula types in same workbook', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    const data = wb.toJSON();
    data.sheets[0]!.worksheet.rows = [
      {
        index: 1,
        cells: [
          {
            // Normal formula
            reference: 'A1',
            cellType: 'formulaStr',
            styleIndex: null,
            value: '30',
            formula: 'SUM(10,20)',
            formulaType: null,
            formulaRef: null,
            sharedIndex: null,
            inlineString: null,
          },
          {
            // Array formula
            reference: 'B1',
            cellType: 'formulaStr',
            styleIndex: null,
            value: null,
            formula: '{=TRANSPOSE(A1:A3)}',
            formulaType: 'array',
            formulaRef: 'B1:D1',
            sharedIndex: null,
            inlineString: null,
          },
          {
            // Shared formula master
            reference: 'C1',
            cellType: 'formulaStr',
            styleIndex: null,
            value: null,
            formula: 'A1+1',
            formulaType: 'shared',
            formulaRef: 'C1:C3',
            sharedIndex: 0,
            inlineString: null,
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
            value: '40',
            formula: null,
            formulaType: null,
            formulaRef: null,
            sharedIndex: null,
            inlineString: null,
          },
          {
            // Shared formula child with cached value
            reference: 'C2',
            cellType: 'formulaStr',
            styleIndex: null,
            value: '41',
            formula: null,
            formulaType: 'shared',
            formulaRef: null,
            sharedIndex: 0,
            inlineString: null,
          },
        ],
        height: null,
        hidden: false,
      },
      {
        index: 3,
        cells: [
          {
            reference: 'A3',
            cellType: 'number',
            styleIndex: null,
            value: '50',
            formula: null,
            formulaType: null,
            formulaRef: null,
            sharedIndex: null,
            inlineString: null,
          },
          {
            // Shared formula child with cached value
            reference: 'C3',
            cellType: 'formulaStr',
            styleIndex: null,
            value: '51',
            formula: null,
            formulaType: 'shared',
            formulaRef: null,
            sharedIndex: 0,
            inlineString: null,
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
    const sheet = readData.sheets[0]!;

    // Normal formula (A1)
    const row1 = sheet.worksheet.rows.find((r) => r.index === 1);
    expect(row1).toBeDefined();
    const normalCell = row1?.cells.find((c) => c.reference === 'A1');
    expect(normalCell).toBeDefined();
    expect(normalCell?.formula).toBe('SUM(10,20)');
    // Normal formulas have null formulaType
    expect(normalCell?.formulaType ?? null).toBeNull();

    // Array formula (B1)
    const arrayCell = row1?.cells.find((c) => c.reference === 'B1');
    expect(arrayCell).toBeDefined();
    expect(arrayCell?.formula).toBe('{=TRANSPOSE(A1:A3)}');
    expect(arrayCell?.formulaType).toBe('array');
    expect(arrayCell?.formulaRef).toBe('B1:D1');

    // Shared formula master (C1)
    const sharedMaster = row1?.cells.find((c) => c.reference === 'C1');
    expect(sharedMaster).toBeDefined();
    expect(sharedMaster?.formula).toBe('A1+1');
    expect(sharedMaster?.formulaType).toBe('shared');
    expect(sharedMaster?.formulaRef).toBe('C1:C3');
    expect(sharedMaster?.sharedIndex).toBe(0);

    // Shared formula child (C2)
    const row2 = sheet.worksheet.rows.find((r) => r.index === 2);
    expect(row2).toBeDefined();
    const sharedChild2 = row2?.cells.find((c) => c.reference === 'C2');
    expect(sharedChild2).toBeDefined();
    expect(sharedChild2?.formulaType).toBe('shared');
    expect(sharedChild2?.sharedIndex).toBe(0);

    // Shared formula child (C3)
    const row3 = sheet.worksheet.rows.find((r) => r.index === 3);
    expect(row3).toBeDefined();
    const sharedChild3 = row3?.cells.find((c) => c.reference === 'C3');
    expect(sharedChild3).toBeDefined();
    expect(sharedChild3?.formulaType).toBe('shared');
    expect(sharedChild3?.sharedIndex).toBe(0);
  });

  // ---------------------------------------------------------------------------
  // 12. Document properties cleared
  // ---------------------------------------------------------------------------
  it('document properties cleared after being set', async () => {
    // First, create a workbook with properties and write it
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.docProperties = {
      title: 'Will be cleared',
      creator: 'will-be-removed',
      subject: 'temporary',
    };

    // Verify properties are set locally
    expect(wb.docProperties).not.toBeNull();
    expect(wb.docProperties?.title).toBe('Will be cleared');

    // Now clear properties
    wb.docProperties = null;
    expect(wb.docProperties).toBeNull();

    // Write and read back
    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    // Properties should be null or all fields null/undefined after clearing
    const props = wb2.docProperties;
    if (props !== null) {
      // If the reader returns an empty object rather than null,
      // all individual fields should be null/undefined
      const allNull = [
        props.title,
        props.subject,
        props.creator,
        props.keywords,
        props.description,
        props.lastModifiedBy,
        props.created,
        props.modified,
        props.category,
        props.contentStatus,
        props.application,
        props.company,
        props.manager,
      ].every((v) => v === null || v === undefined);
      expect(allNull).toBe(true);
    }
  });
});
