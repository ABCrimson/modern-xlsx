import { describe, expect, it } from 'vitest';
import type {
  PivotCacheDefinitionData,
  PivotCacheRecordsData,
  PivotTableData,
  SlicerCacheData,
  SlicerData,
  TimelineCacheData,
  TimelineData,
  Worksheet,
} from '../src/index.js';
import { readBuffer, Workbook } from '../src/index.js';

// ---------------------------------------------------------------------------
// Helpers — minimal valid data objects
// ---------------------------------------------------------------------------

function samplePivotTable(name = 'PivotTable1', cacheId = 0): PivotTableData {
  return {
    name,
    dataCaption: 'Values',
    location: { ref: 'A3:D10', firstHeaderRow: 1, firstDataRow: 2, firstDataCol: 1 },
    pivotFields: [
      {
        axis: 'axisRow',
        items: [{ x: 0 }, { t: 'default' }],
        subtotals: [],
        compact: true,
        outline: true,
      },
      { items: [], subtotals: ['sum'], compact: false, outline: false },
    ],
    rowFields: [{ x: 0 }],
    colFields: [],
    dataFields: [{ name: 'Sum of Amount', fld: 1, subtotal: 'sum' }],
    pageFields: [],
    cacheId,
  };
}

function sampleSlicer(name = 'Slicer_Category'): SlicerData {
  return {
    name,
    caption: 'Category',
    cacheName: 'Slicer_Category',
    columnName: 'Category',
    sortOrder: 'ascending',
    startItem: 0,
  };
}

function sampleTimeline(name = 'Timeline_Date'): TimelineData {
  return {
    name,
    caption: 'Date',
    cacheName: 'NativeTimeline_Date',
    sourceName: 'Date',
    level: 'months',
  };
}

function samplePivotCache(): PivotCacheDefinitionData {
  return {
    source: { ref: 'A1:B5', sheet: 'Sheet1' },
    fields: [
      {
        name: 'Category',
        sharedItems: [
          { type: 'string', v: 'A' },
          { type: 'string', v: 'B' },
        ],
      },
      { name: 'Amount', sharedItems: [] },
    ],
    recordCount: 4,
  };
}

function samplePivotCacheRecords(): PivotCacheRecordsData {
  return {
    records: [
      [
        { type: 'index', v: 0 },
        { type: 'number', v: 100 },
      ],
      [
        { type: 'index', v: 1 },
        { type: 'number', v: 200 },
      ],
    ],
  };
}

function sampleSlicerCache(): SlicerCacheData {
  return {
    name: 'Slicer_Category',
    sourceName: 'Category',
    items: [
      { n: 'A', s: true },
      { n: 'B', s: false },
    ],
  };
}

function sampleTimelineCache(): TimelineCacheData {
  return {
    name: 'NativeTimeline_Date',
    sourceName: 'Date',
    selectionStart: '2024-01-01',
    selectionEnd: '2024-12-31',
  };
}

// ---------------------------------------------------------------------------
// Worksheet-level mutations
// ---------------------------------------------------------------------------

describe('Worksheet pivot table mutations', () => {
  it('addPivotTable adds to the array', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.pivotTables).toHaveLength(0);

    ws.addPivotTable(samplePivotTable());
    expect(ws.pivotTables).toHaveLength(1);
    expect(ws.pivotTables[0]?.name).toBe('PivotTable1');
  });

  it('addPivotTable initializes array if absent', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    // Confirm starting empty
    expect(ws.pivotTables).toEqual([]);

    ws.addPivotTable(samplePivotTable('PT_A'));
    ws.addPivotTable(samplePivotTable('PT_B'));
    expect(ws.pivotTables).toHaveLength(2);
    expect(ws.pivotTables[1]?.name).toBe('PT_B');
  });

  it('removePivotTable removes by index', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addPivotTable(samplePivotTable('PT_A'));
    ws.addPivotTable(samplePivotTable('PT_B'));
    ws.addPivotTable(samplePivotTable('PT_C'));

    expect(ws.removePivotTable(1)).toBe(true);
    expect(ws.pivotTables).toHaveLength(2);
    expect(ws.pivotTables[0]?.name).toBe('PT_A');
    expect(ws.pivotTables[1]?.name).toBe('PT_C');
  });

  it('removePivotTable returns false for invalid index', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.removePivotTable(0)).toBe(false);
    expect(ws.removePivotTable(-1)).toBe(false);
    expect(ws.removePivotTable(99)).toBe(false);
  });
});

describe('Worksheet slicer mutations', () => {
  it('addSlicer adds to the array', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.slicers).toHaveLength(0);

    ws.addSlicer(sampleSlicer());
    expect(ws.slicers).toHaveLength(1);
    expect(ws.slicers[0]?.name).toBe('Slicer_Category');
  });

  it('addSlicer initializes array if absent', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addSlicer(sampleSlicer('S1'));
    ws.addSlicer(sampleSlicer('S2'));
    expect(ws.slicers).toHaveLength(2);
  });

  it('removeSlicer removes by index', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addSlicer(sampleSlicer('S1'));
    ws.addSlicer(sampleSlicer('S2'));

    expect(ws.removeSlicer(0)).toBe(true);
    expect(ws.slicers).toHaveLength(1);
    expect(ws.slicers[0]?.name).toBe('S2');
  });

  it('removeSlicer returns false for invalid index', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.removeSlicer(0)).toBe(false);
    expect(ws.removeSlicer(-1)).toBe(false);
  });
});

describe('Worksheet timeline mutations', () => {
  it('addTimeline adds to the array', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.timelines).toHaveLength(0);

    ws.addTimeline(sampleTimeline());
    expect(ws.timelines).toHaveLength(1);
    expect(ws.timelines[0]?.name).toBe('Timeline_Date');
  });

  it('addTimeline initializes array if absent', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addTimeline(sampleTimeline('T1'));
    ws.addTimeline(sampleTimeline('T2'));
    expect(ws.timelines).toHaveLength(2);
  });

  it('removeTimeline removes by index', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addTimeline(sampleTimeline('T1'));
    ws.addTimeline(sampleTimeline('T2'));

    expect(ws.removeTimeline(0)).toBe(true);
    expect(ws.timelines).toHaveLength(1);
    expect(ws.timelines[0]?.name).toBe('T2');
  });

  it('removeTimeline returns false for invalid index', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.removeTimeline(0)).toBe(false);
    expect(ws.removeTimeline(-1)).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Workbook-level cache methods
// ---------------------------------------------------------------------------

describe('Workbook pivot cache methods', () => {
  it('addPivotCache adds to the array', () => {
    const wb = new Workbook();
    expect(wb.pivotCaches).toHaveLength(0);

    wb.addPivotCache(samplePivotCache());
    expect(wb.pivotCaches).toHaveLength(1);
    expect(wb.pivotCaches[0]?.source.sheet).toBe('Sheet1');
  });

  it('addPivotCacheRecords adds to the array', () => {
    const wb = new Workbook();
    expect(wb.pivotCacheRecords).toHaveLength(0);

    wb.addPivotCacheRecords(samplePivotCacheRecords());
    expect(wb.pivotCacheRecords).toHaveLength(1);
    expect(wb.pivotCacheRecords[0]?.records).toHaveLength(2);
  });
});

describe('Workbook slicer cache methods', () => {
  it('addSlicerCache adds to the array', () => {
    const wb = new Workbook();
    expect(wb.slicerCaches).toHaveLength(0);

    wb.addSlicerCache(sampleSlicerCache());
    expect(wb.slicerCaches).toHaveLength(1);
    expect(wb.slicerCaches[0]?.name).toBe('Slicer_Category');
  });
});

describe('Workbook timeline cache methods', () => {
  it('addTimelineCache adds to the array', () => {
    const wb = new Workbook();
    expect(wb.timelineCaches).toHaveLength(0);

    wb.addTimelineCache(sampleTimelineCache());
    expect(wb.timelineCaches).toHaveLength(1);
    expect(wb.timelineCaches[0]?.name).toBe('NativeTimeline_Date');
  });
});

// ---------------------------------------------------------------------------
// Roundtrip via WASM
// ---------------------------------------------------------------------------

describe('Pivot table roundtrip', () => {
  it('create workbook with pivot table + cache → toBuffer → readBuffer → verify', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    // Seed some data
    ws.cell('A1').value = 'Category';
    ws.cell('B1').value = 'Amount';
    ws.cell('A2').value = 'A';
    ws.cell('B2').value = '100';
    ws.cell('A3').value = 'B';
    ws.cell('B3').value = '200';

    // Add pivot cache + records at workbook level
    wb.addPivotCache(samplePivotCache());
    wb.addPivotCacheRecords(samplePivotCacheRecords());

    // Add pivot table on the sheet
    ws.addPivotTable(samplePivotTable('MyPivot', 0));

    const buffer = await wb.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1') as Worksheet;

    // Pivot table should survive roundtrip
    expect(ws2.pivotTables).toHaveLength(1);
    expect(ws2.pivotTables[0]?.name).toBe('MyPivot');
    expect(ws2.pivotTables[0]?.dataFields).toHaveLength(1);
    expect(ws2.pivotTables[0]?.dataFields[0]?.subtotal).toBe('sum');

    // Pivot caches should survive roundtrip
    expect(wb2.pivotCaches).toHaveLength(1);
    expect(wb2.pivotCaches[0]?.source.sheet).toBe('Sheet1');
    expect(wb2.pivotCaches[0]?.source.ref).toBe('A1:B5');
  });
});

describe('Slicer roundtrip', () => {
  it('create workbook with slicer + cache → toBuffer → readBuffer → verify', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Category';

    wb.addSlicerCache(sampleSlicerCache());
    ws.addSlicer(sampleSlicer());

    const buffer = await wb.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);

    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1') as Worksheet;

    expect(ws2.slicers).toHaveLength(1);
    expect(ws2.slicers[0]?.name).toBe('Slicer_Category');

    expect(wb2.slicerCaches).toHaveLength(1);
    expect(wb2.slicerCaches[0]?.name).toBe('Slicer_Category');
  });
});

describe('Timeline roundtrip', () => {
  it('create workbook with timeline + cache → toBuffer → readBuffer → verify', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Date';

    wb.addTimelineCache(sampleTimelineCache());
    ws.addTimeline(sampleTimeline());

    const buffer = await wb.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);

    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1') as Worksheet;

    expect(ws2.timelines).toHaveLength(1);
    expect(ws2.timelines[0]?.name).toBe('Timeline_Date');

    expect(wb2.timelineCaches).toHaveLength(1);
    expect(wb2.timelineCaches[0]?.name).toBe('NativeTimeline_Date');
  });
});
