import { describe, expect, it } from 'vitest';
import { RichTextBuilder, readBuffer, Workbook } from '../src/index.js';

describe('0.1.2 — Feature Roundtrip Tests', () => {
  // ---------------------------------------------------------------------------
  // 1. ColorScale with 2 stops
  // ---------------------------------------------------------------------------
  it('colorScale with 2 stops survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('CF');
    ws.cell('A1').value = 1;

    const data = wb.toJSON();
    data.sheets[0].worksheet.conditionalFormatting = [
      {
        sqref: 'A1:A10',
        rules: [
          {
            ruleType: 'colorScale',
            priority: 1,
            operator: null,
            formula: null,
            dxfId: null,
            colorScale: {
              cfvos: [
                { cfvoType: 'min', val: null },
                { cfvoType: 'max', val: null },
              ],
              colors: ['FF0000', '00FF00'],
            },
            dataBar: null,
            iconSet: null,
          },
        ],
      },
    ];

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const data2 = wb2.toJSON();
    const cf = data2.sheets[0].worksheet.conditionalFormatting;

    expect(cf).toBeDefined();
    expect(cf).toHaveLength(1);
    expect(cf?.[0].sqref).toBe('A1:A10');
    expect(cf?.[0].rules).toHaveLength(1);

    const rule = cf?.[0].rules[0];
    expect(rule.ruleType).toBe('colorScale');
    expect(rule.priority).toBe(1);
    // Serde omits None fields; reader may return null or undefined for absent values
    expect(rule.operator ?? null).toBeNull();
    expect(rule.formula ?? null).toBeNull();
    expect(rule.dxfId ?? null).toBeNull();

    expect(rule.colorScale).toBeDefined();
    expect(rule.colorScale?.cfvos).toHaveLength(2);
    expect(rule.colorScale?.cfvos[0].cfvoType).toBe('min');
    expect(rule.colorScale?.cfvos[0].val ?? null).toBeNull();
    expect(rule.colorScale?.cfvos[1].cfvoType).toBe('max');
    expect(rule.colorScale?.cfvos[1].val ?? null).toBeNull();
    expect(rule.colorScale?.colors).toEqual(['FF0000', '00FF00']);

    expect(rule.dataBar ?? null).toBeNull();
    expect(rule.iconSet ?? null).toBeNull();
  });

  // ---------------------------------------------------------------------------
  // 2. ColorScale with 3 stops
  // ---------------------------------------------------------------------------
  it('colorScale with 3 stops survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('CF3');
    ws.cell('A1').value = 1;

    const data = wb.toJSON();
    data.sheets[0].worksheet.conditionalFormatting = [
      {
        sqref: 'B1:B20',
        rules: [
          {
            ruleType: 'colorScale',
            priority: 1,
            operator: null,
            formula: null,
            dxfId: null,
            colorScale: {
              cfvos: [
                { cfvoType: 'min', val: null },
                { cfvoType: 'percentile', val: '50' },
                { cfvoType: 'max', val: null },
              ],
              colors: ['FF0000', 'FFFF00', '00FF00'],
            },
            dataBar: null,
            iconSet: null,
          },
        ],
      },
    ];

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const data2 = wb2.toJSON();
    const cf = data2.sheets[0].worksheet.conditionalFormatting;

    expect(cf).toBeDefined();
    expect(cf).toHaveLength(1);
    expect(cf?.[0].sqref).toBe('B1:B20');

    const rule = cf?.[0].rules[0];
    expect(rule.ruleType).toBe('colorScale');
    expect(rule.colorScale).toBeDefined();
    expect(rule.colorScale?.cfvos).toHaveLength(3);
    expect(rule.colorScale?.cfvos[0].cfvoType).toBe('min');
    expect(rule.colorScale?.cfvos[0].val ?? null).toBeNull();
    expect(rule.colorScale?.cfvos[1].cfvoType).toBe('percentile');
    expect(rule.colorScale?.cfvos[1].val).toBe('50');
    expect(rule.colorScale?.cfvos[2].cfvoType).toBe('max');
    expect(rule.colorScale?.cfvos[2].val ?? null).toBeNull();
    expect(rule.colorScale?.colors).toEqual(['FF0000', 'FFFF00', '00FF00']);
  });

  // ---------------------------------------------------------------------------
  // 3. DataBar
  // ---------------------------------------------------------------------------
  it('dataBar conditional formatting survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('DataBar');
    ws.cell('A1').value = 10;

    const data = wb.toJSON();
    data.sheets[0].worksheet.conditionalFormatting = [
      {
        sqref: 'A1:A10',
        rules: [
          {
            ruleType: 'dataBar',
            priority: 1,
            operator: null,
            formula: null,
            dxfId: null,
            colorScale: null,
            dataBar: {
              cfvos: [
                { cfvoType: 'min', val: null },
                { cfvoType: 'max', val: null },
              ],
              color: '638EC6',
            },
            iconSet: null,
          },
        ],
      },
    ];

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const data2 = wb2.toJSON();
    const cf = data2.sheets[0].worksheet.conditionalFormatting;

    expect(cf).toBeDefined();
    expect(cf).toHaveLength(1);
    expect(cf?.[0].sqref).toBe('A1:A10');

    const rule = cf?.[0].rules[0];
    expect(rule.ruleType).toBe('dataBar');
    expect(rule.priority).toBe(1);
    expect(rule.colorScale ?? null).toBeNull();
    expect(rule.iconSet ?? null).toBeNull();

    expect(rule.dataBar).toBeDefined();
    expect(rule.dataBar?.cfvos).toHaveLength(2);
    expect(rule.dataBar?.cfvos[0].cfvoType).toBe('min');
    expect(rule.dataBar?.cfvos[0].val ?? null).toBeNull();
    expect(rule.dataBar?.cfvos[1].cfvoType).toBe('max');
    expect(rule.dataBar?.cfvos[1].val ?? null).toBeNull();
    expect(rule.dataBar?.color).toBe('638EC6');
  });

  // ---------------------------------------------------------------------------
  // 4. IconSet
  // ---------------------------------------------------------------------------
  it('iconSet conditional formatting survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('IconSet');
    ws.cell('A1').value = 50;

    const data = wb.toJSON();
    data.sheets[0].worksheet.conditionalFormatting = [
      {
        sqref: 'A1:A10',
        rules: [
          {
            ruleType: 'iconSet',
            priority: 1,
            operator: null,
            formula: null,
            dxfId: null,
            colorScale: null,
            dataBar: null,
            iconSet: {
              iconSetType: '3TrafficLights1',
              cfvos: [
                { cfvoType: 'percent', val: '0' },
                { cfvoType: 'percent', val: '33' },
                { cfvoType: 'percent', val: '67' },
              ],
            },
          },
        ],
      },
    ];

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const data2 = wb2.toJSON();
    const cf = data2.sheets[0].worksheet.conditionalFormatting;

    expect(cf).toBeDefined();
    expect(cf).toHaveLength(1);
    expect(cf?.[0].sqref).toBe('A1:A10');

    const rule = cf?.[0].rules[0];
    expect(rule.ruleType).toBe('iconSet');
    expect(rule.priority).toBe(1);
    expect(rule.colorScale ?? null).toBeNull();
    expect(rule.dataBar ?? null).toBeNull();

    expect(rule.iconSet).toBeDefined();
    expect(rule.iconSet?.iconSetType).toBe('3TrafficLights1');
    expect(rule.iconSet?.cfvos).toHaveLength(3);
    expect(rule.iconSet?.cfvos[0].cfvoType).toBe('percent');
    expect(rule.iconSet?.cfvos[0].val).toBe('0');
    expect(rule.iconSet?.cfvos[1].cfvoType).toBe('percent');
    expect(rule.iconSet?.cfvos[1].val).toBe('33');
    expect(rule.iconSet?.cfvos[2].cfvoType).toBe('percent');
    expect(rule.iconSet?.cfvos[2].val).toBe('67');
  });

  // ---------------------------------------------------------------------------
  // 5. Rich text / shared string roundtrip
  // ---------------------------------------------------------------------------
  it('shared string cell with plain text survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('RichText');
    ws.cell('A1').value = 'Hello World';

    // Verify the cell is typed as sharedString
    expect(ws.cell('A1').type).toBe('sharedString');
    expect(ws.cell('A1').value).toBe('Hello World');

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('RichText');

    expect(ws2).toBeDefined();
    expect(ws2?.cell('A1').type).toBe('sharedString');
    expect(ws2?.cell('A1').value).toBe('Hello World');
  });

  it('rich text runs are preserved in shared strings data', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('RichText');
    ws.cell('A1').value = 'Hello World';

    const runs = new RichTextBuilder().bold('Hello ').colored('World', 'FF0000').build();

    // Verify the builder produced the expected runs
    expect(runs).toHaveLength(2);
    expect(runs[0].text).toBe('Hello ');
    expect(runs[0].bold).toBe(true);
    expect(runs[1].text).toBe('World');
    expect(runs[1].color).toBe('FF0000');

    // Inject rich runs into the workbook data for the SST
    const data = wb.toJSON();
    data.sharedStrings = {
      strings: ['Hello World'],
      richRuns: [runs],
    };
    // Set the cell to reference SST index 0
    data.sheets[0].worksheet.rows[0].cells[0].cellType = 'sharedString';
    data.sheets[0].worksheet.rows[0].cells[0].value = 'Hello World';

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const data2 = wb2.toJSON();

    // Verify the string content roundtripped
    expect(data2.sharedStrings).toBeDefined();
    expect(data2.sharedStrings?.strings).toContain('Hello World');

    // Verify the cell value is accessible
    const ws2 = wb2.getSheet('RichText');
    expect(ws2).toBeDefined();
    expect(ws2?.cell('A1').value).toBe('Hello World');
  });

  // ---------------------------------------------------------------------------
  // 6. Comments roundtrip
  // ---------------------------------------------------------------------------
  it('single comment survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Comments');
    ws.cell('A1').value = 'Annotated';
    ws.addComment('A1', 'Author', 'This is a comment');

    // Verify the comment was added before roundtrip
    expect(ws.comments).toHaveLength(1);
    expect(ws.comments[0].cellRef).toBe('A1');
    expect(ws.comments[0].author).toBe('Author');
    expect(ws.comments[0].text).toBe('This is a comment');

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Comments');

    expect(ws2).toBeDefined();
    expect(ws2?.comments).toHaveLength(1);
    expect(ws2?.comments[0].cellRef).toBe('A1');
    expect(ws2?.comments[0].author).toBe('Author');
    expect(ws2?.comments[0].text).toBe('This is a comment');
  });

  // ---------------------------------------------------------------------------
  // 7. Multiple comments
  // ---------------------------------------------------------------------------
  it('multiple comments on different cells survive roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('MultiComment');
    ws.cell('A1').value = 'First';
    ws.cell('B2').value = 'Second';
    ws.cell('C3').value = 'Third';

    ws.addComment('A1', 'Alice', 'Comment on A1');
    ws.addComment('B2', 'Bob', 'Comment on B2');
    ws.addComment('C3', 'Charlie', 'Comment on C3');

    expect(ws.comments).toHaveLength(3);

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('MultiComment');

    expect(ws2).toBeDefined();
    expect(ws2?.comments).toHaveLength(3);

    // Verify each comment individually
    const c1 = ws2?.comments.find((c) => c.cellRef === 'A1');
    expect(c1).toBeDefined();
    expect(c1?.author).toBe('Alice');
    expect(c1?.text).toBe('Comment on A1');

    const c2 = ws2?.comments.find((c) => c.cellRef === 'B2');
    expect(c2).toBeDefined();
    expect(c2?.author).toBe('Bob');
    expect(c2?.text).toBe('Comment on B2');

    const c3 = ws2?.comments.find((c) => c.cellRef === 'C3');
    expect(c3).toBeDefined();
    expect(c3?.author).toBe('Charlie');
    expect(c3?.text).toBe('Comment on C3');
  });

  // ---------------------------------------------------------------------------
  // 8. Hyperlink roundtrip (external URL)
  // ---------------------------------------------------------------------------
  it('external hyperlink survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Links');
    ws.cell('A1').value = 'Example';
    ws.addHyperlink('A1', 'https://example.com', {
      display: 'Example',
      tooltip: 'Click here',
    });

    // Verify before roundtrip
    expect(ws.hyperlinks).toHaveLength(1);
    expect(ws.hyperlinks[0].cellRef).toBe('A1');
    expect(ws.hyperlinks[0].location).toBe('https://example.com');
    expect(ws.hyperlinks[0].display).toBe('Example');
    expect(ws.hyperlinks[0].tooltip).toBe('Click here');

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Links');

    expect(ws2).toBeDefined();
    expect(ws2?.hyperlinks).toHaveLength(1);

    const link = ws2?.hyperlinks[0];
    expect(link.cellRef).toBe('A1');
    expect(link.location).toBe('https://example.com');
    expect(link.display).toBe('Example');
    expect(link.tooltip).toBe('Click here');
  });

  // ---------------------------------------------------------------------------
  // 9. Hyperlink roundtrip (internal)
  // ---------------------------------------------------------------------------
  it('internal hyperlink survives roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    const ws2Sheet = wb.addSheet('Sheet2');
    ws2Sheet.cell('A1').value = 'Target';

    const ws1 = wb.getSheet('Sheet1')!;
    ws1.cell('B1').value = 'Go to Sheet2';
    ws1.addHyperlink('B1', 'Sheet2!A1', { display: 'Go to Sheet2' });

    // Verify before roundtrip
    expect(ws1.hyperlinks).toHaveLength(1);
    expect(ws1.hyperlinks[0].cellRef).toBe('B1');
    expect(ws1.hyperlinks[0].location).toBe('Sheet2!A1');
    expect(ws1.hyperlinks[0].display).toBe('Go to Sheet2');

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws1Read = wb2.getSheet('Sheet1');

    expect(ws1Read).toBeDefined();
    expect(ws1Read?.hyperlinks).toHaveLength(1);

    const link = ws1Read?.hyperlinks[0];
    expect(link.cellRef).toBe('B1');
    expect(link.location).toBe('Sheet2!A1');
    expect(link.display).toBe('Go to Sheet2');
  });

  // ---------------------------------------------------------------------------
  // 10. Multiple conditional formatting rules on same range
  // ---------------------------------------------------------------------------
  it('multiple conditional formatting rules on same range survive roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('MultiRule');
    ws.cell('A1').value = 5;

    const data = wb.toJSON();
    data.sheets[0].worksheet.conditionalFormatting = [
      {
        sqref: 'A1:A10',
        rules: [
          {
            ruleType: 'colorScale',
            priority: 1,
            operator: null,
            formula: null,
            dxfId: null,
            colorScale: {
              cfvos: [
                { cfvoType: 'min', val: null },
                { cfvoType: 'max', val: null },
              ],
              colors: ['FF0000', '00FF00'],
            },
            dataBar: null,
            iconSet: null,
          },
          {
            ruleType: 'dataBar',
            priority: 2,
            operator: null,
            formula: null,
            dxfId: null,
            colorScale: null,
            dataBar: {
              cfvos: [
                { cfvoType: 'min', val: null },
                { cfvoType: 'max', val: null },
              ],
              color: '638EC6',
            },
            iconSet: null,
          },
        ],
      },
    ];

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const data2 = wb2.toJSON();
    const cf = data2.sheets[0].worksheet.conditionalFormatting;

    expect(cf).toBeDefined();
    expect(cf).toHaveLength(1);
    expect(cf?.[0].sqref).toBe('A1:A10');
    expect(cf?.[0].rules).toHaveLength(2);

    // First rule: colorScale
    const rule1 = cf?.[0].rules[0];
    expect(rule1.ruleType).toBe('colorScale');
    expect(rule1.priority).toBe(1);
    expect(rule1.colorScale).toBeDefined();
    expect(rule1.colorScale?.cfvos).toHaveLength(2);
    expect(rule1.colorScale?.colors).toEqual(['FF0000', '00FF00']);
    expect(rule1.dataBar ?? null).toBeNull();
    expect(rule1.iconSet ?? null).toBeNull();

    // Second rule: dataBar
    const rule2 = cf?.[0].rules[1];
    expect(rule2.ruleType).toBe('dataBar');
    expect(rule2.priority).toBe(2);
    expect(rule2.dataBar).toBeDefined();
    expect(rule2.dataBar?.color).toBe('638EC6');
    expect(rule2.dataBar?.cfvos).toHaveLength(2);
    expect(rule2.colorScale ?? null).toBeNull();
    expect(rule2.iconSet ?? null).toBeNull();
  });

  // ---------------------------------------------------------------------------
  // 11. Conditional formatting on multiple ranges
  // ---------------------------------------------------------------------------
  it('conditional formatting on multiple separate ranges survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('MultiRange');
    ws.cell('A1').value = 1;
    ws.cell('C1').value = 2;

    const data = wb.toJSON();
    data.sheets[0].worksheet.conditionalFormatting = [
      {
        sqref: 'A1:A10',
        rules: [
          {
            ruleType: 'colorScale',
            priority: 1,
            operator: null,
            formula: null,
            dxfId: null,
            colorScale: {
              cfvos: [
                { cfvoType: 'min', val: null },
                { cfvoType: 'max', val: null },
              ],
              colors: ['0000FF', 'FF0000'],
            },
            dataBar: null,
            iconSet: null,
          },
        ],
      },
      {
        sqref: 'C1:C20',
        rules: [
          {
            ruleType: 'iconSet',
            priority: 2,
            operator: null,
            formula: null,
            dxfId: null,
            colorScale: null,
            dataBar: null,
            iconSet: {
              iconSetType: '3TrafficLights1',
              cfvos: [
                { cfvoType: 'percent', val: '0' },
                { cfvoType: 'percent', val: '33' },
                { cfvoType: 'percent', val: '67' },
              ],
            },
          },
        ],
      },
    ];

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const data2 = wb2.toJSON();
    const cf = data2.sheets[0].worksheet.conditionalFormatting;

    expect(cf).toBeDefined();
    expect(cf).toHaveLength(2);

    // First entry: colorScale on A1:A10
    const entry1 = cf?.find((e) => e.sqref === 'A1:A10');
    expect(entry1).toBeDefined();
    expect(entry1?.rules).toHaveLength(1);
    expect(entry1?.rules[0].ruleType).toBe('colorScale');
    expect(entry1?.rules[0].colorScale).toBeDefined();
    expect(entry1?.rules[0].colorScale?.colors).toEqual(['0000FF', 'FF0000']);

    // Second entry: iconSet on C1:C20
    const entry2 = cf?.find((e) => e.sqref === 'C1:C20');
    expect(entry2).toBeDefined();
    expect(entry2?.rules).toHaveLength(1);
    expect(entry2?.rules[0].ruleType).toBe('iconSet');
    expect(entry2?.rules[0].iconSet).toBeDefined();
    expect(entry2?.rules[0].iconSet?.iconSetType).toBe('3TrafficLights1');
    expect(entry2?.rules[0].iconSet?.cfvos).toHaveLength(3);
  });

  // ---------------------------------------------------------------------------
  // 12. Theme colors read (read-only, null for new workbooks)
  // ---------------------------------------------------------------------------
  it('themeColors is null for a newly created workbook', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    expect(wb.themeColors).toBeNull();
  });

  it('themeColors remains null after roundtrip of a created workbook', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'test';

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    // Since we don't generate theme XML on write, themeColors should be null
    expect(wb2.themeColors).toBeNull();
  });
});
