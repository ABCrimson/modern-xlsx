import { describe, expect, it } from 'vitest';
import {
  aoaToSheet,
  jsonToSheet,
  readBuffer,
  sheetAddAoa,
  sheetAddJson,
  sheetToCsv,
  sheetToFormulae,
  sheetToHtml,
  sheetToJson,
  sheetToTxt,
  Workbook,
} from '../src/index.js';
import type { CellType } from '../src/model/types.js';
import { buildZip, readZipText } from './helpers/zip.js';

describe('aoaToSheet', () => {
  it('creates a worksheet from a 2D array', () => {
    const ws = aoaToSheet([
      ['Name', 'Age'],
      ['Alice', 30],
      ['Bob', 25],
    ]);
    expect(ws.rows).toHaveLength(3);
    expect(ws.cell('A1').value).toBe('Name');
    expect(ws.cell('B1').value).toBe('Age');
    expect(ws.cell('A2').value).toBe('Alice');
    expect(ws.cell('B2').value).toBe(30);
  });

  it('respects origin option', () => {
    const ws = aoaToSheet([['X']], { origin: 'C3' });
    expect(ws.cell('C3').value).toBe('X');
  });

  it('handles empty array', () => {
    const ws = aoaToSheet([]);
    expect(ws.rows).toHaveLength(0);
  });
});

describe('jsonToSheet', () => {
  it('creates a worksheet from JSON objects', () => {
    const ws = jsonToSheet([
      { name: 'Alice', age: 30 },
      { name: 'Bob', age: 25 },
    ]);
    // Header row
    expect(ws.cell('A1').value).toBe('name');
    expect(ws.cell('B1').value).toBe('age');
    // Data rows
    expect(ws.cell('A2').value).toBe('Alice');
    expect(ws.cell('B2').value).toBe(30);
  });

  it('skipHeader omits the header row', () => {
    const ws = jsonToSheet([{ x: 1 }], { skipHeader: true });
    expect(ws.cell('A1').value).toBe(1);
  });

  it('custom header order', () => {
    const ws = jsonToSheet([{ b: 2, a: 1 }], { header: ['a', 'b'] });
    expect(ws.cell('A1').value).toBe('a');
    expect(ws.cell('B1').value).toBe('b');
    expect(ws.cell('A2').value).toBe(1);
    expect(ws.cell('B2').value).toBe(2);
  });

  it('handles empty array', () => {
    const ws = jsonToSheet([]);
    expect(ws.rows).toHaveLength(0);
  });
});

describe('sheetToCsv', () => {
  it('converts worksheet to CSV', () => {
    const ws = aoaToSheet([
      ['Name', 'Age'],
      ['Alice', 30],
    ]);
    const csv = sheetToCsv(ws);
    expect(csv).toBe('Name,Age\nAlice,30');
  });

  it('quotes fields with commas', () => {
    const ws = aoaToSheet([['Hello, World']]);
    const csv = sheetToCsv(ws);
    expect(csv).toBe('"Hello, World"');
  });

  it('quotes fields with double quotes', () => {
    const ws = aoaToSheet([['He said "hi"']]);
    const csv = sheetToCsv(ws);
    expect(csv).toBe('"He said ""hi"""');
  });

  it('custom separator', () => {
    const ws = aoaToSheet([['A', 'B']]);
    const csv = sheetToCsv(ws, { separator: '\t' });
    expect(csv).toBe('A\tB');
  });

  it('forceQuote wraps all fields', () => {
    const ws = aoaToSheet([['A', 'B']]);
    const csv = sheetToCsv(ws, { forceQuote: true });
    expect(csv).toBe('"A","B"');
  });

  it('handles empty worksheet', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Empty');
    const csv = sheetToCsv(ws);
    expect(csv).toBe('');
  });
});

describe('sheetToJson', () => {
  it('converts worksheet with header row to JSON', () => {
    const ws = aoaToSheet([
      ['Name', 'Age'],
      ['Alice', 30],
      ['Bob', 25],
    ]);
    const data = sheetToJson(ws);
    expect(data).toHaveLength(2);
    expect(data[0]).toEqual({ Name: 'Alice', Age: 30 });
    expect(data[1]).toEqual({ Name: 'Bob', Age: 25 });
  });

  it('uses column letters as keys with header "A"', () => {
    const ws = aoaToSheet([
      ['Alice', 30],
      ['Bob', 25],
    ]);
    const data = sheetToJson(ws, { header: 'A' });
    expect(data.length).toBeGreaterThan(0);
    // With header: 'A', keys are column letters
    expect(data[0]).toHaveProperty('A');
    expect(data[0]).toHaveProperty('B');
  });
});

describe('roundtrip: aoaToSheet -> sheetToJson', () => {
  it('roundtrips basic data', () => {
    const input = [
      ['Name', 'Score'],
      ['Alice', 95],
      ['Bob', 87],
    ];
    const ws = aoaToSheet(input);
    const json = sheetToJson(ws);
    expect(json).toHaveLength(2);
    expect(json[0]).toEqual({ Name: 'Alice', Score: 95 });
    expect(json[1]).toEqual({ Name: 'Bob', Score: 87 });
  });
});

describe('sheetToJson range, headers and prototype keys', () => {
  const sample = () =>
    aoaToSheet([
      ['Name', 'Score'],
      ['a', 1],
      ['b', 2],
      ['c', 3],
    ]);

  it("range 'A1:B3' includes row 3 (endRow is 1-based inclusive)", () => {
    expect(sheetToJson(sample(), { range: 'A1:B3' })).toEqual([
      { Name: 'a', Score: 1 },
      { Name: 'b', Score: 2 },
    ]);
  });

  it("range 'A2:B3' with header 'A' returns rows 2 and 3", () => {
    expect(sheetToJson(sample(), { range: 'A2:B3', header: 'A' })).toEqual([
      { A: 'a', B: 1 },
      { A: 'b', B: 2 },
    ]);
  });

  it('without a range returns all three data rows', () => {
    expect(sheetToJson(sample())).toHaveLength(3);
  });

  it('suffixes duplicate headers', () => {
    const ws = aoaToSheet([
      ['Name', 'Name'],
      ['x', 'y'],
    ]);
    expect(sheetToJson(ws)).toEqual([{ Name: 'x', Name_1: 'y' }]);
  });

  it('skips suffixes that collide with another raw header', () => {
    const ws = aoaToSheet([
      ['Name', 'Name', 'Name_1'],
      ['x', 'y', 'z'],
    ]);
    expect(sheetToJson(ws)).toEqual([{ Name: 'x', Name_2: 'y', Name_1: 'z' }]);
  });

  it('keeps a __proto__ header as an own key', () => {
    const ws = aoaToSheet([['__proto__'], ['x']]);
    const [row] = sheetToJson(ws);
    expect(Object.hasOwn(row as object, '__proto__')).toBe(true);
    expect(Object.getPrototypeOf(row)).toBe(Object.prototype);
    expect(Object.keys(row as object)).toEqual(['__proto__']);
  });

  it('sheetRows counts data objects for JSON and output lines for CSV', () => {
    const sparse = aoaToSheet([['H']]);
    sparse.cell('A100').value = 'r100';
    expect(sheetToJson(sparse, { sheetRows: 2 })).toEqual([{ H: 'r100' }]);
    expect(sheetToCsv(sparse, { sheetRows: 2 }).split('\n')).toEqual(['H', '']);
  });

  it('header: number uses that 1-based row of the range as the keys', () => {
    const ws = aoaToSheet([['x'], ['Name', 'Age'], ['a', 1]]);
    expect(sheetToJson(ws, { header: 2 })).toEqual([{ Name: 'a', Age: 1 }]);
  });

  it('header: string[] supplies the keys and keeps the first row as data', () => {
    const ws = aoaToSheet([['x'], ['Name', 'Age'], ['a', 1]]);
    expect(sheetToJson(ws, { header: ['p', 'q'], range: 'A2:B3' })).toEqual([
      { p: 'Name', q: 'Age' },
      { p: 'a', q: 1 },
    ]);
  });

  it('defval fills missing cells and is absent otherwise', () => {
    const ws = aoaToSheet([['a', 'b'], [1]]);
    expect(sheetToJson(ws, { defval: null })).toEqual([{ a: 1, b: null }]);
    expect(sheetToJson(ws)).toEqual([{ a: 1 }]);
  });
});

describe('sheetToCsv quoting and formula escaping', () => {
  it('quotes a field containing a carriage return', () => {
    const ws = aoaToSheet([['a\rb', 'c']]);
    expect(sheetToCsv(ws)).toBe('"a\rb",c');
  });

  it('escapeFormulas prefixes and force-quotes text that starts with a formula trigger', () => {
    const ws = aoaToSheet([['=1+1', '+x', '@y', '-z', '\tw', -5, 'ok']]);
    expect(sheetToCsv(ws, { escapeFormulas: true })).toBe(`"'=1+1","'+x","'@y","'-z","'\tw",-5,ok`);
  });

  it('leaves output unchanged without the option', () => {
    const ws = aoaToSheet([['=1+1', -5]]);
    expect(sheetToCsv(ws)).toBe('=1+1,-5');
  });
});

describe('sheetToHtml', () => {
  it('returns an empty table for an empty worksheet', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Empty');
    const html = sheetToHtml(ws);
    expect(html).toBe('<table></table>');
  });

  it('renders basic data as correct HTML', () => {
    const ws = aoaToSheet([
      ['Name', 'Age'],
      ['Alice', 30],
    ]);
    const html = sheetToHtml(ws);
    expect(html).toContain('<table>');
    expect(html).toContain('</table>');
    expect(html).toContain('<tr>');
    expect(html).toContain('<td>Name</td>');
    expect(html).toContain('<td>Age</td>');
    expect(html).toContain('<td>Alice</td>');
    expect(html).toContain('<td>30</td>');
  });

  it('applies className option to the table element', () => {
    const ws = aoaToSheet([['A']]);
    const html = sheetToHtml(ws, { className: 'my-table' });
    expect(html).toContain('<table class="my-table">');
  });

  it('creates thead/tbody when header option is true', () => {
    const ws = aoaToSheet([
      ['Name', 'Age'],
      ['Alice', 30],
    ]);
    const html = sheetToHtml(ws, { header: true });
    expect(html).toContain('<thead>');
    expect(html).toContain('</thead>');
    expect(html).toContain('<tbody>');
    expect(html).toContain('</tbody>');
    expect(html).toContain('<th>Name</th>');
    expect(html).toContain('<th>Age</th>');
    expect(html).toContain('<td>Alice</td>');
  });

  it('escapes HTML entities in cell values', () => {
    const ws = aoaToSheet([['<script>alert("xss")</script>']]);
    const html = sheetToHtml(ws);
    expect(html).not.toContain('<script>');
    expect(html).toContain('&lt;script&gt;');
    expect(html).toContain('&quot;');
  });

  it('applies className with special characters escaped', () => {
    const ws = aoaToSheet([['Test']]);
    const html = sheetToHtml(ws, { className: 'table "special"' });
    expect(html).toContain('class="table &quot;special&quot;"');
  });

  it('includeStyle aligns numbers right and booleans center', () => {
    const ws = aoaToSheet([[1, true, 't']]);
    expect(sheetToHtml(ws, { includeStyle: true })).toBe(
      '<table><tr><td style="text-align:right">1</td><td style="text-align:center">TRUE</td><td>t</td></tr></table>',
    );
  });
});

describe('sheetAddAoa', () => {
  it('appends data at the next row by default', () => {
    const ws = aoaToSheet([
      ['Name', 'Age'],
      ['Alice', 30],
    ]);
    sheetAddAoa(ws, [['Bob', 25]]);
    expect(ws.cell('A3').value).toBe('Bob');
    expect(ws.cell('B3').value).toBe(25);
  });

  it('uses origin option to start at a specified cell', () => {
    const ws = aoaToSheet([['Existing']]);
    sheetAddAoa(ws, [['New']], { origin: 'C5' });
    expect(ws.cell('C5').value).toBe('New');
  });

  it('handles mixed data types', () => {
    const ws = aoaToSheet([['Header']]);
    sheetAddAoa(ws, [['text', 42, true]]);
    expect(ws.cell('A2').value).toBe('text');
    expect(ws.cell('B2').value).toBe(42);
    expect(ws.cell('C2').value).toBe(true);
  });

  it('skips null and undefined values', () => {
    const ws = aoaToSheet([['A']]);
    sheetAddAoa(ws, [[null, 'B', undefined, 'D']]);
    // Inspect the row before any ws.cell() call: cell() creates a missing record.
    expect(ws.rows.find((r) => r.index === 2)?.cells.map((c) => c.reference)).toEqual(['B2', 'D2']);
    expect(ws.cell('A2').value).toBeNull();
    expect(ws.cell('B2').value).toBe('B');
    expect(ws.cell('D2').value).toBe('D');
  });

  it('appends multiple rows at once', () => {
    const ws = aoaToSheet([['H1']]);
    sheetAddAoa(ws, [['R1'], ['R2'], ['R3']]);
    expect(ws.cell('A2').value).toBe('R1');
    expect(ws.cell('A3').value).toBe('R2');
    expect(ws.cell('A4').value).toBe('R3');
  });
});

describe('sheetAddJson', () => {
  it('appends with headers by default', () => {
    const ws = aoaToSheet([['Existing']]);
    sheetAddJson(ws, [{ x: 1, y: 2 }]);
    // Headers are written at the next empty row after existing data
    expect(sheetToCsv(ws)).toBe('Existing,\nx,y\n1,2');
  });

  it('skipHeader option skips the header row', () => {
    const ws = aoaToSheet([['x', 'y']]);
    sheetAddJson(ws, [{ x: 10, y: 20 }], { skipHeader: true });
    // The data should start right after existing data without a new header row
    const csv = sheetToCsv(ws);
    const lines = csv.split('\n');
    // First line is the original header
    expect(lines[0]).toBe('x,y');
    // Second line should be the data, not another header row
    expect(lines[1]).toBe('10,20');
  });

  it('uses origin option to start at a specified cell', () => {
    const ws = aoaToSheet([['Existing']]);
    sheetAddJson(ws, [{ a: 1 }], { origin: 'C3' });
    expect(ws.cell('C3').value).toBe('a'); // header
    expect(ws.cell('C4').value).toBe(1); // data
  });

  it('uses custom header order', () => {
    const ws = aoaToSheet([['Existing']]);
    sheetAddJson(ws, [{ b: 2, a: 1 }], { header: ['a', 'b'] });
    // The headers are in the specified order: a, b
    expect(sheetToCsv(ws)).toBe('Existing,\na,b\n1,2');
    expect(ws.cell('A2').value).toBe('a');
    expect(ws.cell('B2').value).toBe('b');
    expect(ws.cell('A3').value).toBe(1);
    expect(ws.cell('B3').value).toBe(2);
  });

  it('handles empty data array gracefully', () => {
    const ws = aoaToSheet([['Existing']]);
    const rowsBefore = ws.rows.length;
    sheetAddJson(ws, []);
    expect(ws.rows.length).toBe(rowsBefore);
  });
});

describe('sheetToTxt', () => {
  it('produces tab-separated output', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 'Name';
    ws.cell('B1').value = 'Age';
    ws.cell('A2').value = 'Alice';
    ws.cell('B2').value = 30;

    const txt = sheetToTxt(ws);
    const lines = txt.split('\n');
    expect(lines[0]).toBe('Name\tAge');
    expect(lines[1]).toBe('Alice\t30');
  });

  it('respects sheetRows limit', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 'r1';
    ws.cell('A2').value = 'r2';
    ws.cell('A3').value = 'r3';

    const txt = sheetToTxt(ws, { sheetRows: 2 });
    expect(txt.split('\n')).toHaveLength(2);
  });
});

describe('sheetToFormulae', () => {
  it('returns cell references with values and formulas', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 100;
    ws.cell('A2').value = 200;
    ws.cell('A3').formula = 'SUM(A1:A2)';

    const formulae = sheetToFormulae(ws);
    expect(formulae).toContain('A1=100');
    expect(formulae).toContain('A2=200');
    expect(formulae).toContain("A3='SUM(A1:A2)");
  });

  it('handles string values', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 'hello';

    const formulae = sheetToFormulae(ws);
    expect(formulae).toContain("A1='hello");
  });

  it('handles empty sheet', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    expect(sheetToFormulae(ws)).toEqual([]);
  });
});

describe('sheetToJson header de-duplication', () => {
  const headers = (row: string[]) =>
    Object.keys(sheetToJson(aoaToSheet([row, row.map(() => 1)]))[0] ?? {});

  it('de-duplicates 16,384 identical headers in linear time', () => {
    const row = Array.from({ length: 16_384 }, () => 'a');
    // Each repeat resumes from the last suffix issued for its base name, so N identical headers cost O(N) probes; the 10 s testTimeout bounds this case.
    const keys = headers(row);
    expect(keys).toHaveLength(16_384);
    expect(keys[0]).toBe('a');
    expect(keys[1]).toBe('a_1');
    expect(keys[16_383]).toBe('a_16383');
  });

  it('skips candidates that collide with a raw header name', () => {
    expect(headers(['a', 'a', 'a_1'])).toEqual(['a', 'a_2', 'a_1']);
  });

  it('numbers repeats of one name from 1', () => {
    expect(headers(['x', 'x', 'x'])).toEqual(['x', 'x_1', 'x_2']);
  });
});

describe('sheetToCsv formula escaping does not trust the declared cellType', () => {
  const typed = (value: string, cellType: CellType) => {
    const ws = aoaToSheet([['placeholder']]);
    const cell = ws.rows[0]?.cells[0];
    if (!cell) throw new Error('no cell');
    cell.cellType = cellType;
    cell.value = value;
    return ws;
  };

  it.each([
    ['=1+1', 'number'],
    ['-1+cmd', 'number'],
    ['@SUM(A1)', 'number'],
    ['=x', 'boolean'],
    ['+x', 'error'],
    ['=x', 'sharedString'],
  ] as const)('prefixes %s declared %s', (value, cellType) => {
    expect(sheetToCsv(typed(value, cellType), { escapeFormulas: true })).toBe(`"'${value}"`);
  });

  it.each(['-5', '-1.5e3', '+2'])('leaves number-typed %s unchanged', (value) => {
    expect(sheetToCsv(typed(value, 'number'), { escapeFormulas: true })).toBe(value);
  });

  it('prefixes nothing when escapeFormulas is false', () => {
    expect(sheetToCsv(typed('=1+1', 'number'))).toBe('=1+1');
  });
});

describe('cells and rows without an r attribute', () => {
  const enc = new TextEncoder();
  const main = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

  /** A real one-sheet workbook whose sheet1.xml holds `sheetData` instead. */
  async function workbookWithSheetData(sheetData: string): Promise<Uint8Array> {
    const wb = new Workbook();
    wb.addSheet('S').cell('A1').value = 'x';
    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ${main}><sheetData>${sheetData}</sheetData></worksheet>`;
    const parts = Object.entries(readZipText(await wb.toBuffer())).map(
      ([name, body]) =>
        [name, enc.encode(name === 'xl/worksheets/sheet1.xml' ? sheet : body)] as const,
    );
    return buildZip(parts);
  }

  it('take the next position, so the sheet converts (ECMA-376 Part 1 §18.3.1.4)', async () => {
    const wb = await readBuffer(
      await workbookWithSheetData('<row><c><v>1</v></c><c><v>2</v></c></row>'),
    );
    const ws = wb.getSheet('S');
    if (!ws) throw new Error('no sheet S');
    expect(ws.rows.map((r) => [r.index, r.cells.map((c) => c.reference)])).toEqual([
      [1, ['A1', 'B1']],
    ]);
    expect(sheetToCsv(ws)).toBe('1,2');
    expect(sheetToJson(ws, { header: 'A' })).toEqual([{ A: 1, B: 2 }]);
  });

  it('continue after a cell or row that has one', async () => {
    const wb = await readBuffer(
      await workbookWithSheetData(
        '<row><c><v>1</v></c><c><v>2</v></c></row><row><c r="C2"><v>3</v></c><c><v>4</v></c></row>',
      ),
    );
    const ws = wb.getSheet('S');
    if (!ws) throw new Error('no sheet S');
    expect(sheetToCsv(ws)).toBe('1,2,,\n,,3,4');
  });
});

describe('sheet-to-text conversions reject output no string can hold', () => {
  const corners = () => {
    const ws = aoaToSheet([['a']]);
    ws.cell('XFD1048576').value = 'z';
    return ws;
  };

  it.each([
    ['sheetToCsv', sheetToCsv],
    ['sheetToTxt', sheetToTxt],
    ['sheetToHtml', sheetToHtml],
  ] as const)('%s throws SIZE_LIMIT for A1 + XFD1048576 in under 100 ms', (_, convert) => {
    const ws = corners();
    const start = performance.now();
    expect(() => convert(ws)).toThrow(expect.objectContaining({ code: 'SIZE_LIMIT' }));
    expect(performance.now() - start).toBeLessThan(100);
  });
});

describe('Date values', () => {
  it('aoaToSheet writes a Date as ISO 8601 UTC text', () => {
    const ws = aoaToSheet([[new Date(Date.UTC(2024, 0, 1))]]);
    expect(ws.cell('A1').value).toBe('2024-01-01T00:00:00.000Z');
  });

  it('jsonToSheet skips an invalid Date', () => {
    const ws = jsonToSheet([{ d: new Date(Number.NaN) }]);
    expect(ws.cell('A1').value).toBe('d');
    expect(ws.cell('A2').value).toBeNull();
  });

  it('an invalid Date leaves no cell behind, like null', () => {
    expect(sheetToCsv(aoaToSheet([[1, new Date(Number.NaN)]]))).toBe('1');
    expect(sheetToCsv(jsonToSheet([{ a: 1, d: new Date(Number.NaN) }], { skipHeader: true }))).toBe(
      '1',
    );
  });

  it('sheetAddAoa and sheetAddJson write a Date as ISO 8601 UTC text', () => {
    const ws = aoaToSheet([['h']]);
    sheetAddAoa(ws, [[new Date(Date.UTC(2024, 1, 29, 12, 30))]]);
    sheetAddJson(ws, [{ d: new Date(Date.UTC(1999, 11, 31)) }], { skipHeader: true });
    expect(ws.cell('A2').value).toBe('2024-02-29T12:30:00.000Z');
    expect(ws.cell('A3').value).toBe('1999-12-31T00:00:00.000Z');
  });
});

describe('sheetToJson sheetRows stops early', () => {
  it('sheetRows: 2 on a 5-row sheet returns the first 2 data objects', () => {
    const ws = aoaToSheet([['n'], [1], [2], [3], [4]]);
    expect(sheetToJson(ws, { sheetRows: 2 })).toEqual([{ n: 1 }, { n: 2 }]);
    expect(sheetToJson(ws, { sheetRows: 0 })).toHaveLength(4);
    expect(sheetToJson(ws, { sheetRows: -1 })).toHaveLength(4);
    // A fractional limit truncates, as the earlier slice did.
    expect(sheetToJson(ws, { sheetRows: 2.5 })).toEqual([{ n: 1 }, { n: 2 }]);
    expect(sheetToJson(ws, { sheetRows: 0.5 })).toEqual([]);
  });
});

describe('sheetToHtml escaping', () => {
  it('escapes & < > and " and nothing else', () => {
    const ws = aoaToSheet([['a&b<c>"d'], ["it's"]]);
    expect(sheetToHtml(ws)).toBe(
      "<table><tr><td>a&amp;b&lt;c&gt;&quot;d</td></tr><tr><td>it's</td></tr></table>",
    );
  });
});

describe('JSON records never write values inherited from Object.prototype', () => {
  const header = ['constructor', 'toString', 'hasOwnProperty'];
  const data = () => [{ constructor: 'a', toString: 'b' }, {}];
  const refs = (ws: ReturnType<typeof aoaToSheet>, index: number) =>
    ws.rows.find((r) => r.index === index)?.cells.map((c) => [c.reference, c.value]) ?? [];

  it('jsonToSheet writes own values and leaves inherited ones out', () => {
    const ws = jsonToSheet(data(), { header });
    expect(refs(ws, 2)).toEqual([
      ['A2', 'a'],
      ['B2', 'b'],
    ]);
    expect(refs(ws, 3)).toEqual([]);
  });

  it('sheetAddJson writes own values and leaves inherited ones out', () => {
    const ws = aoaToSheet([['top']]);
    sheetAddJson(ws, data(), { header });
    expect(refs(ws, 3)).toEqual([
      ['A3', 'a'],
      ['B3', 'b'],
    ]);
    expect(refs(ws, 4)).toEqual([]);
  });

  it('still reads a getter a class defines on its own prototype', () => {
    class Person {
      get name(): string {
        return 'Ada';
      }
    }
    const ws = jsonToSheet([new Person() as unknown as Record<string, unknown>], {
      header: ['name'],
    });
    expect(ws.cell('A2').value).toBe('Ada');
  });
});
