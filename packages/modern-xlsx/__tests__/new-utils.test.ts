import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import {
  aoaToSheet,
  formatCell,
  getBuiltinFormat,
  initWasm,
  RichTextBuilder,
  sheetAddAoa,
  sheetAddJson,
  sheetToCsv,
  sheetToHtml,
  Workbook,
} from '../src/index.js';
import { initSync } from '../wasm/modern_xlsx_wasm.js';

beforeAll(async () => {
  const wasmBytes = await readFile(new URL('../wasm/modern_xlsx_wasm_bg.wasm', import.meta.url));
  initSync({ module: wasmBytes });
  await initWasm();
});

// ---------------------------------------------------------------------------
// formatCell
// ---------------------------------------------------------------------------

describe('formatCell', () => {
  it('returns the value as-is for General format', () => {
    expect(formatCell(123, 'General')).toBe('123');
    expect(formatCell('hello', 'General')).toBe('hello');
  });

  it('formats numbers with #,##0.00', () => {
    expect(formatCell(1234567.891, '#,##0.00')).toBe('1,234,567.89');
    expect(formatCell(0, '#,##0.00')).toBe('0.00');
    expect(formatCell(42, '#,##0.00')).toBe('42.00');
  });

  it('formats percentages with 0.00%', () => {
    expect(formatCell(0.1234, '0.00%')).toBe('12.34%');
    expect(formatCell(1, '0.00%')).toBe('100.00%');
    expect(formatCell(0, '0.00%')).toBe('0.00%');
  });

  it('formats dates with yyyy-mm-dd (serial 45292 = 2024-01-01)', () => {
    // Serial 45292 corresponds to 2024-01-01 in the date1900 system
    const result = formatCell(45292, 'yyyy-mm-dd');
    expect(result).toBe('2024-01-01');
  });

  it('formats dates with mm/dd/yy', () => {
    const result = formatCell(45292, 'mm/dd/yy');
    expect(result).toBe('01/01/24');
  });

  it('formats dates using built-in format ID 14 (mm-dd-yy)', () => {
    // Built-in format 14 is 'mm-dd-yy'
    const result = formatCell(45292, 14);
    expect(result).toBe('01-01-24');
  });

  it('formats scientific notation with 0.00E+00', () => {
    const result = formatCell(12345, '0.00E+00');
    // toExponential(2).toUpperCase() => "1.23E+4"
    expect(result).toBe('1.23E+4');
  });

  it('returns correct built-in format codes via getBuiltinFormat', () => {
    expect(getBuiltinFormat(0)).toBe('General');
    expect(getBuiltinFormat(1)).toBe('0');
    expect(getBuiltinFormat(2)).toBe('0.00');
    expect(getBuiltinFormat(9)).toBe('0%');
    expect(getBuiltinFormat(10)).toBe('0.00%');
    expect(getBuiltinFormat(14)).toBe('mm-dd-yy');
    expect(getBuiltinFormat(49)).toBe('@');
    expect(getBuiltinFormat(999)).toBeUndefined();
  });

  it('returns empty string for null and undefined', () => {
    expect(formatCell(null, 'General')).toBe('');
    expect(formatCell(undefined as unknown as null, '#,##0')).toBe('');
  });

  it('returns string as-is for text format @', () => {
    expect(formatCell('hello world', '@')).toBe('hello world');
    expect(formatCell(42, '@')).toBe('42');
  });

  it('formats integer percentages with 0%', () => {
    expect(formatCell(0.75, '0%')).toBe('75%');
  });

  it('handles built-in format ID 0 as General', () => {
    expect(formatCell(42, 0)).toBe('42');
  });
});

// ---------------------------------------------------------------------------
// RichTextBuilder
// ---------------------------------------------------------------------------

describe('RichTextBuilder', () => {
  it('.text() adds a plain run', () => {
    const builder = new RichTextBuilder();
    const runs = builder.text('Hello').build();
    expect(runs).toHaveLength(1);
    expect(runs[0]).toEqual({ text: 'Hello' });
  });

  it('.bold() and .italic() set their respective flags', () => {
    const builder = new RichTextBuilder();
    const runs = builder.bold('Bold').italic('Italic').build();
    expect(runs).toHaveLength(2);
    expect(runs[0]).toEqual({ text: 'Bold', bold: true });
    expect(runs[1]).toEqual({ text: 'Italic', italic: true });
  });

  it('.styled() applies custom options', () => {
    const builder = new RichTextBuilder();
    const runs = builder
      .styled('Custom', { bold: true, fontSize: 16, fontName: 'Arial', color: 'FF0000' })
      .build();
    expect(runs).toHaveLength(1);
    expect(runs[0]).toEqual({
      text: 'Custom',
      bold: true,
      fontSize: 16,
      fontName: 'Arial',
      color: 'FF0000',
    });
  });

  it('.plainText() concatenates all runs', () => {
    const builder = new RichTextBuilder();
    const text = builder.text('Hello ').bold('World').italic('!').plainText();
    expect(text).toBe('Hello World!');
  });

  it('.build() returns a copy of the runs array', () => {
    const builder = new RichTextBuilder();
    builder.text('A').bold('B');
    const runs1 = builder.build();
    const runs2 = builder.build();
    expect(runs1).toEqual(runs2);
    expect(runs1).not.toBe(runs2); // different array instances
  });

  it('.boldItalic() sets both bold and italic', () => {
    const builder = new RichTextBuilder();
    const runs = builder.boldItalic('Both').build();
    expect(runs[0]).toEqual({ text: 'Both', bold: true, italic: true });
  });

  it('.colored() sets the color property', () => {
    const builder = new RichTextBuilder();
    const runs = builder.colored('Red text', 'FF0000').build();
    expect(runs[0]).toEqual({ text: 'Red text', color: 'FF0000' });
  });
});

// ---------------------------------------------------------------------------
// sheetToHtml
// ---------------------------------------------------------------------------

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
});

// ---------------------------------------------------------------------------
// sheetAddAoa
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// sheetAddJson
// ---------------------------------------------------------------------------

describe('sheetAddJson', () => {
  it('appends with headers by default', () => {
    const ws = aoaToSheet([['Existing']]);
    sheetAddJson(ws, [{ x: 1, y: 2 }]);
    // Headers should be written at the next empty row after existing data
    const csv = sheetToCsv(ws);
    expect(csv).toContain('x');
    expect(csv).toContain('y');
    expect(csv).toContain('1');
    expect(csv).toContain('2');
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
    const csv = sheetToCsv(ws);
    // The headers should be in the specified order: a, b
    expect(csv).toContain('a');
    expect(csv).toContain('b');
  });

  it('handles empty data array gracefully', () => {
    const ws = aoaToSheet([['Existing']]);
    const rowsBefore = ws.rows.length;
    sheetAddJson(ws, []);
    expect(ws.rows.length).toBe(rowsBefore);
  });
});
