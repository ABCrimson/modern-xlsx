import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import { initWasm, Workbook } from '../src/index.js';
import { aoaToSheet, jsonToSheet, sheetToCsv, sheetToJson } from '../src/utils.js';
import { initSync } from '../wasm/modern_xlsx_wasm.js';

beforeAll(async () => {
  const wasmBytes = await readFile(new URL('../wasm/modern_xlsx_wasm_bg.wasm', import.meta.url));
  initSync({ module: wasmBytes });
  await initWasm();
});

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
