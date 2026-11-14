import { execFile } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { run } from '../src/cli-run.js';
import { Workbook } from '../src/index.js';
import { sheetToJson } from '../src/io/sheet-utils.js';

/** CLI tests: the getters and sheet utilities the commands use, run() in-process with captured stdout/stderr, and one real-process smoke test of the built dist/cli.mjs (skipped when dist is absent). */
describe('getters and utilities the CLI uses', () => {
  it('Worksheet.dimension returns the source dimension', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Test');
    // New sheet has no dimension
    expect(ws.dimension).toBeNull();
  });

  it('Worksheet.rowCount returns the number of rows', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Test');
    expect(ws.rowCount).toBe(0);
    ws.cell('A1').value = 'hello';
    expect(ws.rowCount).toBe(1);
    ws.cell('A5').value = 'world';
    expect(ws.rowCount).toBe(2);
  });

  it('sheetToJson converts worksheet to JSON objects', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'Name';
    ws.cell('B1').value = 'Age';
    ws.cell('A2').value = 'Alice';
    ws.cell('B2').value = 30;
    ws.cell('A3').value = 'Bob';
    ws.cell('B3').value = 25;

    const json = sheetToJson(ws);
    expect(json).toHaveLength(2);
    expect(json[0]).toEqual({ Name: 'Alice', Age: 30 });
    expect(json[1]).toEqual({ Name: 'Bob', Age: 25 });
  });

  it('info command properties are accessible', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.addSheet('Sheet2');
    expect(wb.sheetCount).toBe(2);
    expect(wb.sheetNames).toEqual(['Sheet1', 'Sheet2']);
    for (let i = 0; i < wb.sheetCount; i++) {
      const ws = wb.getSheetByIndex(i);
      expect(ws).toBeDefined();
      expect(ws?.rowCount).toBe(0);
      expect(ws?.dimension).toBeNull();
    }
  });
});

describe('run()', () => {
  let dir = '';
  let xlsx = '';

  beforeAll(async () => {
    dir = await mkdtemp(join(tmpdir(), 'mx-cli-'));
    xlsx = join(dir, 'in.xlsx');
    const wb = new Workbook();
    const ws = wb.addSheet('__proto__');
    ws.cell('A1').value = 'Name';
    ws.cell('A2').value = 'Alice';
    await writeFile(xlsx, await wb.toBuffer());
  });

  afterAll(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  const capture = () => {
    const out: string[] = [];
    const err: string[] = [];
    return {
      out,
      err,
      io: { stdout: (s: string) => out.push(s), stderr: (s: string) => err.push(s) },
    };
  };

  it('info returns 0 and lists the sheets', async () => {
    const c = capture();
    expect(await run(['info', xlsx], c.io)).toBe(0);
    expect(c.out.join('')).toContain('Sheets: 1');
  });

  it('convert writes CSV', async () => {
    const out = join(dir, 'out.csv');
    const c = capture();
    expect(await run(['convert', xlsx, out, '--format', 'csv'], c.io)).toBe(0);
    expect(await readFile(out, 'utf8')).toBe('Name\nAlice');
  });

  it('convert writes JSON and keeps a sheet named __proto__', async () => {
    const out = join(dir, 'out.json');
    const c = capture();
    expect(await run(['convert', xlsx, out], c.io)).toBe(0);
    const parsed: unknown = JSON.parse(await readFile(out, 'utf8'));
    expect(Object.hasOwn(parsed as object, '__proto__')).toBe(true);
    expect(Object.getOwnPropertyDescriptor(parsed, '__proto__')?.value).toEqual([
      { Name: 'Alice' },
    ]);
  });

  it('convert --sheet 0 writes one sheet as JSON', async () => {
    const out = join(dir, 'one.json');
    expect(await run(['convert', xlsx, out, '--sheet', '0'], capture().io)).toBe(0);
    expect(JSON.parse(await readFile(out, 'utf8'))).toEqual([{ Name: 'Alice' }]);
  });

  it('returns 2 for usage errors', async () => {
    const out = join(dir, 'never.out');
    for (const argv of [
      ['convert', xlsx, out, '--format', 'xml'],
      ['frobnicate'],
      ['info'],
      ['convert', xlsx],
      ['convert', xlsx, out, '--sheet', '-1'],
      ['convert', xlsx, out, '--sheet', 'x'],
      ['info', xlsx, '--nope'],
    ]) {
      const c = capture();
      expect(await run(argv, c.io), argv.join(' ')).toBe(2);
      expect(c.err.join('')).not.toBe('');
    }
    expect(existsSync(out)).toBe(false);
  });

  it('returns 1 for a sheet index out of range', async () => {
    const c = capture();
    expect(await run(['convert', xlsx, join(dir, 'x.json'), '--sheet', '5'], c.io)).toBe(1);
  });

  it('returns 1 for a missing file, without a stack trace', async () => {
    const c = capture();
    expect(await run(['info', join(dir, 'missing.xlsx')], c.io)).toBe(1);
    const text = c.err.join('');
    expect(text).toContain('modern-xlsx:');
    expect(text).not.toMatch(/\n\s+at /);
  });

  it('--help returns 0 and prints usage to stdout', async () => {
    const c = capture();
    expect(await run(['--help'], c.io)).toBe(0);
    expect(c.out.join('')).toContain('Usage:');
  });
});

describe('dist/cli.mjs', () => {
  const cli = fileURLToPath(new URL('../dist/cli.mjs', import.meta.url));

  it.skipIf(!existsSync(cli))('runs info as a real process and exits 0', async () => {
    const wb = new Workbook();
    wb.addSheet('S');
    const dir = await mkdtemp(join(tmpdir(), 'mx-cli-bin-'));
    try {
      const file = join(dir, 'f.xlsx');
      await writeFile(file, await wb.toBuffer());
      const { stdout } = await promisify(execFile)(process.execPath, [cli, 'info', file]);
      expect(stdout).toContain('Sheets: 1');
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
