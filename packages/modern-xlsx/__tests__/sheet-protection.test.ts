import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import type { SheetProtectionData } from '../src/index.js';
import { INVALID_WORKBOOK, ModernXlsxError, readBuffer, Workbook } from '../src/index.js';
import {
  initWasmSync as initLite,
  Workbook as LiteWorkbook,
  readBuffer as readLite,
} from '../src/index-lite.js';
import { readZipText } from './helpers/zip.js';

/** The `<sheetProtection .../>` element of the first worksheet in `bytes`. */
function protectionElement(bytes: Uint8Array): string {
  const xml = readZipText(bytes)['xl/worksheets/sheet1.xml'] ?? '';
  return xml.match(/<sheetProtection[^>]*\/>/)?.[0] ?? '';
}

/** Every flag of a protected sheet locked except sorting and filtering. */
const SORT_AND_FILTER_ALLOWED: SheetProtectionData = {
  sheet: true,
  objects: true,
  scenarios: true,
  formatCells: true,
  formatColumns: true,
  formatRows: true,
  insertColumns: true,
  insertRows: true,
  deleteColumns: true,
  deleteRows: true,
  sort: false,
  autoFilter: false,
  plainPassword: 'secret',
};

describe('sheet protection', () => {
  it('writes allowed sort and autoFilter as explicit 0 and reads them back', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'x';
    ws.sheetProtection = SORT_AND_FILTER_ALLOWED;
    const bytes = await wb.toBuffer();

    const element = protectionElement(bytes);
    expect(element).toContain('sort="0"');
    expect(element).toContain('autoFilter="0"');
    expect(element).toMatch(
      /^<sheetProtection algorithmName="SHA-512" hashValue="[A-Za-z0-9+/]{86}==" saltValue="[A-Za-z0-9+/]{22}==" spinCount="100000" sheet="1" objects="1" scenarios="1" sort="0" autoFilter="0"\/>$/,
    );
    expect(element).not.toContain('secret');

    const back = (await readBuffer(bytes)).getSheet('Data')?.sheetProtection;
    expect(back).toMatchObject({
      sheet: true,
      objects: true,
      scenarios: true,
      formatCells: true,
      insertRows: true,
      deleteColumns: true,
      sort: false,
      autoFilter: false,
      algorithmName: 'SHA-512',
      spinCount: 100000,
    });
    expect(back?.plainPassword).toBeUndefined();
    expect(back?.insertHyperlinks).toBeUndefined();
  });

  it('a 1.2.x-shaped object with every flag false allows each of those actions', async () => {
    const wb = new Workbook();
    wb.addSheet('Data').sheetProtection = {
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
    const bytes = await wb.toBuffer();
    expect(protectionElement(bytes)).toBe(
      '<sheetProtection sheet="1" formatCells="0" formatColumns="0" formatRows="0" insertColumns="0" insertRows="0" deleteColumns="0" deleteRows="0" sort="0" autoFilter="0"/>',
    );
    const back = (await readBuffer(bytes)).getSheet('Data')?.sheetProtection;
    expect(back).toMatchObject({ sheet: true, formatCells: false, sort: false, autoFilter: false });
  });

  it('a read workbook writes its protection back unchanged', async () => {
    const wb = new Workbook();
    wb.addSheet('Data').sheetProtection = {
      ...SORT_AND_FILTER_ALLOWED,
      selectLockedCells: true,
      pivotTables: false,
    };
    const first = await wb.toBuffer();
    const second = await (await readBuffer(first)).toBuffer();
    expect(protectionElement(second)).toBe(protectionElement(first));
    expect(protectionElement(second)).toContain(
      'selectLockedCells="1" sort="0" autoFilter="0" pivotTables="0"',
    );
  });

  it('rejects a password that is not a stored legacy hash with INVALID_WORKBOOK', async () => {
    const wb = new Workbook();
    wb.addSheet('Payroll').sheetProtection = {
      sheet: true,
      objects: false,
      scenarios: false,
      formatCells: true,
      formatColumns: true,
      formatRows: true,
      insertColumns: true,
      insertRows: true,
      deleteColumns: true,
      deleteRows: true,
      sort: true,
      autoFilter: true,
      password: 'secret',
    };
    const rejection = wb.toBuffer();
    await expect(rejection).rejects.toBeInstanceOf(ModernXlsxError);
    await expect(rejection).rejects.toMatchObject({ code: INVALID_WORKBOOK });
    await expect(rejection).rejects.toThrow(/'Payroll'.*plainPassword/);
    await expect(rejection).rejects.not.toThrow(/secret/);

    const issue = wb.validate().issues.find((i) => i.category === 'sheetProtection');
    expect(issue).toMatchObject({
      severity: 'error',
      location: 'Payroll!sheetProtection',
      autoFixable: false,
    });
    expect(issue?.message).toContain('plainPassword');
  });

  it('plainPassword "secret" still round-trips to hash attributes, replacing a bad password', async () => {
    const wb = new Workbook();
    wb.addSheet('Data').sheetProtection = {
      ...SORT_AND_FILTER_ALLOWED,
      password: 'secret',
      plainPassword: 'secret',
    };
    expect(wb.validate().issues.some((i) => i.category === 'sheetProtection')).toBe(false);
    const bytes = await wb.toBuffer();
    const element = protectionElement(bytes);
    expect(element).toMatch(
      /^<sheetProtection algorithmName="SHA-512" hashValue="[A-Za-z0-9+/]{86}==" saltValue="[A-Za-z0-9+/]{22}==" spinCount="100000" sheet="1"/,
    );
    expect(element).not.toContain('secret');
    expect(element).not.toContain('password=');
    const back = (await readBuffer(bytes)).getSheet('Data')?.sheetProtection;
    expect(back).toMatchObject({ algorithmName: 'SHA-512', spinCount: 100000 });
    expect(back?.password).toBeUndefined();
    expect(back?.plainPassword).toBeUndefined();
  });

  it('modern-xlsx/lite hashes plainPassword with the legacy 16-bit hash', async () => {
    initLite(readFileSync(new URL('../wasm-lite/modern_xlsx_wasm_bg.wasm', import.meta.url)));
    const wb = new LiteWorkbook();
    wb.addSheet('Data').sheetProtection = SORT_AND_FILTER_ALLOWED;
    const bytes = await wb.toBuffer();
    expect(protectionElement(bytes)).toBe(
      '<sheetProtection password="DAA7" sheet="1" objects="1" scenarios="1" sort="0" autoFilter="0"/>',
    );
    const back = (await readLite(bytes)).getSheet('Data')?.sheetProtection;
    expect(back?.password).toBe('DAA7');
  });
});
