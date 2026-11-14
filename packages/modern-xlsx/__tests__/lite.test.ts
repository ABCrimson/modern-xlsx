import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import {
  INVALID_ARGUMENT,
  initWasmSync,
  ModernXlsxError,
  readBuffer,
  readFile,
  StreamingXlsxWriter,
  Workbook,
} from '../src/index-lite.js';
import { buildOle2 } from './helpers/ole2.js';

const LITE_PASSWORD_MESSAGE =
  "Password-protected workbooks need the full build: import from 'modern-xlsx', not 'modern-xlsx/lite'.";

/** A one-sheet lite workbook with a string and a number. */
function sampleWorkbook(): Workbook {
  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  ws.cell('A1').value = 'lite';
  ws.cell('B1').value = 42;
  return wb;
}

describe('modern-xlsx/lite', () => {
  beforeAll(() => {
    initWasmSync(readFileSync(new URL('../wasm-lite/modern_xlsx_wasm_bg.wasm', import.meta.url)));
  });

  it('toBuffer writes a ZIP through the lite build', async () => {
    const bytes = await sampleWorkbook().toBuffer();
    expect(bytes.subarray(0, 2)).toEqual(Uint8Array.of(0x50, 0x4b));
  });

  it('validate reports a fresh workbook as valid', () => {
    expect(sampleWorkbook().validate().isValid).toBe(true);
  });

  it('repair returns a lite Workbook', () => {
    const { workbook } = sampleWorkbook().repair();
    expect(workbook).toBeInstanceOf(Workbook);
    expect(workbook.constructor).toBe(Workbook);
  });

  it('readBuffer round-trips the data and returns a lite Workbook', async () => {
    const wb = await readBuffer(await sampleWorkbook().toBuffer());
    expect(wb).toBeInstanceOf(Workbook);
    expect(wb.sheetNames).toEqual(['Data']);
    const ws = wb.getSheet('Data');
    expect(ws?.cell('A1').value).toBe('lite');
    expect(ws?.cell('B1').value).toBe(42);
  });

  it('StreamingXlsxWriter writes through the lite build', () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('S');
    writer.writeRow([{ value: '1', cellType: 'number' }]);
    expect(writer.finish().subarray(0, 2)).toEqual(Uint8Array.of(0x50, 0x4b));
  });

  it('readBuffer with a password throws INVALID_ARGUMENT before any WASM call', async () => {
    const bytes = await sampleWorkbook().toBuffer();
    const err = await readBuffer(bytes, { password: 'x' }).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(ModernXlsxError);
    expect(err).toMatchObject({ code: INVALID_ARGUMENT, message: LITE_PASSWORD_MESSAGE });
  });

  it('readFile with a password throws INVALID_ARGUMENT before reading the file', async () => {
    const err = await readFile('does-not-exist.xlsx', { password: 'x' }).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(ModernXlsxError);
    expect(err).toMatchObject({ code: INVALID_ARGUMENT, message: LITE_PASSWORD_MESSAGE });
  });

  it('toBuffer with a password throws INVALID_ARGUMENT', async () => {
    const err = await sampleWorkbook()
      .toBuffer({ password: 'x' })
      .catch((e: unknown) => e);
    expect(err).toBeInstanceOf(ModernXlsxError);
    expect(err).toMatchObject({ code: INVALID_ARGUMENT, message: LITE_PASSWORD_MESSAGE });
  });

  it('readBuffer classifies non-ZIP input through the lite build', async () => {
    const encrypted = readFileSync(
      new URL(
        '../../../crates/modern-xlsx-core/tests/fixtures/msoffcrypto/example_password.xlsx',
        import.meta.url,
      ),
    );
    await expect(readBuffer(encrypted)).rejects.toMatchObject({
      code: 'PASSWORD_PROTECTED',
      message: expect.stringContaining('lite'),
    });
    await expect(
      readBuffer(buildOle2([{ name: 'Workbook', data: Uint8Array.of(0x78) }])),
    ).rejects.toMatchObject({ code: 'LEGACY_FORMAT' });
    await expect(
      readBuffer(buildOle2([{ name: 'Foo', data: Uint8Array.of(0x78) }])),
    ).rejects.toMatchObject({ code: 'UNRECOGNIZED_FORMAT' });
    await expect(readBuffer(new TextEncoder().encode('garbage'))).rejects.toMatchObject({
      code: 'UNRECOGNIZED_FORMAT',
    });
  });

  it('readBuffer enforces ZIP limits through the lite build', async () => {
    await expect(
      readBuffer(await sampleWorkbook().toBuffer(), { limits: { maxEntries: 1 } }),
    ).rejects.toMatchObject({ code: 'SECURITY' });
  });
});
