import { describe, expect, it } from 'vitest';
import * as full from '../src/index.js';
import { dateToSerial } from '../src/index.js';
import * as lite from '../src/index-lite.js';
import { StreamingXlsxWriter as StreamingXlsxWriterBase } from '../src/io/streaming-writer.js';
import { ModernXlsxError, WASM_ERROR, WASM_INIT_FAILED } from '../src/model/errors.js';
import { Workbook as WorkbookBase } from '../src/model/workbook.js';
import { createLoader, type StreamingWriterHandle, type WasmGlue } from '../src/wasm-backend.js';

/** Runtime exports only the full entry has: barcodes and the worker client. */
const FULL_ONLY = [
  'createXlsxWorker',
  'encodeCode39',
  'encodeCode128',
  'encodeDataMatrix',
  'encodeEAN13',
  'encodeGS1128',
  'encodeITF14',
  'encodePDF417',
  'encodeQR',
  'encodeUPCA',
  'generateBarcode',
  'generateDrawingRels',
  'generateDrawingXml',
  'renderBarcodePNG',
].sort();

class FakeStreamingWriter implements StreamingWriterHandle {
  free(): void {}
  [Symbol.dispose](): void {}
  finish(): Uint8Array {
    return new Uint8Array();
  }
  setStylesXml(): void {}
  startSheet(): void {}
  writeRow(): void {}
}

/** A glue stub whose read/validate/repair return `'{}'`, which no guard accepts. */
function fakeGlue(overrides: Partial<WasmGlue> = {}): WasmGlue {
  return {
    default: async () => ({}),
    initSync: () => ({}),
    read: () => '{}',
    write: () => new Uint8Array(),
    writeBlob: () => new Blob(),
    validate: () => '{}',
    repair: () => '{}',
    version: () => '0.0.0-fake',
    StreamingWriter: FakeStreamingWriter,
    ...overrides,
  };
}

describe('entry exports', () => {
  it('lite exports every full export except the documented full-only set', () => {
    const fullKeys = Object.keys(full).sort();
    const liteKeys = Object.keys(lite).sort();
    expect(fullKeys.filter((k) => !(k in lite))).toEqual(FULL_ONLY);
    expect(liteKeys).toEqual(fullKeys.filter((k) => !FULL_ONLY.includes(k)));
  });

  it('shared exports are the same values in both entries', () => {
    expect(lite.Worksheet).toBe(full.Worksheet);
    expect(lite.PivotTableBuilder).toBe(full.PivotTableBuilder);
    expect(lite.VERSION).toBe(full.VERSION);
  });

  it('each entry binds its own Workbook and StreamingXlsxWriter subclass', () => {
    expect(full.Workbook).not.toBe(lite.Workbook);
    expect(Object.getPrototypeOf(full.Workbook)).toBe(WorkbookBase);
    expect(Object.getPrototypeOf(lite.Workbook)).toBe(WorkbookBase);
    expect(full.StreamingXlsxWriter).not.toBe(lite.StreamingXlsxWriter);
    expect(Object.getPrototypeOf(full.StreamingXlsxWriter)).toBe(StreamingXlsxWriterBase);
    expect(Object.getPrototypeOf(lite.StreamingXlsxWriter)).toBe(StreamingXlsxWriterBase);
  });

  it('the unbound base Workbook rejects toBuffer with WASM_INIT_FAILED', async () => {
    const err = await new WorkbookBase().toBuffer().catch((e: unknown) => e);
    expect(err).toBeInstanceOf(ModernXlsxError);
    expect(err).toMatchObject({ code: WASM_INIT_FAILED });
  });

  it('the unbound base StreamingXlsxWriter.create throws WASM_INIT_FAILED', () => {
    expect(() => StreamingXlsxWriterBase.create()).toThrow(
      expect.objectContaining({ code: WASM_INIT_FAILED }),
    );
  });
});

describe('createLoader', () => {
  it('throws WASM_ERROR when the glue returns an invalid structure', () => {
    const loader = createLoader(fakeGlue(), 'full');
    loader.initWasmSync(new Uint8Array());
    for (const call of [
      () => loader.wasmRead(new Uint8Array()),
      () => loader.wasmValidate(new WorkbookBase().toJSON()),
      () => loader.wasmRepair(new WorkbookBase().toJSON()),
    ]) {
      expect(call).toThrow(expect.objectContaining({ name: 'ModernXlsxError', code: WASM_ERROR }));
    }
  });

  it('wraps an init rejection in WASM_INIT_FAILED and keeps the cause', async () => {
    const cause = new TypeError('fetch failed');
    const loader = createLoader(
      fakeGlue({
        default: async () => {
          throw cause;
        },
      }),
      'full',
    );
    const err = await loader.initWasm('https://example.invalid/x.wasm').catch((e: unknown) => e);
    expect(err).toBeInstanceOf(ModernXlsxError);
    expect(err).toMatchObject({ code: WASM_INIT_FAILED, message: 'fetch failed', cause });
    expect(() => loader.ensureInitialized()).toThrow(
      expect.objectContaining({ code: WASM_INIT_FAILED }),
    );
  });
});

describe('TemporalLike', () => {
  it.skipIf(typeof Temporal === 'undefined')('dateToSerial takes a native PlainDateTime', () => {
    expect(dateToSerial(Temporal.PlainDateTime.from('2026-03-01T12:00'))).toBe(46082.5);
  });
});
