import vm from 'node:vm';
import { describe, expect, it } from 'vitest';
import { decodeCellRef, decodeRange, decodeRow, splitCellRef } from '../src/format/cell-ref.js';
import * as full from '../src/index.js';
import * as lite from '../src/index-lite.js';
import {
  INVALID_ARGUMENT,
  INVALID_CELL_REF,
  ModernXlsxError,
  WASM_ERROR,
} from '../src/model/errors.js';

describe('ModernXlsxError', () => {
  it('keeps the cause passed through ErrorOptions', () => {
    const cause = new Error('root');
    const err = new ModernXlsxError('X', 'm', { cause });
    expect(err.cause).toBe(cause);
    expect(err.name).toBe('ModernXlsxError');
    expect(err.code).toBe('X');
    expect(err.message).toBe('m');
  });

  it('has no cause when no options are given', () => {
    expect('cause' in new ModernXlsxError('X', 'm')).toBe(false);
  });
});

describe('ModernXlsxError.fromWasmError', () => {
  it('parses the [CODE] prefix and keeps the original error as cause', () => {
    const original = new Error('[PASSWORD_PROTECTED] Incorrect password.');
    const err = ModernXlsxError.fromWasmError(original);
    expect(err.code).toBe('PASSWORD_PROTECTED');
    expect(err.message).toBe('Incorrect password.');
    expect(err.cause).toBe(original);
  });

  it('falls back to WASM_ERROR for a plain string and keeps it as cause', () => {
    const err = ModernXlsxError.fromWasmError('plain text');
    expect(err.code).toBe(WASM_ERROR);
    expect(err.message).toBe('plain text');
    expect(err.cause).toBe('plain text');
  });

  it('parses a coded message that spans lines', () => {
    const err = ModernXlsxError.fromWasmError(new Error('[ZIP_READ] a\nb'));
    expect(err.code).toBe('ZIP_READ');
    expect(err.message).toBe('a\nb');
  });

  it('recognises errors from another realm', () => {
    const foreign = vm.runInNewContext('new Error("[ZIP_READ] bad")') as unknown;
    expect(foreign instanceof Error).toBe(false);
    const err = ModernXlsxError.fromWasmError(foreign);
    expect(err.code).toBe('ZIP_READ');
    expect(err.message).toBe('bad');
    expect(err.cause).toBe(foreign);
  });
});

describe('cell-reference codec errors', () => {
  function thrown(fn: () => unknown): ModernXlsxError {
    try {
      fn();
    } catch (e) {
      expect(e).toBeInstanceOf(ModernXlsxError);
      return e as ModernXlsxError;
    }
    throw new Error('did not throw');
  }

  it('decodeCellRef, splitCellRef and decodeRange throw INVALID_CELL_REF', () => {
    const a = thrown(() => decodeCellRef(''));
    expect(a.code).toBe(INVALID_CELL_REF);
    expect(a.message).toBe('Invalid cell reference: ');
    const b = thrown(() => splitCellRef('!!'));
    expect(b.code).toBe(INVALID_CELL_REF);
    expect(b.message).toBe('Invalid cell reference: !!');
    const c = thrown(() => decodeRange('A1:'));
    expect(c.code).toBe(INVALID_CELL_REF);
    expect(c.message).toBe('Invalid range: A1:');
  });

  it('decodeRow throws INVALID_ARGUMENT', () => {
    const e = thrown(() => decodeRow('x'));
    expect(e.code).toBe(INVALID_ARGUMENT);
    expect(e.message).toBe('Invalid row string: x');
  });
});

describe('error code exports', () => {
  it('both entries export WORKER_ERROR', () => {
    expect(full.WORKER_ERROR).toBe('WORKER_ERROR');
    expect(lite.WORKER_ERROR).toBe('WORKER_ERROR');
  });
});
