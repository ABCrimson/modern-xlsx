import { describe, expect, it } from 'vitest';
import { StreamingXlsxWriter } from '../src/index.js';

describe('StreamingXlsxWriter disposal', () => {
  it('throws INVALID_ARGUMENT after a using scope', () => {
    let writer!: StreamingXlsxWriter;
    {
      using w = StreamingXlsxWriter.create();
      writer = w;
      w.startSheet('S');
    }
    expect(() => writer.startSheet('T')).toThrow(
      expect.objectContaining({
        code: 'INVALID_ARGUMENT',
        message: 'StreamingXlsxWriter has been disposed',
      }),
    );
    expect(() => writer.writeRow([])).toThrow(
      expect.objectContaining({ code: 'INVALID_ARGUMENT' }),
    );
  });

  it('is idempotent when disposed twice', () => {
    const writer = StreamingXlsxWriter.create();
    writer[Symbol.dispose]();
    expect(() => writer[Symbol.dispose]()).not.toThrow();
  });

  it('allows dispose after finish', () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('S');
    writer.finish();
    expect(() => writer[Symbol.dispose]()).not.toThrow();
  });
});
