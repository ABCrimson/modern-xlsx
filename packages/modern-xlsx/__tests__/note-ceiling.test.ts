/**
 * The ceiling on the notes one write generates (`MAX_GENERATED_NOTE_SHAPES`
 * in the Rust core, 131,072): a VML note shape for each legacy comment and a
 * `tc=` placeholder for each thread root. Past it the write fails with
 * SIZE_LIMIT before the oversized parts are built.
 */
import { describe, expect, it } from 'vitest';
import { ModernXlsxError, SIZE_LIMIT, Workbook } from '../src/index.js';

const CEILING = 131_072;

describe('generated-note ceiling', () => {
  it('rejects toBuffer() with SIZE_LIMIT for one note past the ceiling', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Notes');
    // No cell values and the shortest note text keep the model small.
    for (let row = 1; row <= CEILING + 1; row++) {
      ws.addComment(`A${row}`, '', '');
    }
    const error: unknown = await wb.toBuffer().catch((e: unknown) => e);
    expect(error).toBeInstanceOf(ModernXlsxError);
    expect(error).toMatchObject({
      code: SIZE_LIMIT,
      message: expect.stringContaining(`more than ${CEILING}`),
    });
  });
});
