import { describe, expectTypeOf, it } from 'vitest';
import type { Workbook } from '../src/index.js';

describe('toBuffer types', () => {
  // TypeScript's Uint8Array defaults to Uint8Array<ArrayBufferLike>, which the DOM and Deno
  // BodyInit/BlobPart types reject; `new Response(await wb.toBuffer())` needs ArrayBuffer.
  it('resolves to bytes backed by a plain ArrayBuffer', () => {
    expectTypeOf<Awaited<ReturnType<Workbook['toBuffer']>>>().toEqualTypeOf<
      Uint8Array<ArrayBuffer>
    >();
  });
});
