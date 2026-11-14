import { describe, expectTypeOf, it } from 'vitest';
import type { QrEcLevel } from '../src/index.js';

describe('barcode option types', () => {
  it('exports QrEcLevel as the four QR error-correction levels', () => {
    expectTypeOf<QrEcLevel>().toEqualTypeOf<'L' | 'M' | 'Q' | 'H'>();
    expectTypeOf<'L'>().toExtend<QrEcLevel>();
  });
});
