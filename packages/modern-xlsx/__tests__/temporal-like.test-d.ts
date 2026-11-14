import { describe, expectTypeOf, it } from 'vitest';
import type { TemporalLike } from '../src/index.js';

describe('TemporalLike', () => {
  it('accepts native Temporal.PlainDate and PlainDateTime', () => {
    expectTypeOf<Temporal.PlainDate>().toExtend<TemporalLike>();
    expectTypeOf<Temporal.PlainDateTime>().toExtend<TemporalLike>();
  });
});
