import { describe, expectTypeOf, it } from 'vitest';
import {
  type KnownModernXlsxErrorCode,
  type ModernXlsxError,
  type ModernXlsxErrorCode,
  WORKER_ERROR,
} from '../src/model/errors.js';

describe('error code types', () => {
  it('types the code constants and the code union', () => {
    expectTypeOf(WORKER_ERROR).toEqualTypeOf<'WORKER_ERROR'>();
    expectTypeOf<'PASSWORD_PROTECTED'>().toExtend<KnownModernXlsxErrorCode>();
    expectTypeOf<'WORKER_ERROR'>().toExtend<KnownModernXlsxErrorCode>();
    expectTypeOf<ModernXlsxError['code']>().toEqualTypeOf<ModernXlsxErrorCode>();
    expectTypeOf<'FUTURE_CODE'>().toExtend<ModernXlsxErrorCode>();
  });
});
