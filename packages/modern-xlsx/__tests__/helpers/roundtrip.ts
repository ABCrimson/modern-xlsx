import type { ReadOptions, Workbook, WriteOptions } from '../../src/index.js';
import { readBuffer } from '../../src/index.js';

/**
 * Writes `wb` with `toBuffer` and reads the bytes straight back. `opts` feeds
 * both calls, so `{ password }` encrypts on the way out and decrypts on the way in.
 */
export async function roundTrip(
  wb: Workbook,
  opts?: WriteOptions & ReadOptions,
): Promise<Workbook> {
  return readBuffer(await wb.toBuffer(opts), opts);
}
