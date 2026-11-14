import { INVALID_ARGUMENT, ModernXlsxError } from './model/errors.js';
import type { ReadLimits } from './model/types.js';

/** Fields that must be whole numbers: they count bytes or items. */
const INTEGER_FIELDS: ReadonlySet<string> = new Set([
  'maxDecompressedSize',
  'maxEntries',
  'maxJsonLength',
]);

/** Every field of {@link ReadLimits}. */
const KNOWN_FIELDS: ReadonlySet<string> = new Set([...INTEGER_FIELDS, 'maxCompressionRatio']);

/**
 * Validate `ReadOptions.limits` and serialize the provided fields to the JSON
 * object the WASM `read` functions take.
 *
 * Returns `undefined` when `limits` is `undefined`. Throws `INVALID_ARGUMENT` for an
 * unknown key, or for a value that is not a finite number greater than zero (and,
 * for the byte and item counts, a safe integer).
 */
export function serializeReadLimits(limits: ReadLimits | undefined): string | undefined {
  if (limits === undefined) {
    return undefined;
  }
  const provided: Record<string, number> = {};
  for (const [key, value] of Object.entries(limits)) {
    if (value === undefined) {
      continue;
    }
    const valid =
      KNOWN_FIELDS.has(key) &&
      typeof value === 'number' &&
      Number.isFinite(value) &&
      value > 0 &&
      (!INTEGER_FIELDS.has(key) || Number.isSafeInteger(value));
    if (!valid) {
      throw new ModernXlsxError(
        INVALID_ARGUMENT,
        `Invalid ReadOptions.limits.${key}: ${String(value)}`,
      );
    }
    provided[key] = value;
  }
  return JSON.stringify(provided);
}
