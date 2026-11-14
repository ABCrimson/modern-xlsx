/**
 * String and logical formula functions.
 *
 * Registers the built-in Excel functions covering string manipulation
 * (CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, TEXT, VALUE,
 * EXACT, SUBSTITUTE, REPT, FIND, SEARCH) and logical evaluation
 * (IF, AND, OR, NOT, IFERROR).
 *
 * TEXT renders through `formatCell` (honouring `ctx.dateSystem`); TRIM touches
 * only U+0020; AND/OR evaluate every argument and fold reference/array cells
 * (blanks and text skipped); SEARCH supports `?`, `*` and `~` wildcards while
 * FIND stays literal and case-sensitive; SUBSTITUTE's nth instance counts
 * non-overlapping matches; CONCATENATE needs at least one argument; an omitted
 * IF/IFERROR result argument yields 0. Text results honour Excel's 32,767
 * character cap: CONCATENATE keeps the first 32,767 characters, while REPT and
 * SUBSTITUTE give `#VALUE!` past it.
 *
 * @module formula/functions/string-logical
 */

import { formatCell } from '../../format/format-cell.js';
import { coerceToString, MAX_TEXT_LENGTH, toBool, toNumber } from '../coerce.js';
import { isError } from '../errors.js';
import type { ASTNode } from '../parser.js';
import type { CellValue, EvalContext, EvaluateFn, FormulaFunction } from '../resolver.js';
import { argValues, isReferenceNode } from '../resolver.js';
import { searchGlob } from './glob.js';

// ---------------------------------------------------------------------------
// Function implementations
// ---------------------------------------------------------------------------

/** Resolve an optional start-position argument for FIND/SEARCH (1-based). */
function resolveStartPos(
  args: ASTNode[],
  argIndex: number,
  ctx: EvalContext,
  evaluate: EvaluateFn,
): number | string {
  if (args.length <= argIndex) return 0;
  const arg = args[argIndex];
  if (!arg) return 0;
  const sVal = toNumber(evaluate(arg, ctx));
  if (typeof sVal === 'string') return sVal;
  if (sVal < 1) return '#VALUE!';
  return Math.floor(sVal) - 1;
}

/** Replace the nth non-overlapping occurrence of `oldStr` in `text` with `newStr` (each match resumes the scan after its end, as Excel and LibreOffice count); `text` is returned unchanged when there are fewer than `instance` matches, and `#VALUE!` when the result would exceed 32,767 characters. */
function replaceNthOccurrence(
  text: string,
  oldStr: string,
  newStr: string,
  instance: number,
): string {
  let count = 0;
  let pos = 0;
  while (pos < text.length) {
    const idx = text.indexOf(oldStr, pos);
    if (idx === -1) break;
    count++;
    if (count === instance) {
      if (text.length - oldStr.length + newStr.length > MAX_TEXT_LENGTH) return '#VALUE!';
      return text.slice(0, idx) + newStr + text.slice(idx + oldStr.length);
    }
    pos = idx + oldStr.length;
  }
  return text;
}

/**
 * Fold the arguments of AND/OR. Every argument is evaluated left to right
 * (no early return), so the first error anywhere wins. Reference and array
 * arguments contribute their cells: blanks and non-error text are skipped,
 * booleans count as themselves, numbers as `n !== 0`. Any other argument is
 * a scalar read with {@link toBool} (a `'missing'` argument is FALSE). When no
 * logical value was seen at all the result is `#VALUE!`.
 */
function foldLogical(
  args: ASTNode[],
  ctx: EvalContext,
  evaluate: EvaluateFn,
  isAnd: boolean,
): CellValue {
  if (args.length === 0) return '#VALUE!';
  let firstError: string | undefined;
  let seen = false;
  let result = isAnd;
  const take = (b: boolean): void => {
    seen = true;
    result = isAnd ? result && b : result || b;
  };
  for (const a of args) {
    if (a.type === 'missing') {
      take(false);
    } else if (a.type === 'array' || isReferenceNode(a)) {
      for (const v of argValues(a, ctx, evaluate)) {
        if (v === null) continue;
        if (typeof v === 'string') {
          if (isError(v)) firstError ??= v;
          continue;
        }
        take(typeof v === 'boolean' ? v : v !== 0);
      }
    } else {
      const v = evaluate(a, ctx);
      if (isError(v)) {
        firstError ??= v;
        continue;
      }
      const b = toBool(v);
      if (typeof b === 'string') firstError ??= b;
      else take(b);
    }
  }
  if (firstError !== undefined) return firstError;
  return seen ? result : '#VALUE!';
}

/**
 * SUBSTITUTE(text, old_text, new_text, [instance_num]): replaces `old_text`
 * with `new_text` inside `text`, case-sensitively and literally.
 *
 * An empty `old_text` returns `text` unchanged. Without `instance_num` every
 * occurrence is replaced. With it, only that occurrence is replaced;
 * `instance_num` is floored and must be at least 1, else `#VALUE!`. A result
 * longer than 32,767 characters (Excel's text-length cap) is `#VALUE!`, checked
 * before the result is built. Fewer than three arguments gives `#VALUE!`; an
 * error in `text`, `old_text` or `new_text` is returned as is, and an error in
 * `instance_num` too unless `old_text` is empty (that argument is then never
 * evaluated).
 */
const substituteImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 3) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  const arg2 = args[2];
  if (!arg0 || !arg1 || !arg2) return '#VALUE!';

  const textVal = evaluate(arg0, ctx);
  if (isError(textVal)) return textVal;
  const oldVal = evaluate(arg1, ctx);
  if (isError(oldVal)) return oldVal;
  const newVal = evaluate(arg2, ctx);
  if (isError(newVal)) return newVal;

  const text = coerceToString(textVal);
  const oldStr = coerceToString(oldVal);
  const newStr = coerceToString(newVal);

  if (oldStr === '') return text;

  if (args.length >= 4) {
    const arg3 = args[3];
    if (!arg3) return '#VALUE!';
    const instVal = toNumber(evaluate(arg3, ctx));
    if (typeof instVal === 'string') return instVal;
    const instance = Math.floor(instVal);
    if (instance < 1) return '#VALUE!';
    return replaceNthOccurrence(text, oldStr, newStr, instance);
  }

  // Count the matches first so an over-long result is never allocated.
  let count = 0;
  for (let i = text.indexOf(oldStr); i !== -1; i = text.indexOf(oldStr, i + oldStr.length)) {
    count++;
  }
  if (text.length + count * (newStr.length - oldStr.length) > MAX_TEXT_LENGTH) return '#VALUE!';
  return text.split(oldStr).join(newStr);
};

/**
 * REPT(text, number_times): repeats `text` `number_times` times.
 *
 * `number_times` is floored. A negative or non-finite count gives `#VALUE!`,
 * as does a result longer than 32,767 characters (Excel's text-length cap).
 */
const reptImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 2) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  if (!arg0 || !arg1) return '#VALUE!';

  const textVal = evaluate(arg0, ctx);
  if (isError(textVal)) return textVal;
  const nVal = toNumber(evaluate(arg1, ctx));
  if (typeof nVal === 'string') return nVal;
  if (nVal < 0) return '#VALUE!';
  const count = Math.floor(nVal);
  if (!Number.isFinite(count)) return '#VALUE!';
  const text = coerceToString(textVal);
  // Excel caps a text result at 32,767 characters.
  if (text.length * count > MAX_TEXT_LENGTH) return '#VALUE!';
  return text.repeat(count);
};

/**
 * FIND(find_text, within_text, [start_num]): the 1-based position of the first
 * case-sensitive, literal (no wildcards) match of `find_text` in `within_text`.
 *
 * The search begins at `start_num` (default 1, floored). A `start_num` below 1
 * or past the text length + 1 gives `#VALUE!`, as does no match; so an empty
 * `find_text` matches at `start_num` itself up to the text length + 1, as in
 * Excel (`FIND("","abc",4)` is 4, `FIND("","abc",5)` is `#VALUE!`).
 */
const findImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 2) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  if (!arg0 || !arg1) return '#VALUE!';

  const findVal = evaluate(arg0, ctx);
  if (isError(findVal)) return findVal;
  const withinVal = evaluate(arg1, ctx);
  if (isError(withinVal)) return withinVal;

  const startPos = resolveStartPos(args, 2, ctx, evaluate);
  if (typeof startPos === 'string') return startPos;

  const within = coerceToString(withinVal);
  // String#indexOf clamps fromIndex to the length; startPos is 0-based.
  if (startPos > within.length) return '#VALUE!';
  const idx = within.indexOf(coerceToString(findVal), startPos);
  return idx === -1 ? '#VALUE!' : idx + 1;
};

/**
 * SEARCH(find_text, within_text, [start_num]): the 1-based position of the
 * first case-insensitive match of `find_text` in `within_text`.
 *
 * `?` matches any single character, `*` any run of characters, and `~` escapes
 * a following `?`, `*` or `~`. `start_num` defaults to 1 and is floored; below
 * 1, or past the end of the text, gives `#VALUE!`, as does no match.
 */
const searchImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 2) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  if (!arg0 || !arg1) return '#VALUE!';

  const findVal = evaluate(arg0, ctx);
  if (isError(findVal)) return findVal;
  const withinVal = evaluate(arg1, ctx);
  if (isError(withinVal)) return withinVal;

  const startPos = resolveStartPos(args, 2, ctx, evaluate);
  if (typeof startPos === 'string') return startPos;

  const within = coerceToString(withinVal);
  // start_num (1-based) past the end of the text is #VALUE!; startPos is 0-based.
  if (startPos >= within.length && startPos > 0) return '#VALUE!';
  const idx = searchGlob(coerceToString(findVal), within, startPos);
  return idx === -1 ? '#VALUE!' : idx + 1;
};

// ---------------------------------------------------------------------------
// Registration
// ---------------------------------------------------------------------------

/**
 * Register all string and logical functions into the given registry.
 */
export function registerStringLogicalFunctions(registry: Map<string, FormulaFunction>): void {
  // ---- IF ----------------------------------------------------------------
  registry.set('IF', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = args[0];
    const arg1 = args[1];
    if (!arg0 || !arg1) return '#VALUE!';
    const test = evaluate(arg0, ctx);
    if (isError(test)) return test;
    const condition = toBool(test);
    if (typeof condition === 'string') return condition;
    const chosen = condition ? arg1 : args[2];
    if (!chosen) return false;
    // An omitted argument (`IF(TRUE,,2)`) yields 0, not blank.
    if (chosen.type === 'missing') return 0;
    return evaluate(chosen, ctx);
  });

  // ---- AND ---------------------------------------------------------------
  registry.set('AND', (args, ctx, evaluate): CellValue => foldLogical(args, ctx, evaluate, true));

  // ---- OR ----------------------------------------------------------------
  registry.set('OR', (args, ctx, evaluate): CellValue => foldLogical(args, ctx, evaluate, false));

  // ---- NOT ---------------------------------------------------------------
  registry.set('NOT', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const val = evaluate(arg0, ctx);
    if (isError(val)) return val;
    const b = toBool(val);
    if (typeof b === 'string') return b;
    return !b;
  });

  // ---- IFERROR -----------------------------------------------------------
  registry.set('IFERROR', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = args[0];
    const arg1 = args[1];
    if (!arg0 || !arg1) return '#VALUE!';
    const val = evaluate(arg0, ctx);
    if (!isError(val)) return val;
    // An omitted fallback (`IFERROR(1/0,)`) yields 0, not blank.
    return arg1.type === 'missing' ? 0 : evaluate(arg1, ctx);
  });

  // ---- CONCATENATE -------------------------------------------------------
  registry.set('CONCATENATE', (args, ctx, evaluate): CellValue => {
    if (args.length === 0) return '#VALUE!';
    // Excel keeps the first 32,767 characters; every argument is still
    // evaluated so a later error propagates.
    let result = '';
    for (const a of args) {
      const val = evaluate(a, ctx);
      if (isError(val)) return val;
      const room = MAX_TEXT_LENGTH - result.length;
      if (room > 0) result += coerceToString(val).slice(0, room);
    }
    return result;
  });

  // ---- LEFT --------------------------------------------------------------
  registry.set('LEFT', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const textVal = evaluate(arg0, ctx);
    if (isError(textVal)) return textVal;
    const text = coerceToString(textVal);
    let n = 1;
    if (args.length >= 2) {
      const arg1 = args[1];
      if (!arg1) return '#VALUE!';
      const nVal = toNumber(evaluate(arg1, ctx));
      if (typeof nVal === 'string') return nVal;
      if (nVal < 0) return '#VALUE!';
      n = Math.floor(nVal);
    }
    return text.slice(0, n);
  });

  // ---- RIGHT -------------------------------------------------------------
  registry.set('RIGHT', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const textVal = evaluate(arg0, ctx);
    if (isError(textVal)) return textVal;
    const text = coerceToString(textVal);
    let n = 1;
    if (args.length >= 2) {
      const arg1 = args[1];
      if (!arg1) return '#VALUE!';
      const nVal = toNumber(evaluate(arg1, ctx));
      if (typeof nVal === 'string') return nVal;
      if (nVal < 0) return '#VALUE!';
      n = Math.floor(nVal);
    }
    if (n === 0) return '';
    return text.slice(-n);
  });

  // ---- MID ---------------------------------------------------------------
  registry.set('MID', (args, ctx, evaluate): CellValue => {
    if (args.length < 3) return '#VALUE!';
    const arg0 = args[0];
    const arg1 = args[1];
    const arg2 = args[2];
    if (!arg0 || !arg1 || !arg2) return '#VALUE!';
    const textVal = evaluate(arg0, ctx);
    if (isError(textVal)) return textVal;
    const text = coerceToString(textVal);
    const startVal = toNumber(evaluate(arg1, ctx));
    if (typeof startVal === 'string') return startVal;
    const nVal = toNumber(evaluate(arg2, ctx));
    if (typeof nVal === 'string') return nVal;
    if (startVal < 1 || nVal < 0) return '#VALUE!';
    const start = Math.floor(startVal) - 1;
    return text.slice(start, start + Math.floor(nVal));
  });

  // ---- LEN ---------------------------------------------------------------
  registry.set('LEN', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const val = evaluate(arg0, ctx);
    if (isError(val)) return val;
    return coerceToString(val).length;
  });

  // ---- TRIM --------------------------------------------------------------
  registry.set('TRIM', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const val = evaluate(arg0, ctx);
    if (isError(val)) return val;
    // Only U+0020 is trimmed or collapsed; tabs, NBSP and other spaces stay.
    // The collapse runs first, so no regex rescans a long run of spaces.
    let text = coerceToString(val).replace(/ {2,}/g, ' ');
    if (text.startsWith(' ')) text = text.slice(1);
    if (text.endsWith(' ')) text = text.slice(0, -1);
    return text;
  });

  // ---- UPPER -------------------------------------------------------------
  registry.set('UPPER', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const val = evaluate(arg0, ctx);
    if (isError(val)) return val;
    return coerceToString(val).toUpperCase();
  });

  // ---- LOWER -------------------------------------------------------------
  registry.set('LOWER', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const val = evaluate(arg0, ctx);
    if (isError(val)) return val;
    return coerceToString(val).toLowerCase();
  });

  // ---- TEXT --------------------------------------------------------------
  registry.set('TEXT', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = args[0];
    const arg1 = args[1];
    if (!arg0 || !arg1) return '#VALUE!';
    const valRaw = evaluate(arg0, ctx);
    if (isError(valRaw)) return valRaw;
    const fmtRaw = evaluate(arg1, ctx);
    if (isError(fmtRaw)) return fmtRaw;
    const fmt = coerceToString(fmtRaw);
    let value: string | number | boolean;
    if (valRaw === null) value = 0;
    else if (typeof valRaw === 'string') {
      // Numeric text is formatted as a number; other text passes through.
      const n = toNumber(valRaw);
      value = typeof n === 'number' ? n : valRaw;
    } else value = valRaw;
    return formatCell(
      value,
      fmt,
      ctx.dateSystem === undefined ? undefined : { dateSystem: ctx.dateSystem },
    );
  });

  // ---- VALUE -------------------------------------------------------------
  registry.set('VALUE', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const val = evaluate(arg0, ctx);
    if (isError(val)) return val;
    if (typeof val === 'number') return val;
    if (typeof val === 'boolean') return val ? 1 : 0;
    if (val === null) return 0;
    return toNumber(val);
  });

  // ---- EXACT -------------------------------------------------------------
  registry.set('EXACT', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = args[0];
    const arg1 = args[1];
    if (!arg0 || !arg1) return '#VALUE!';
    const a = evaluate(arg0, ctx);
    if (isError(a)) return a;
    const b = evaluate(arg1, ctx);
    if (isError(b)) return b;
    return coerceToString(a) === coerceToString(b);
  });

  registry.set('SUBSTITUTE', substituteImpl);
  registry.set('REPT', reptImpl);
  registry.set('FIND', findImpl);
  registry.set('SEARCH', searchImpl);
}
