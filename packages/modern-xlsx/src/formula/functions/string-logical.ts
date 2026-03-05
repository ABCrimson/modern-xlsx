/**
 * String and logical formula functions.
 *
 * Registers 20 built-in Excel functions covering string manipulation
 * (CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, TEXT, VALUE,
 * EXACT, SUBSTITUTE, REPT, FIND, SEARCH) and logical evaluation
 * (IF, AND, OR, NOT, IFERROR).
 *
 * @module formula/functions/string-logical
 */

import type { ASTNode } from '../parser.js';
import type { CellValue, EvalContext, FormulaFunction } from '../resolver.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Safely get an element from an array by index. */
function at<T>(arr: T[], idx: number): T | undefined {
  return arr[idx];
}

function isError(val: CellValue): val is string {
  return typeof val === 'string' && val.length > 0 && val.charAt(0) === '#';
}

function toNumber(val: CellValue): number | string {
  if (typeof val === 'number') return val;
  if (typeof val === 'boolean') return val ? 1 : 0;
  if (val === null) return 0;
  if (isError(val)) return val;
  const n = Number(val);
  return Number.isNaN(n) ? '#VALUE!' : n;
}

function coerceToString(val: CellValue): string {
  if (val === null) return '';
  if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE';
  return String(val);
}

function toBool(val: CellValue): boolean | string {
  if (typeof val === 'boolean') return val;
  if (typeof val === 'number') return val !== 0;
  if (val === null) return false;
  if (typeof val === 'string') {
    if (val.length > 0 && val.charAt(0) === '#') return val;
    const upper = val.toUpperCase();
    if (upper === 'TRUE') return true;
    if (upper === 'FALSE') return false;
    return '#VALUE!';
  }
  return '#VALUE!';
}

// ---------------------------------------------------------------------------
// Extracted implementations (reduce cognitive complexity)
// ---------------------------------------------------------------------------

/** Resolve an optional start-position argument for FIND/SEARCH (1-based). */
function resolveStartPos(
  args: ASTNode[],
  argIndex: number,
  ctx: EvalContext,
  evaluate: (node: ASTNode, ctx: EvalContext) => CellValue,
): number | string {
  if (args.length <= argIndex) return 0;
  const arg = at(args, argIndex);
  if (!arg) return 0;
  const sVal = toNumber(evaluate(arg, ctx));
  if (typeof sVal === 'string') return sVal;
  if (sVal < 1) return '#VALUE!';
  return Math.floor(sVal) - 1;
}

/** Replace the nth occurrence of `oldStr` in `text` with `newStr`. */
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
      return text.slice(0, idx) + newStr + text.slice(idx + oldStr.length);
    }
    pos = idx + 1;
  }
  return text;
}

const substituteImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 3) return '#VALUE!';
  const arg0 = at(args, 0);
  const arg1 = at(args, 1);
  const arg2 = at(args, 2);
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
    const arg3 = at(args, 3);
    if (!arg3) return '#VALUE!';
    const instVal = toNumber(evaluate(arg3, ctx));
    if (typeof instVal === 'string') return instVal;
    const instance = Math.floor(instVal);
    if (instance < 1) return '#VALUE!';
    return replaceNthOccurrence(text, oldStr, newStr, instance);
  }

  return text.split(oldStr).join(newStr);
};

const reptImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 2) return '#VALUE!';
  const arg0 = at(args, 0);
  const arg1 = at(args, 1);
  if (!arg0 || !arg1) return '#VALUE!';

  const textVal = evaluate(arg0, ctx);
  if (isError(textVal)) return textVal;
  const nVal = toNumber(evaluate(arg1, ctx));
  if (typeof nVal === 'string') return nVal;
  if (nVal < 0) return '#VALUE!';
  return coerceToString(textVal).repeat(Math.floor(nVal));
};

const findImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 2) return '#VALUE!';
  const arg0 = at(args, 0);
  const arg1 = at(args, 1);
  if (!arg0 || !arg1) return '#VALUE!';

  const findVal = evaluate(arg0, ctx);
  if (isError(findVal)) return findVal;
  const withinVal = evaluate(arg1, ctx);
  if (isError(withinVal)) return withinVal;

  const startPos = resolveStartPos(args, 2, ctx, evaluate);
  if (typeof startPos === 'string') return startPos;

  const idx = coerceToString(withinVal).indexOf(coerceToString(findVal), startPos);
  return idx === -1 ? '#VALUE!' : idx + 1;
};

const searchImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 2) return '#VALUE!';
  const arg0 = at(args, 0);
  const arg1 = at(args, 1);
  if (!arg0 || !arg1) return '#VALUE!';

  const findVal = evaluate(arg0, ctx);
  if (isError(findVal)) return findVal;
  const withinVal = evaluate(arg1, ctx);
  if (isError(withinVal)) return withinVal;

  const startPos = resolveStartPos(args, 2, ctx, evaluate);
  if (typeof startPos === 'string') return startPos;

  const idx = coerceToString(withinVal)
    .toLowerCase()
    .indexOf(coerceToString(findVal).toLowerCase(), startPos);
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
    const arg0 = at(args, 0);
    const arg1 = at(args, 1);
    if (!arg0 || !arg1) return '#VALUE!';
    const test = evaluate(arg0, ctx);
    if (isError(test)) return test;
    const condition = toBool(test);
    if (typeof condition === 'string') return condition;
    if (condition) {
      return evaluate(arg1, ctx);
    }
    return args[2] ? evaluate(args[2], ctx) : false;
  });

  // ---- AND ---------------------------------------------------------------
  registry.set('AND', (args, ctx, evaluate): CellValue => {
    if (args.length === 0) return '#VALUE!';
    for (const a of args) {
      const val = evaluate(a, ctx);
      if (isError(val)) return val;
      const b = toBool(val);
      if (typeof b === 'string') return b;
      if (!b) return false;
    }
    return true;
  });

  // ---- OR ----------------------------------------------------------------
  registry.set('OR', (args, ctx, evaluate): CellValue => {
    if (args.length === 0) return '#VALUE!';
    for (const a of args) {
      const val = evaluate(a, ctx);
      if (isError(val)) return val;
      const b = toBool(val);
      if (typeof b === 'string') return b;
      if (b) return true;
    }
    return false;
  });

  // ---- NOT ---------------------------------------------------------------
  registry.set('NOT', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = at(args, 0);
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
    const arg0 = at(args, 0);
    const arg1 = at(args, 1);
    if (!arg0 || !arg1) return '#VALUE!';
    const val = evaluate(arg0, ctx);
    if (isError(val)) return evaluate(arg1, ctx);
    return val;
  });

  // ---- CONCATENATE -------------------------------------------------------
  registry.set('CONCATENATE', (args, ctx, evaluate): CellValue => {
    let result = '';
    for (const a of args) {
      const val = evaluate(a, ctx);
      if (isError(val)) return val;
      result += coerceToString(val);
    }
    return result;
  });

  // ---- LEFT --------------------------------------------------------------
  registry.set('LEFT', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = at(args, 0);
    if (!arg0) return '#VALUE!';
    const textVal = evaluate(arg0, ctx);
    if (isError(textVal)) return textVal;
    const text = coerceToString(textVal);
    let n = 1;
    if (args.length >= 2) {
      const arg1 = at(args, 1);
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
    const arg0 = at(args, 0);
    if (!arg0) return '#VALUE!';
    const textVal = evaluate(arg0, ctx);
    if (isError(textVal)) return textVal;
    const text = coerceToString(textVal);
    let n = 1;
    if (args.length >= 2) {
      const arg1 = at(args, 1);
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
    const arg0 = at(args, 0);
    const arg1 = at(args, 1);
    const arg2 = at(args, 2);
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
    const arg0 = at(args, 0);
    if (!arg0) return '#VALUE!';
    const val = evaluate(arg0, ctx);
    if (isError(val)) return val;
    return coerceToString(val).length;
  });

  // ---- TRIM --------------------------------------------------------------
  registry.set('TRIM', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = at(args, 0);
    if (!arg0) return '#VALUE!';
    const val = evaluate(arg0, ctx);
    if (isError(val)) return val;
    return coerceToString(val).trim().replace(/ +/g, ' ');
  });

  // ---- UPPER -------------------------------------------------------------
  registry.set('UPPER', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = at(args, 0);
    if (!arg0) return '#VALUE!';
    const val = evaluate(arg0, ctx);
    if (isError(val)) return val;
    return coerceToString(val).toUpperCase();
  });

  // ---- LOWER -------------------------------------------------------------
  registry.set('LOWER', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = at(args, 0);
    if (!arg0) return '#VALUE!';
    const val = evaluate(arg0, ctx);
    if (isError(val)) return val;
    return coerceToString(val).toLowerCase();
  });

  // ---- TEXT --------------------------------------------------------------
  registry.set('TEXT', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = at(args, 0);
    const arg1 = at(args, 1);
    if (!arg0 || !arg1) return '#VALUE!';
    const valRaw = evaluate(arg0, ctx);
    if (isError(valRaw)) return valRaw;
    const fmtRaw = evaluate(arg1, ctx);
    if (isError(fmtRaw)) return fmtRaw;
    const num = toNumber(valRaw);
    if (typeof num === 'string') return num;
    const fmt = coerceToString(fmtRaw);
    const decMatch = fmt.match(/\.(0+)/);
    if (decMatch) {
      return num.toFixed(decMatch[1]?.length ?? 0);
    }
    if (fmt.includes('%')) {
      const pctMatch = fmt.match(/\.(0+)%/);
      if (pctMatch) {
        return `${(num * 100).toFixed(pctMatch[1]?.length ?? 0)}%`;
      }
      return `${Math.round(num * 100)}%`;
    }
    return String(num);
  });

  // ---- VALUE -------------------------------------------------------------
  registry.set('VALUE', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = at(args, 0);
    if (!arg0) return '#VALUE!';
    const val = evaluate(arg0, ctx);
    if (isError(val)) return val;
    if (typeof val === 'number') return val;
    if (typeof val === 'boolean') return val ? 1 : 0;
    if (val === null) return 0;
    const n = Number(val);
    return Number.isNaN(n) ? '#VALUE!' : n;
  });

  // ---- EXACT -------------------------------------------------------------
  registry.set('EXACT', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = at(args, 0);
    const arg1 = at(args, 1);
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
