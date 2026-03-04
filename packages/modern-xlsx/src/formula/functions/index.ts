/**
 * Built-in Excel formula function library.
 *
 * Provides a pre-populated function registry containing all supported
 * string, logical, math, statistical, and lookup functions.
 *
 * @module formula/functions
 */

import type { FormulaFunction } from '../resolver.js';
import { registerLookupFunctions } from './lookup.js';
import { registerMathStatsFunctions } from './math-stats.js';
import { registerStringLogicalFunctions } from './string-logical.js';

/**
 * Create a function registry pre-loaded with all built-in Excel functions.
 * Pass this as `ctx.functions` when calling `evaluateFormula`.
 */
export function createDefaultFunctions(): Map<string, FormulaFunction> {
  const registry = new Map<string, FormulaFunction>();
  registerStringLogicalFunctions(registry);
  registerMathStatsFunctions(registry);
  registerLookupFunctions(registry);
  return registry;
}
