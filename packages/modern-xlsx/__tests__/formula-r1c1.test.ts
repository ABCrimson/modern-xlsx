import { describe, expect, it } from 'vitest';
import { tokenize } from '../src/index.js';

/**
 * modern-xlsx has no R1C1 conversion API, so there is no Excel oracle
 * (Application.ConvertFormula) to compare against. These tests pin only what the
 * public tokenizer does today with R1C1-looking text: `R1` is a row-only
 * A1-style cell reference, `R1C1` is a defined-name token.
 */
describe('R1C1-looking text in the tokenizer', () => {
  it('tokenizes R1 as a cell_ref', () => {
    expect(tokenize('R1')).toEqual({
      tokens: [{ type: 'cell_ref', value: 'R1', start: 0, end: 2 }],
      errors: [],
    });
  });

  it('tokenizes R1C1 as a name', () => {
    expect(tokenize('R1C1')).toEqual({
      tokens: [{ type: 'name', value: 'R1C1', start: 0, end: 4 }],
      errors: [],
    });
  });

  it('tokenizes =R1C1 as a name at 1..5', () => {
    const { tokens, errors } = tokenize('=R1C1');
    expect(errors).toEqual([]);
    expect(tokens).toContainEqual({ type: 'name', value: 'R1C1', start: 1, end: 5 });
  });
});
