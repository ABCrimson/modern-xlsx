import { describe, expect, it } from 'vitest';
import { parseFormula } from '../src/formula/parser.js';
import { rewriteFormula } from '../src/formula/rewriter.js';
import { quoteSheetName, serializeFormula } from '../src/formula/serializer.js';
import { tokenize } from '../src/formula/tokenizer.js';

describe('quoteSheetName', () => {
  it.each([
    ['1Q', "'1Q'"],
    ['A1', "'A1'"],
    ['xfd1048576', "'xfd1048576'"],
    ['R1C1', "'R1C1'"],
    ['r', "'r'"],
    ['C', "'C'"],
    ['R2C', "'R2C'"],
    ['TRUE', "'TRUE'"],
    ['false', "'false'"],
    ['a b', "'a b'"],
    ["O'Brien", "'O''Brien'"],
    ['Umsätze', "'Umsätze'"],
    ['Data-2', "'Data-2'"],
    ['.x', "'.x'"],
    ['RC1x', "'RC1x'"],
  ])('quotes %j as %s', (name, quoted) => {
    expect(quoteSheetName(name)).toBe(quoted);
  });

  it.each(['Sheet1', 'Q_1', 'Tax.Rate', 'ABCD1', 'TRUEish', '_x'])('leaves %j unquoted', (name) => {
    expect(quoteSheetName(name)).toBe(name);
  });

  it('round-trips an escaped quoted sheet through the parser and serializer', () => {
    const { ast } = parseFormula("='It''s'!A1");
    expect(ast).not.toBeNull();
    if (ast) expect(serializeFormula(ast)).toBe("'It''s'!A1");
  });

  it('quotes a cell-reference-like sheet name when serializing', () => {
    const { ast } = parseFormula("='Q1'!A1+'2024'!B2");
    expect(ast).not.toBeNull();
    if (ast) expect(serializeFormula(ast)).toBe("'Q1'!A1+'2024'!B2");
  });
});

describe('tokenizer', () => {
  it("never emits a 'range' token", () => {
    for (const formula of ['SUM(A1:B2)', 'Sheet1!A1:B2', "'My Sheet'!$A$1:$C$3", 'A:A', '1:3']) {
      expect(tokenize(formula).tokens.some((t) => t.type === 'range')).toBe(false);
    }
  });
});

describe('RewriteAction bases', () => {
  it('insert_cols start is a 0-based column index', () => {
    expect(rewriteFormula('C1', { type: 'insert_cols', start: 2, count: 1 })).toBe('D1');
  });

  it('insert_rows start is a 1-based row number', () => {
    expect(rewriteFormula('A3', { type: 'insert_rows', start: 3, count: 1 })).toBe('A4');
  });
});
