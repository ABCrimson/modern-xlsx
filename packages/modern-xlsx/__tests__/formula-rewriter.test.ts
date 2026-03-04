import { describe, expect, it } from 'vitest';
import { rewriteFormula } from '../src/formula/rewriter.js';

describe('rewriteFormula', () => {
  // -----------------------------------------------------------------------
  // Insert rows
  // -----------------------------------------------------------------------
  describe('insert_rows', () => {
    it('shifts ref at the insert point', () => {
      expect(rewriteFormula('A1', { type: 'insert_rows', start: 1, count: 2 })).toBe('A3');
    });

    it('shifts ref below the insert point', () => {
      expect(rewriteFormula('A5', { type: 'insert_rows', start: 3, count: 2 })).toBe('A7');
    });

    it('does not shift ref above the insert point', () => {
      expect(rewriteFormula('A1', { type: 'insert_rows', start: 5, count: 2 })).toBe('A1');
    });

    it('shifts absolute row refs', () => {
      expect(rewriteFormula('$A$1', { type: 'insert_rows', start: 1, count: 2 })).toBe('$A$3');
    });
  });

  // -----------------------------------------------------------------------
  // Delete rows
  // -----------------------------------------------------------------------
  describe('delete_rows', () => {
    it('shifts ref below deleted range', () => {
      // A5, delete rows 2-3 (start=2, count=2) → A3
      expect(rewriteFormula('A5', { type: 'delete_rows', start: 2, count: 2 })).toBe('A3');
    });

    it('replaces ref in deleted range with #REF!', () => {
      expect(rewriteFormula('A2', { type: 'delete_rows', start: 2, count: 1 })).toBe('#REF!');
    });

    it('does not shift ref above deleted range', () => {
      expect(rewriteFormula('A1', { type: 'delete_rows', start: 3, count: 2 })).toBe('A1');
    });
  });

  // -----------------------------------------------------------------------
  // Insert columns
  // -----------------------------------------------------------------------
  describe('insert_cols', () => {
    it('shifts ref at the insert column', () => {
      // C1, insert 1 col at C (index 2) → D1
      expect(rewriteFormula('C1', { type: 'insert_cols', start: 2, count: 1 })).toBe('D1');
    });

    it('shifts ref after the insert column', () => {
      // D1, insert 2 cols at B (index 1) → F1
      expect(rewriteFormula('D1', { type: 'insert_cols', start: 1, count: 2 })).toBe('F1');
    });

    it('does not shift ref before the insert column', () => {
      expect(rewriteFormula('A1', { type: 'insert_cols', start: 2, count: 1 })).toBe('A1');
    });
  });

  // -----------------------------------------------------------------------
  // Delete columns
  // -----------------------------------------------------------------------
  describe('delete_cols', () => {
    it('shifts ref after deleted column', () => {
      // D1, delete col B (index 1, count 1) → C1
      expect(rewriteFormula('D1', { type: 'delete_cols', start: 1, count: 1 })).toBe('C1');
    });

    it('replaces ref in deleted column with #REF!', () => {
      // B1, delete col B (index 1, count 1) → #REF!
      expect(rewriteFormula('B1', { type: 'delete_cols', start: 1, count: 1 })).toBe('#REF!');
    });

    it('does not shift ref before deleted column', () => {
      expect(rewriteFormula('A1', { type: 'delete_cols', start: 1, count: 1 })).toBe('A1');
    });
  });

  // -----------------------------------------------------------------------
  // Ranges
  // -----------------------------------------------------------------------
  describe('ranges', () => {
    it('adjusts range end on row insert', () => {
      // A1:A10, insert 2 rows at 5 → A1:A12
      expect(rewriteFormula('A1:A10', { type: 'insert_rows', start: 5, count: 2 })).toBe('A1:A12');
    });

    it('adjusts both start and end when both are affected', () => {
      // B5:B10, insert 3 rows at 1 → B8:B13
      expect(rewriteFormula('B5:B10', { type: 'insert_rows', start: 1, count: 3 })).toBe('B8:B13');
    });

    it('adjusts column range on col insert', () => {
      // A1:C1, insert 1 col at B (index 1) → A1:D1
      expect(rewriteFormula('A1:C1', { type: 'insert_cols', start: 1, count: 1 })).toBe('A1:D1');
    });

    it('replaces range endpoint with #REF! on delete', () => {
      // A1:A5, delete rows 5-5 → A1:#REF!
      expect(rewriteFormula('A1:A5', { type: 'delete_rows', start: 5, count: 1 })).toBe('A1:#REF!');
    });
  });

  // -----------------------------------------------------------------------
  // Sheet-qualified references
  // -----------------------------------------------------------------------
  describe('sheet-qualified refs', () => {
    it('does not modify ref on different sheet', () => {
      // Sheet2!A1 when inserting rows on Sheet1 → unchanged
      expect(
        rewriteFormula('Sheet2!A1', { type: 'insert_rows', sheet: 'Sheet1', start: 1, count: 2 }),
      ).toBe('Sheet2!A1');
    });

    it('modifies ref when sheet matches', () => {
      expect(
        rewriteFormula('Sheet1!A1', { type: 'insert_rows', sheet: 'Sheet1', start: 1, count: 2 }),
      ).toBe('Sheet1!A3');
    });

    it('does not modify unqualified ref when action targets a specific sheet', () => {
      // Unqualified A1 should not change when action specifies Sheet1
      expect(
        rewriteFormula('A1', { type: 'insert_rows', sheet: 'Sheet1', start: 1, count: 2 }),
      ).toBe('A1');
    });
  });

  // -----------------------------------------------------------------------
  // Formulas with functions
  // -----------------------------------------------------------------------
  describe('formulas with functions', () => {
    it('adjusts refs inside SUM', () => {
      // SUM(A1:A10), insert row at 5 → SUM(A1:A12)
      expect(rewriteFormula('SUM(A1:A10)', { type: 'insert_rows', start: 5, count: 2 })).toBe(
        'SUM(A1:A12)',
      );
    });

    it('handles #REF! inside complex formula on delete', () => {
      // IF(A1>0,SUM(B2:B10),0), delete row 1 → IF(#REF!>0,SUM(B1:B9),0)
      expect(
        rewriteFormula('IF(A1>0,SUM(B2:B10),0)', { type: 'delete_rows', start: 1, count: 1 }),
      ).toBe('IF(#REF!>0,SUM(B1:B9),0)');
    });
  });

  // -----------------------------------------------------------------------
  // Edge cases
  // -----------------------------------------------------------------------
  describe('edge cases', () => {
    it('handles formula with no cell refs', () => {
      expect(rewriteFormula('1+2', { type: 'insert_rows', start: 1, count: 1 })).toBe('1+2');
    });

    it('handles named range (no transformation)', () => {
      expect(rewriteFormula('SUM(myRange)', { type: 'insert_rows', start: 1, count: 1 })).toBe(
        'SUM(myRange)',
      );
    });

    it('returns formula unchanged on parse error', () => {
      expect(rewriteFormula('=??invalid', { type: 'insert_rows', start: 1, count: 1 })).toBe(
        '=??invalid',
      );
    });
  });
});
