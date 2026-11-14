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

    it('shrinks the range when its end row is deleted', () => {
      // A1:A5, delete row 5 → A1:A4 (Excel shrinks; it never emits A1:#REF!)
      expect(rewriteFormula('A1:A5', { type: 'delete_rows', start: 5, count: 1 })).toBe('A1:A4');
    });
  });

  // -----------------------------------------------------------------------
  // Range shrink on delete
  // -----------------------------------------------------------------------
  describe('range shrink on delete', () => {
    it('shrinks when the start row is deleted', () => {
      expect(rewriteFormula('SUM(A1:A5)', { type: 'delete_rows', start: 1, count: 1 })).toBe(
        'SUM(A1:A4)',
      );
    });

    it('shrinks when the deleted block overlaps the start', () => {
      // A3:A10, delete rows 2-4 → rows 3-4 go, 5-10 move up 3 → A2:A7
      expect(rewriteFormula('A3:A10', { type: 'delete_rows', start: 2, count: 3 })).toBe('A2:A7');
    });

    it('a range deleted entirely becomes one #REF!', () => {
      expect(rewriteFormula('A2:A3', { type: 'delete_rows', start: 2, count: 2 })).toBe('#REF!');
      expect(rewriteFormula('SUM(A2:A3)', { type: 'delete_rows', start: 2, count: 2 })).toBe(
        'SUM(#REF!)',
      );
    });

    it('shrinks a column span and keeps $ flags', () => {
      // $B$1:$D$1, delete col B → $B$1:$C$1
      expect(rewriteFormula('$B$1:$D$1', { type: 'delete_cols', start: 1, count: 1 })).toBe(
        '$B$1:$C$1',
      );
    });

    it('a shrunk range inside an operator needs no parentheses', () => {
      expect(rewriteFormula('A1:A5+1', { type: 'delete_rows', start: 5, count: 1 })).toBe(
        'A1:A4+1',
      );
      expect(rewriteFormula('A2:A3+1', { type: 'delete_rows', start: 2, count: 2 })).toBe(
        '#REF!+1',
      );
    });

    it('writes each value back to the endpoint that held it', () => {
      // A5:A1 (end above start), delete row 5 → the start endpoint held the high row
      expect(rewriteFormula('A5:A1', { type: 'delete_rows', start: 5, count: 1 })).toBe('A4:A1');
    });
  });

  // -----------------------------------------------------------------------
  // Range end inherits the start's sheet
  // -----------------------------------------------------------------------
  describe('sheet-qualified ranges', () => {
    it('shifts both endpoints of a sheet-qualified range', () => {
      expect(
        rewriteFormula('Sheet1!A1:B10', {
          type: 'insert_rows',
          sheet: 'Sheet1',
          start: 1,
          count: 2,
        }),
      ).toBe('Sheet1!A3:B12');
    });

    it('shrinks a sheet-qualified range on column delete', () => {
      expect(
        rewriteFormula('SUM(Sheet1!A1:C3)', {
          type: 'delete_cols',
          sheet: 'Sheet1',
          start: 1,
          count: 1,
        }),
      ).toBe('SUM(Sheet1!A1:B3)');
    });

    it('leaves a sheet-qualified range alone for an unqualified action', () => {
      expect(rewriteFormula('Sheet1!A1:B10', { type: 'insert_rows', start: 1, count: 2 })).toBe(
        'Sheet1!A1:B10',
      );
    });

    it('shifts a quoted sheet name', () => {
      expect(
        rewriteFormula("'My Sheet'!A1", {
          type: 'insert_rows',
          sheet: 'My Sheet',
          start: 1,
          count: 1,
        }),
      ).toBe("'My Sheet'!A2");
    });

    it('keeps a digit-leading sheet name quoted', () => {
      expect(
        rewriteFormula("'2024'!A1+1", { type: 'insert_rows', sheet: '2024', start: 1, count: 1 }),
      ).toBe("'2024'!A2+1");
    });

    it('passes a 3-D reference through unchanged', () => {
      expect(
        rewriteFormula('SUM(Sheet1:Sheet3!A1)', {
          type: 'insert_rows',
          sheet: 'Sheet1',
          start: 1,
          count: 1,
        }),
      ).toBe('SUM(Sheet1:Sheet3!A1)');
      expect(
        rewriteFormula('SUM(Sheet1:Sheet3!A1:B2)', {
          type: 'delete_rows',
          sheet: 'Sheet1',
          start: 1,
          count: 1,
        }),
      ).toBe('SUM(Sheet1:Sheet3!A1:B2)');
    });
  });

  // -----------------------------------------------------------------------
  // Whole-column/row ranges and other reference forms
  // -----------------------------------------------------------------------
  describe('reference forms', () => {
    it('shifts through modern error literals', () => {
      expect(
        rewriteFormula('IF(A1=#SPILL!,1,B1)', { type: 'insert_rows', start: 1, count: 1 }),
      ).toBe('IF(A2=#SPILL!,1,B2)');
    });

    it('shifts both areas of a union', () => {
      expect(rewriteFormula('SUM((A1,B2))', { type: 'insert_rows', start: 1, count: 1 })).toBe(
        'SUM((A2,B3))',
      );
    });

    it('shifts both operands of an intersection', () => {
      expect(rewriteFormula('SUM(A1:B5 B2:C3)', { type: 'insert_rows', start: 1, count: 1 })).toBe(
        'SUM(A2:B6 B3:C4)',
      );
    });

    it('keeps an empty argument', () => {
      expect(rewriteFormula('IF(A1,,2)', { type: 'insert_rows', start: 1, count: 1 })).toBe(
        'IF(A2,,2)',
      );
    });

    it('shifts a whole-column range on a column insert', () => {
      expect(rewriteFormula('SUM(A:A)', { type: 'insert_cols', start: 0, count: 1 })).toBe(
        'SUM(B:B)',
      );
    });

    it('shifts a whole-row range on a row insert', () => {
      expect(rewriteFormula('SUM(1:1)', { type: 'insert_rows', start: 1, count: 1 })).toBe(
        'SUM(2:2)',
      );
    });

    it('leaves a whole-column range alone on a row action and vice versa', () => {
      expect(rewriteFormula('SUM(A:A)', { type: 'delete_rows', start: 1, count: 1 })).toBe(
        'SUM(A:A)',
      );
      expect(rewriteFormula('SUM(1:1)', { type: 'delete_cols', start: 0, count: 1 })).toBe(
        'SUM(1:1)',
      );
    });

    it('shrinks a whole-column range and keeps $ flags', () => {
      expect(rewriteFormula('SUM($B:$D)', { type: 'delete_cols', start: 2, count: 1 })).toBe(
        'SUM($B:$C)',
      );
      expect(rewriteFormula('SUM(2:5)', { type: 'delete_rows', start: 1, count: 2 })).toBe(
        'SUM(1:3)',
      );
    });

    it('a whole-column or whole-row range deleted entirely becomes #REF!', () => {
      expect(rewriteFormula('SUM(B:B)', { type: 'delete_cols', start: 1, count: 1 })).toBe(
        'SUM(#REF!)',
      );
      expect(rewriteFormula('SUM(2:3)', { type: 'delete_rows', start: 2, count: 2 })).toBe(
        'SUM(#REF!)',
      );
    });

    it('shifts a sheet-qualified whole-column range only on its sheet', () => {
      expect(
        rewriteFormula('SUM(Data!B:C)', { type: 'insert_cols', sheet: 'Data', start: 0, count: 1 }),
      ).toBe('SUM(Data!C:D)');
      expect(rewriteFormula('SUM(Data!B:C)', { type: 'insert_cols', start: 0, count: 1 })).toBe(
        'SUM(Data!B:C)',
      );
    });

    it('passes structured references through', () => {
      expect(
        rewriteFormula('SUM(Table1[Col])+A1', { type: 'insert_rows', start: 1, count: 1 }),
      ).toBe('SUM(Table1[Col])+A2');
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

    it('matches the sheet name ignoring case, as Excel does', () => {
      expect(
        rewriteFormula('sheet1!A5', { type: 'insert_rows', sheet: 'Sheet1', start: 1, count: 1 }),
      ).toBe('sheet1!A6');
      expect(
        rewriteFormula('SUM(DATA!B:C)', { type: 'insert_cols', sheet: 'data', start: 0, count: 1 }),
      ).toBe('SUM(DATA!C:D)');
    });
  });

  // -----------------------------------------------------------------------
  // Range operator with a function endpoint
  // -----------------------------------------------------------------------
  describe('range operator endpoints', () => {
    it('shifts both sides of A2:INDEX(A:A,B1)', () => {
      expect(
        rewriteFormula('SUM(A2:INDEX(A:A,B1))', { type: 'insert_rows', start: 1, count: 1 }),
      ).toBe('SUM(A3:INDEX(A:A,B2))');
    });

    it('shifts every endpoint of A1:B1:C1', () => {
      expect(rewriteFormula('A1:B1:C1', { type: 'insert_cols', start: 1, count: 1 })).toBe(
        'A1:C1:D1',
      );
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

describe('rewriteFormula sheet edges', () => {
  it('insert_cols past XFD gives #REF!', () => {
    expect(rewriteFormula('XFD1+1', { type: 'insert_cols', start: 0, count: 1 })).toBe('#REF!+1');
  });

  it('insert_rows past row 1048576 gives #REF!', () => {
    expect(rewriteFormula('A1048576', { type: 'insert_rows', start: 0, count: 1 })).toBe('#REF!');
    expect(rewriteFormula('A1048576', { type: 'insert_rows', start: 5, count: 1 })).toBe('#REF!');
  });

  it('insert_cols pushing XFD1 off the sheet gives #REF!', () => {
    expect(rewriteFormula('XFD1', { type: 'insert_cols', start: 0, count: 1 })).toBe('#REF!');
  });

  it('clamps a range end that an insert pushes past the last row', () => {
    expect(rewriteFormula('SUM(A2:A1048576)', { type: 'insert_rows', start: 1, count: 1 })).toBe(
      'SUM(A3:A1048576)',
    );
  });

  it('clamps a range end that an insert pushes past XFD', () => {
    expect(rewriteFormula('SUM(B1:XFD1)', { type: 'insert_cols', start: 0, count: 1 })).toBe(
      'SUM(C1:XFD1)',
    );
  });

  it('a range whose start an insert pushes off the sheet becomes #REF!', () => {
    expect(
      rewriteFormula('SUM(A1048575:A1048576)', { type: 'insert_rows', start: 1, count: 2 }),
    ).toBe('SUM(#REF!)');
  });

  it('clamps or invalidates whole-column and whole-row ranges at the edge', () => {
    expect(rewriteFormula('SUM(B:XFD)', { type: 'insert_cols', start: 0, count: 1 })).toBe(
      'SUM(C:XFD)',
    );
    expect(rewriteFormula('SUM(XFD:XFD)', { type: 'insert_cols', start: 0, count: 1 })).toBe(
      'SUM(#REF!)',
    );
    expect(
      rewriteFormula('SUM(1048576:1048576)', { type: 'insert_rows', start: 1, count: 1 }),
    ).toBe('SUM(#REF!)');
  });
});

describe('#REF! range endpoints', () => {
  it('writes #REF! as the endpoint of a deleted chain and keeps rewriting it', () => {
    expect(rewriteFormula('SUM(A2:A2:C3)', { type: 'delete_rows', start: 2, count: 1 })).toBe(
      'SUM(#REF!:C2)',
    );
    expect(rewriteFormula('SUM(#REF!:C2)', { type: 'insert_rows', start: 1, count: 1 })).toBe(
      'SUM(#REF!:C3)',
    );
  });
});
