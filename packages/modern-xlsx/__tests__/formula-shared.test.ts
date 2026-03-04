import { describe, expect, it } from 'vitest';
import { expandSharedFormula } from '../src/formula/shared.js';

describe('expandSharedFormula', () => {
  // -----------------------------------------------------------------------
  // No offset
  // -----------------------------------------------------------------------
  it('returns same formula when master and child are the same cell', () => {
    expect(expandSharedFormula('A1*2', 'B1', 'B1')).toBe('A1*2');
  });

  // -----------------------------------------------------------------------
  // Row offset only
  // -----------------------------------------------------------------------
  it('shifts relative row refs by row delta', () => {
    // master B1, child B3 → rowDelta=2
    expect(expandSharedFormula('A1*2', 'B1', 'B3')).toBe('A3*2');
  });

  // -----------------------------------------------------------------------
  // Column offset only
  // -----------------------------------------------------------------------
  it('shifts relative col refs by col delta', () => {
    // master B1, child D1 → colDelta=2
    expect(expandSharedFormula('A1*2', 'B1', 'D1')).toBe('C1*2');
  });

  // -----------------------------------------------------------------------
  // Both offsets
  // -----------------------------------------------------------------------
  it('shifts both row and col for multiple refs', () => {
    // master B1, child D3 → rowDelta=2, colDelta=2
    expect(expandSharedFormula('A1+B1', 'B1', 'D3')).toBe('C3+D3');
  });

  // -----------------------------------------------------------------------
  // Absolute references preserved
  // -----------------------------------------------------------------------
  it('preserves fully absolute refs', () => {
    // $A$1 should not shift, A1 shifts by rowDelta=2
    expect(expandSharedFormula('$A$1*A1', 'B1', 'B3')).toBe('$A$1*A3');
  });

  // -----------------------------------------------------------------------
  // Mixed absolute references
  // -----------------------------------------------------------------------
  it('handles mixed absolute/relative correctly', () => {
    // $A1: absCol=true (no col shift), absRow=false (row 1→3)
    // A$1: absCol=false (col A stays A, delta=0), absRow=true (no row shift)
    expect(expandSharedFormula('$A1+A$1', 'B1', 'B3')).toBe('$A3+A$1');
  });

  // -----------------------------------------------------------------------
  // Range adjustment
  // -----------------------------------------------------------------------
  it('shifts range references', () => {
    // master B1, child B3 → rowDelta=2
    expect(expandSharedFormula('SUM(A1:A10)', 'B1', 'B3')).toBe('SUM(A3:A12)');
  });

  // -----------------------------------------------------------------------
  // Function with mixed refs
  // -----------------------------------------------------------------------
  it('preserves absolute refs in functions', () => {
    // $A$1 stays, B1 shifts to B3
    expect(expandSharedFormula('IF($A$1>0,B1,0)', 'B1', 'B3')).toBe('IF($A$1>0,B3,0)');
  });

  // -----------------------------------------------------------------------
  // Edge case: #REF! for negative row
  // -----------------------------------------------------------------------
  it('produces #REF! when adjusted row goes below 1', () => {
    // master B2, child B1 → rowDelta=-1
    // A1: row 1 + (-1) = 0 < 1 → #REF!
    expect(expandSharedFormula('A1*2', 'B2', 'B1')).toBe('#REF!*2');
  });

  // -----------------------------------------------------------------------
  // Edge case: #REF! for negative column
  // -----------------------------------------------------------------------
  it('produces #REF! when adjusted col goes below A', () => {
    // master B1, child A1 → colDelta=-1
    // A1: col A (0) + (-1) = -1 < 0 → #REF!
    expect(expandSharedFormula('A1*2', 'B1', 'A1')).toBe('#REF!*2');
  });

  // -----------------------------------------------------------------------
  // Complex formula
  // -----------------------------------------------------------------------
  it('handles complex formula with nested functions', () => {
    // master B1, child C2 → rowDelta=1, colDelta=1
    expect(expandSharedFormula('IF(A1>0,SUM(A1:A10),0)', 'B1', 'C2')).toBe(
      'IF(B2>0,SUM(B2:B11),0)',
    );
  });

  // -----------------------------------------------------------------------
  // Sheet-qualified references
  // -----------------------------------------------------------------------
  it('shifts sheet-qualified relative refs', () => {
    // master B1, child B3 → rowDelta=2
    expect(expandSharedFormula('Sheet1!A1+A1', 'B1', 'B3')).toBe('Sheet1!A3+A3');
  });

  // -----------------------------------------------------------------------
  // Parse error returns formula unchanged
  // -----------------------------------------------------------------------
  it('returns original formula on parse error', () => {
    expect(expandSharedFormula('=??invalid', 'B1', 'B3')).toBe('=??invalid');
  });
});
