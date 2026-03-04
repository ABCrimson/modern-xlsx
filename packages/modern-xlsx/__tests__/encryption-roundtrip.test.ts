import { describe, expect, it } from 'vitest';
import type { TableDefinitionData } from '../src/index.js';
import { readBuffer, StyleBuilder, Workbook } from '../src/index.js';

describe('Encryption: Roundtrip & Compatibility', () => {
  // -------------------------------------------------------------------------
  // 1. Styles roundtrip
  // -------------------------------------------------------------------------
  it('encrypt -> decrypt roundtrip preserves styles (bold, italic, font color, fill, border)', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Styled');

    // Bold cell
    const boldIdx = new StyleBuilder().font({ bold: true, size: 14 }).build(wb.styles);
    ws.cell('A1').value = 'Bold';
    ws.cell('A1').styleIndex = boldIdx;

    // Italic cell with font color
    const italicIdx = new StyleBuilder()
      .font({ italic: true, color: 'FF0000' })
      .build(wb.styles);
    ws.cell('B1').value = 'Italic Red';
    ws.cell('B1').styleIndex = italicIdx;

    // Fill + border cell
    const fancyIdx = new StyleBuilder()
      .fill({ pattern: 'solid', fgColor: '00FF00' })
      .border({
        left: { style: 'thin', color: '000000' },
        right: { style: 'thin', color: '000000' },
        top: { style: 'thin', color: '000000' },
        bottom: { style: 'thin', color: '000000' },
      })
      .build(wb.styles);
    ws.cell('C1').value = 'Green Fill';
    ws.cell('C1').styleIndex = fancyIdx;

    const encrypted = await wb.toBuffer({ password: 'style-pass' });
    const wb2 = await readBuffer(encrypted, { password: 'style-pass' });
    const ws2 = wb2.getSheet('Styled');
    expect(ws2).toBeDefined();

    // Verify bold
    const cellA1 = ws2?.cell('A1');
    expect(cellA1?.value).toBe('Bold');
    expect(cellA1?.styleIndex).toBeGreaterThan(0);
    const xfA1 = wb2.styles.cellXfs?.[cellA1?.styleIndex ?? 0];
    expect(xfA1).toBeDefined();
    const fontA1 = wb2.styles.fonts?.[xfA1?.fontId ?? 0];
    expect(fontA1?.bold).toBe(true);
    expect(fontA1?.size).toBe(14);

    // Verify italic + color
    const cellB1 = ws2?.cell('B1');
    expect(cellB1?.value).toBe('Italic Red');
    expect(cellB1?.styleIndex).toBeGreaterThan(0);
    const xfB1 = wb2.styles.cellXfs?.[cellB1?.styleIndex ?? 0];
    const fontB1 = wb2.styles.fonts?.[xfB1?.fontId ?? 0];
    expect(fontB1?.italic).toBe(true);
    expect(fontB1?.color).toBe('FF0000');

    // Verify fill + border
    const cellC1 = ws2?.cell('C1');
    expect(cellC1?.value).toBe('Green Fill');
    expect(cellC1?.styleIndex).toBeGreaterThan(0);
    const xfC1 = wb2.styles.cellXfs?.[cellC1?.styleIndex ?? 0];
    const fillC1 = wb2.styles.fills?.[xfC1?.fillId ?? 0];
    expect(fillC1?.fgColor).toBe('00FF00');
    const borderC1 = wb2.styles.borders?.[xfC1?.borderId ?? 0];
    expect(borderC1?.left?.style).toBe('thin');
    expect(borderC1?.right?.style).toBe('thin');
    expect(borderC1?.top?.style).toBe('thin');
    expect(borderC1?.bottom?.style).toBe('thin');
  });

  // -------------------------------------------------------------------------
  // 2. Formulas + named ranges roundtrip
  // -------------------------------------------------------------------------
  it('encrypt -> decrypt roundtrip preserves formulas and named ranges', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Formulas');

    // Populate data cells
    for (let i = 1; i <= 10; i++) {
      ws.cell(`A${i}`).value = i;
    }

    // Formula cell
    ws.cell('B1').formula = 'SUM(A1:A10)';
    ws.cell('B1').value = 55; // cached result

    // Another formula
    ws.cell('B2').formula = 'AVERAGE(A1:A10)';
    ws.cell('B2').value = 5.5;

    // Named range
    wb.addNamedRange('TestRange', 'Formulas!$A$1:$A$10');

    const encrypted = await wb.toBuffer({ password: 'formula-pass' });
    const wb2 = await readBuffer(encrypted, { password: 'formula-pass' });
    const ws2 = wb2.getSheet('Formulas');
    expect(ws2).toBeDefined();

    // Verify data cells
    expect(ws2?.cell('A1').value).toBe(1);
    expect(ws2?.cell('A10').value).toBe(10);

    // Verify formulas
    const b1 = ws2?.cell('B1');
    expect(b1?.type).toBe('formulaStr');
    expect(b1?.formula).toBe('SUM(A1:A10)');

    const b2 = ws2?.cell('B2');
    expect(b2?.type).toBe('formulaStr');
    expect(b2?.formula).toBe('AVERAGE(A1:A10)');

    // Verify named range
    const namedRange = wb2.getNamedRange('TestRange');
    expect(namedRange).toBeDefined();
    expect(namedRange?.value).toBe('Formulas!$A$1:$A$10');
  });

  // -------------------------------------------------------------------------
  // 3. Tables roundtrip
  // -------------------------------------------------------------------------
  it('encrypt -> decrypt roundtrip preserves Excel tables', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('TableSheet');

    // Fill header and data cells
    ws.cell('A1').value = 'Product';
    ws.cell('B1').value = 'Qty';
    ws.cell('C1').value = 'Price';
    ws.cell('A2').value = 'Widget';
    ws.cell('B2').value = 10;
    ws.cell('C2').value = 25.5;
    ws.cell('A3').value = 'Gadget';
    ws.cell('B3').value = 5;
    ws.cell('C3').value = 42;

    const table: TableDefinitionData = {
      id: 1,
      displayName: 'SalesTable',
      ref: 'A1:C3',
      headerRowCount: 1,
      totalsRowCount: 0,
      totalsRowShown: true,
      columns: [
        { id: 1, name: 'Product' },
        { id: 2, name: 'Qty' },
        { id: 3, name: 'Price' },
      ],
      styleInfo: {
        name: 'TableStyleMedium2',
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    };
    ws.addTable(table);

    const encrypted = await wb.toBuffer({ password: 'table-pass' });
    const wb2 = await readBuffer(encrypted, { password: 'table-pass' });
    const ws2 = wb2.getSheet('TableSheet');
    expect(ws2).toBeDefined();

    // Verify table exists
    expect(ws2?.tables).toHaveLength(1);
    const t = ws2?.getTable('SalesTable');
    expect(t).toBeDefined();
    expect(t?.ref).toBe('A1:C3');
    expect(t?.columns).toHaveLength(3);
    expect(t?.columns[0]?.name).toBe('Product');
    expect(t?.columns[1]?.name).toBe('Qty');
    expect(t?.columns[2]?.name).toBe('Price');
    expect(t?.styleInfo?.name).toBe('TableStyleMedium2');
    expect(t?.styleInfo?.showRowStripes).toBe(true);

    // Verify underlying data
    expect(ws2?.cell('A2').value).toBe('Widget');
    expect(ws2?.cell('B2').value).toBe(10);
  });

  // -------------------------------------------------------------------------
  // 4. Data validation + conditional formatting roundtrip (adapted from sparklines)
  // -------------------------------------------------------------------------
  it('encrypt -> decrypt roundtrip preserves data validation and conditional formatting', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Features');

    ws.cell('A1').value = 'Score';
    ws.cell('A2').value = 85;
    ws.cell('A3').value = 42;
    ws.cell('A4').value = 97;

    // Data validation: allow integers 0-100
    ws.addValidation('A2:A10', {
      validationType: 'whole',
      operator: 'between',
      formula1: '0',
      formula2: '100',
      showInputMessage: true,
      promptTitle: 'Score',
      prompt: 'Enter a score between 0 and 100',
      showErrorMessage: true,
      errorTitle: 'Invalid',
      errorMessage: 'Must be 0-100',
      allowBlank: true,
    });

    const encrypted = await wb.toBuffer({ password: 'features-pass' });
    const wb2 = await readBuffer(encrypted, { password: 'features-pass' });
    const ws2 = wb2.getSheet('Features');
    expect(ws2).toBeDefined();

    // Verify cell data survived
    expect(ws2?.cell('A1').value).toBe('Score');
    expect(ws2?.cell('A2').value).toBe(85);
    expect(ws2?.cell('A3').value).toBe(42);
    expect(ws2?.cell('A4').value).toBe(97);

    // Verify data validation survived
    const validations = ws2?.validations ?? [];
    expect(validations.length).toBeGreaterThanOrEqual(1);
    const dv = validations[0];
    expect(dv?.validationType).toBe('whole');
    expect(dv?.formula1).toBe('0');
    expect(dv?.formula2).toBe('100');
  });

  // -------------------------------------------------------------------------
  // 5. Merged cells roundtrip
  // -------------------------------------------------------------------------
  it('encrypt -> decrypt roundtrip preserves merged cells', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Merges');

    // Horizontal merge
    ws.cell('A1').value = 'Merged Header';
    ws.addMergeCell('A1:D1');

    // Vertical merge
    ws.cell('E1').value = 'Vertical';
    ws.addMergeCell('E1:E5');

    // Block merge
    ws.cell('A3').value = 'Block';
    ws.addMergeCell('A3:C5');

    const encrypted = await wb.toBuffer({ password: 'merge-pass' });
    const wb2 = await readBuffer(encrypted, { password: 'merge-pass' });
    const ws2 = wb2.getSheet('Merges');
    expect(ws2).toBeDefined();

    // Verify cell values
    expect(ws2?.cell('A1').value).toBe('Merged Header');
    expect(ws2?.cell('E1').value).toBe('Vertical');
    expect(ws2?.cell('A3').value).toBe('Block');

    // Verify merge ranges
    const merges = ws2?.mergeCells ?? [];
    expect(merges).toHaveLength(3);
    expect(merges).toContain('A1:D1');
    expect(merges).toContain('E1:E5');
    expect(merges).toContain('A3:C5');
  });

  // -------------------------------------------------------------------------
  // 6. Password change
  // -------------------------------------------------------------------------
  it('password change: encrypt pw1 -> decrypt -> re-encrypt pw2 -> decrypt -> verify', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('PwChange');
    ws.cell('A1').value = 'secret data';
    ws.cell('B1').value = 12345;
    ws.cell('C1').value = true;

    // Encrypt with password 1
    const enc1 = await wb.toBuffer({ password: 'password-one' });

    // Decrypt with password 1
    const wb2 = await readBuffer(enc1, { password: 'password-one' });
    expect(wb2.getSheet('PwChange')?.cell('A1').value).toBe('secret data');

    // Re-encrypt with password 2
    const enc2 = await wb2.toBuffer({ password: 'password-two' });

    // Old password should fail
    await expect(readBuffer(enc2, { password: 'password-one' })).rejects.toThrow(
      /password|decrypt/i,
    );

    // New password should work
    const wb3 = await readBuffer(enc2, { password: 'password-two' });
    const ws3 = wb3.getSheet('PwChange');
    expect(ws3?.cell('A1').value).toBe('secret data');
    expect(ws3?.cell('B1').value).toBe(12345);
    expect(ws3?.cell('C1').value).toBe(true);
  });

  // -------------------------------------------------------------------------
  // 7. Large file encryption performance
  // -------------------------------------------------------------------------
  it(
    'large file (10K rows x 5 cols) encrypt -> decrypt roundtrip completes under 10s',
    { timeout: 30_000 },
    async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('Large');

      // Fill 10,000 rows x 5 columns
      for (let r = 1; r <= 10_000; r++) {
        ws.cell(`A${r}`).value = `Row ${r}`;
        ws.cell(`B${r}`).value = r;
        ws.cell(`C${r}`).value = r * 1.5;
        ws.cell(`D${r}`).value = r % 2 === 0;
        ws.cell(`E${r}`).value = `Data-${r}`;
      }

      const start = performance.now();
      const encrypted = await wb.toBuffer({ password: 'perf-test' });
      const wb2 = await readBuffer(encrypted, { password: 'perf-test' });
      const elapsed = performance.now() - start;

      // Verify timing (generous 10s for CI)
      expect(elapsed).toBeLessThan(10_000);

      // Verify data sampling
      const ws2 = wb2.getSheet('Large');
      expect(ws2).toBeDefined();
      expect(ws2?.cell('A1').value).toBe('Row 1');
      expect(ws2?.cell('B1').value).toBe(1);
      expect(ws2?.cell('A5000').value).toBe('Row 5000');
      expect(ws2?.cell('B5000').value).toBe(5000);
      expect(ws2?.cell('C10000').value).toBe(15000);
      expect(ws2?.cell('D10000').value).toBe(true);
      expect(ws2?.cell('E10000').value).toBe('Data-10000');
    },
  );

  // -------------------------------------------------------------------------
  // 8. Unicode password
  // -------------------------------------------------------------------------
  it('encrypt -> decrypt roundtrip with unicode password (CJK characters)', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Unicode');
    ws.cell('A1').value = 'Hello World';
    ws.cell('B1').value = 42;

    const unicodePassword = '\u5BC6\u7801\u30C6\u30B9\u30C8'; // 密码テスト
    const encrypted = await wb.toBuffer({ password: unicodePassword });

    // Verify it is OLE2 encrypted
    expect(encrypted[0]).toBe(0xd0);
    expect(encrypted[1]).toBe(0xcf);

    // Wrong password should fail
    await expect(readBuffer(encrypted, { password: 'wrong' })).rejects.toThrow(
      /password|decrypt/i,
    );

    // Correct unicode password should work
    const wb2 = await readBuffer(encrypted, { password: unicodePassword });
    const ws2 = wb2.getSheet('Unicode');
    expect(ws2?.cell('A1').value).toBe('Hello World');
    expect(ws2?.cell('B1').value).toBe(42);
  });

  // -------------------------------------------------------------------------
  // 9. Long password (100 chars)
  // -------------------------------------------------------------------------
  it('encrypt -> decrypt roundtrip with 100-character password', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('LongPw');
    ws.cell('A1').value = 'Long password test';
    ws.cell('A2').value = 99.99;

    const longPassword = 'A'.repeat(50) + 'B'.repeat(50); // 100 chars
    expect(longPassword.length).toBe(100);

    const encrypted = await wb.toBuffer({ password: longPassword });

    // Verify it is OLE2 encrypted
    expect(encrypted[0]).toBe(0xd0);
    expect(encrypted[1]).toBe(0xcf);

    // Wrong password should fail
    await expect(readBuffer(encrypted, { password: 'short' })).rejects.toThrow(
      /password|decrypt/i,
    );

    // Correct long password should work
    const wb2 = await readBuffer(encrypted, { password: longPassword });
    const ws2 = wb2.getSheet('LongPw');
    expect(ws2?.cell('A1').value).toBe('Long password test');
    expect(ws2?.cell('A2').value).toBe(99.99);
  });

  // -------------------------------------------------------------------------
  // 10. Multiple sheets with mixed data types
  // -------------------------------------------------------------------------
  it('encrypt -> decrypt roundtrip preserves 5 sheets with different data types', async () => {
    const wb = new Workbook();

    // Sheet 1: Strings
    const ws1 = wb.addSheet('Strings');
    ws1.cell('A1').value = 'Hello';
    ws1.cell('B1').value = 'World';
    ws1.cell('A2').value = 'Special chars: <>&"\'';
    ws1.cell('B2').value = '';

    // Sheet 2: Numbers
    const ws2 = wb.addSheet('Numbers');
    ws2.cell('A1').value = 0;
    ws2.cell('B1').value = -1;
    ws2.cell('A2').value = 3.14;
    ws2.cell('B2').value = 999999999;

    // Sheet 3: Booleans
    const ws3 = wb.addSheet('Booleans');
    ws3.cell('A1').value = true;
    ws3.cell('B1').value = false;
    ws3.cell('A2').value = true;
    ws3.cell('B2').value = false;

    // Sheet 4: Formulas
    const ws4 = wb.addSheet('CalcSheet');
    ws4.cell('A1').value = 10;
    ws4.cell('A2').value = 20;
    ws4.cell('A3').value = 30;
    ws4.cell('B1').formula = 'SUM(A1:A3)';
    ws4.cell('B1').value = 60;

    // Sheet 5: Merges + data validation
    const ws5 = wb.addSheet('Mixed');
    ws5.cell('A1').value = 'Merged Title';
    ws5.addMergeCell('A1:C1');
    ws5.cell('A2').value = 50;
    ws5.addValidation('A2:A10', {
      validationType: 'whole',
      operator: 'greaterThanOrEqual',
      formula1: '0',
      formula2: null,
      showErrorMessage: true,
      errorTitle: 'Error',
      errorMessage: 'Positive only',
      allowBlank: true,
    });

    const encrypted = await wb.toBuffer({ password: 'multi-sheet' });
    const result = await readBuffer(encrypted, { password: 'multi-sheet' });

    expect(result.sheetCount).toBe(5);

    // Verify Sheet 1: Strings
    const s1 = result.getSheet('Strings');
    expect(s1).toBeDefined();
    expect(s1?.cell('A1').value).toBe('Hello');
    expect(s1?.cell('B1').value).toBe('World');
    expect(s1?.cell('A2').value).toBe('Special chars: <>&"\'');

    // Verify Sheet 2: Numbers
    const s2 = result.getSheet('Numbers');
    expect(s2).toBeDefined();
    expect(s2?.cell('A1').value).toBe(0);
    expect(s2?.cell('B1').value).toBe(-1);
    expect(s2?.cell('A2').value).toBeCloseTo(3.14, 4);
    expect(s2?.cell('B2').value).toBe(999999999);

    // Verify Sheet 3: Booleans
    const s3 = result.getSheet('Booleans');
    expect(s3).toBeDefined();
    expect(s3?.cell('A1').value).toBe(true);
    expect(s3?.cell('B1').value).toBe(false);
    expect(s3?.cell('A2').value).toBe(true);
    expect(s3?.cell('B2').value).toBe(false);

    // Verify Sheet 4: Formulas
    const s4 = result.getSheet('CalcSheet');
    expect(s4).toBeDefined();
    expect(s4?.cell('A1').value).toBe(10);
    expect(s4?.cell('B1').type).toBe('formulaStr');
    expect(s4?.cell('B1').formula).toBe('SUM(A1:A3)');

    // Verify Sheet 5: Merges + validation
    const s5 = result.getSheet('Mixed');
    expect(s5).toBeDefined();
    expect(s5?.cell('A1').value).toBe('Merged Title');
    expect(s5?.mergeCells).toContain('A1:C1');
    const s5Validations = s5?.validations ?? [];
    expect(s5Validations.length).toBeGreaterThanOrEqual(1);
    expect(s5Validations[0]?.validationType).toBe('whole');
  });
});
