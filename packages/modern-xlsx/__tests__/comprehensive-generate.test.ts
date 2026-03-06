import { mkdirSync, writeFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readBuffer } from '../src/index.js';
import { StyleBuilder } from '../src/style-builder.js';
import { initWasm } from '../src/wasm-loader.js';
import { Workbook } from '../src/workbook.js';

describe('Comprehensive XLSX Generation', () => {
  it('generates comprehensive test file with all features', async () => {
    await initWasm();
    const wb = new Workbook();

    // --- Sheet 1: Styles showcase ---
    const ws1 = wb.addSheet('Styles');

    // Header row
    const headerStyle = new StyleBuilder()
      .font({ bold: true, size: 14, color: 'FFFFFF' })
      .fill({ pattern: 'solid', fgColor: '4472C4' })
      .border({
        left: { style: 'thin', color: '000000' },
        right: { style: 'thin', color: '000000' },
        top: { style: 'thin', color: '000000' },
        bottom: { style: 'thin', color: '000000' },
      })
      .build(wb.styles);
    ws1.cell('A1').value = 'Feature';
    ws1.cell('A1').styleIndex = headerStyle;
    ws1.cell('B1').value = 'Demo';
    ws1.cell('B1').styleIndex = headerStyle;
    ws1.cell('C1').value = 'Notes';
    ws1.cell('C1').styleIndex = headerStyle;

    // Bold
    const boldStyle = new StyleBuilder().font({ bold: true }).build(wb.styles);
    ws1.cell('A2').value = 'Bold';
    ws1.cell('B2').value = 'Bold text';
    ws1.cell('B2').styleIndex = boldStyle;

    // Italic
    const italicStyle = new StyleBuilder().font({ italic: true }).build(wb.styles);
    ws1.cell('A3').value = 'Italic';
    ws1.cell('B3').value = 'Italic text';
    ws1.cell('B3').styleIndex = italicStyle;

    // Number format
    const numFmtStyle = new StyleBuilder().numberFormat('#,##0.00').build(wb.styles);
    ws1.cell('A4').value = 'Number format';
    ws1.cell('B4').value = 1234567.89;
    ws1.cell('B4').styleIndex = numFmtStyle;

    // Date format
    const dateFmtStyle = new StyleBuilder().numberFormat('yyyy-mm-dd').build(wb.styles);
    ws1.cell('A5').value = 'Date format';
    ws1.cell('B5').value = 45658; // 2024-12-31
    ws1.cell('B5').styleIndex = dateFmtStyle;

    // Fill colors
    const redFill = new StyleBuilder()
      .fill({ pattern: 'solid', fgColor: 'FF0000' })
      .build(wb.styles);
    ws1.cell('A6').value = 'Red fill';
    ws1.cell('B6').value = 'Red background';
    ws1.cell('B6').styleIndex = redFill;

    const greenFill = new StyleBuilder()
      .fill({ pattern: 'solid', fgColor: '00FF00' })
      .build(wb.styles);
    ws1.cell('A7').value = 'Green fill';
    ws1.cell('B7').value = 'Green background';
    ws1.cell('B7').styleIndex = greenFill;

    // Borders
    const borderStyle = new StyleBuilder()
      .border({
        left: { style: 'medium', color: '000000' },
        right: { style: 'medium', color: '000000' },
        top: { style: 'medium', color: '000000' },
        bottom: { style: 'medium', color: '000000' },
      })
      .build(wb.styles);
    ws1.cell('A8').value = 'Borders';
    ws1.cell('B8').value = 'Medium borders';
    ws1.cell('B8').styleIndex = borderStyle;

    // Column widths
    ws1.setColumnWidth(1, 20);
    ws1.setColumnWidth(2, 25);
    ws1.setColumnWidth(3, 30);

    // --- Sheet 2: Data features ---
    const ws2 = wb.addSheet('Data');

    // Headers
    const labels = ['ID', 'Name', 'Score', 'Grade', 'Email'];
    for (const [col, label] of labels.entries()) {
      ws2.cell(`${String.fromCharCode(65 + col)}1`).value = label;
    }

    // Data rows
    const students = [
      { id: 1, name: 'Alice', score: 95, grade: 'A', email: 'alice@example.com' },
      { id: 2, name: 'Bob', score: 82, grade: 'B', email: 'bob@example.com' },
      { id: 3, name: 'Charlie', score: 67, grade: 'D', email: 'charlie@example.com' },
      { id: 4, name: 'Diana', score: 91, grade: 'A', email: 'diana@example.com' },
      { id: 5, name: 'Eve', score: 78, grade: 'C', email: 'eve@example.com' },
    ];

    for (const [i, s] of students.entries()) {
      const row = i + 2;
      ws2.cell(`A${row}`).value = s.id;
      ws2.cell(`B${row}`).value = s.name;
      ws2.cell(`C${row}`).value = s.score;
      ws2.cell(`D${row}`).value = s.grade;
      ws2.cell(`E${row}`).value = s.email;
    }

    // Merge cells
    ws2.cell('A8').value = 'Merged Header';
    ws2.addMergeCell('A8:E8');

    // Data validation (list)
    ws2.addValidation('D2:D6', {
      validationType: 'list',
      operator: null,
      formula1: '"A,B,C,D,F"',
      formula2: null,
      allowBlank: true,
      showErrorMessage: true,
      errorTitle: 'Invalid Grade',
      errorMessage: 'Please enter A, B, C, D, or F',
    });

    // Hyperlinks
    ws2.addHyperlink('E2', 'mailto:alice@example.com', {
      display: 'alice@example.com',
      tooltip: 'Send email to Alice',
    });

    // Comments
    ws2.addComment('C3', 'Teacher', 'Needs improvement in math');
    ws2.addComment('C2', 'Teacher', 'Excellent performance');

    // Auto filter
    ws2.autoFilter = 'A1:E6';

    // --- Sheet 3: Formulas ---
    const ws3 = wb.addSheet('Formulas');

    ws3.cell('A1').value = 'Value 1';
    ws3.cell('B1').value = 'Value 2';
    ws3.cell('C1').value = 'SUM';
    ws3.cell('D1').value = 'AVERAGE';
    ws3.cell('E1').value = 'IF Result';

    for (let i = 2; i <= 11; i++) {
      ws3.cell(`A${i}`).value = i * 10;
      ws3.cell(`B${i}`).value = i * 5;
      ws3.cell(`C${i}`).formula = `SUM(A${i},B${i})`;
      ws3.cell(`C${i}`).value = i * 15; // cached value
      ws3.cell(`D${i}`).formula = `AVERAGE(A${i},B${i})`;
      ws3.cell(`D${i}`).value = i * 7.5; // cached value
      ws3.cell(`E${i}`).formula = `IF(A${i}>50,"High","Low")`;
      ws3.cell(`E${i}`).value = i * 10 > 50 ? 'High' : 'Low'; // cached value
    }

    // Totals row
    ws3.cell('A12').value = 'Totals:';
    ws3.cell('C12').formula = 'SUM(C2:C11)';
    ws3.cell('C12').value = 825; // cached value
    ws3.cell('D12').formula = 'AVERAGE(D2:D11)';
    ws3.cell('D12').value = 48.75; // cached value

    // --- Sheet 4: Large data ---
    const ws4 = wb.addSheet('LargeData');

    // Headers
    ws4.cell('A1').value = 'Row';
    ws4.cell('B1').value = 'Random';
    ws4.cell('C1').value = 'Category';
    ws4.cell('D1').value = 'Amount';

    const categories = ['Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon'];
    for (let i = 2; i <= 1001; i++) {
      ws4.cell(`A${i}`).value = i - 1;
      ws4.cell(`B${i}`).value = Math.round(Math.random() * 10000) / 100;
      ws4.cell(`C${i}`).value = categories[(i - 2) % categories.length] ?? 'Alpha';
      ws4.cell(`D${i}`).value = Math.round(Math.random() * 100000) / 100;
    }

    // Freeze the header row
    ws4.frozenPane = { rows: 1, cols: 0 };

    // Column widths
    ws4.setColumnWidth(1, 10);
    ws4.setColumnWidth(2, 12);
    ws4.setColumnWidth(3, 15);
    ws4.setColumnWidth(4, 15);

    // --- Write and roundtrip verify ---
    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    // Roundtrip verify
    const wb2 = await readBuffer(buffer);
    expect(wb2.sheetCount).toBe(4);
    expect(wb2.sheetNames).toEqual(['Styles', 'Data', 'Formulas', 'LargeData']);

    // Verify Styles sheet
    const ws1r = wb2.getSheetByIndex(0);
    expect(ws1r?.rowCount).toBeGreaterThanOrEqual(8);

    // Verify Data sheet
    const ws2r = wb2.getSheetByIndex(1);
    expect(ws2r?.mergeCells).toContain('A8:E8');
    expect(ws2r?.comments.length).toBe(2);
    expect(ws2r?.hyperlinks.length).toBe(1);
    expect(ws2r?.validations.length).toBe(1);
    expect(ws2r?.autoFilter).toBeTruthy();

    // Verify Formulas sheet
    const ws3r = wb2.getSheetByIndex(2);
    expect(ws3r?.rowCount).toBeGreaterThanOrEqual(12);

    // Verify LargeData sheet
    const ws4r = wb2.getSheetByIndex(3);
    expect(ws4r?.rowCount).toBe(1001);
    expect(ws4r?.frozenPane).toEqual({ rows: 1, cols: 0 });

    // Save for manual inspection
    mkdirSync('test-output', { recursive: true });
    writeFileSync('test-output/comprehensive-test.xlsx', buffer);
  });
});
