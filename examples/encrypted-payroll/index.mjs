/**
 * Encrypted Payroll — modern-xlsx example
 *
 * Creates a password-protected payroll workbook demonstrating:
 *   - Employee data (name, department, salary, bonus, net pay)
 *   - Hidden formulas via sheet protection
 *   - Workbook structure protection
 *   - AES-256 file encryption via toBuffer({ password })
 *   - Professional styling with department color coding
 *
 * The output file requires the password "payroll2025" to open.
 */

import { writeFileSync } from 'node:fs';
import { Workbook, StyleBuilder, initWasm } from 'modern-xlsx';

// ---------------------------------------------------------------------------
// 1. Initialize WASM
// ---------------------------------------------------------------------------
await initWasm();

const PASSWORD = 'payroll2025';

// ---------------------------------------------------------------------------
// 2. Employee data
// ---------------------------------------------------------------------------
const employees = [
  { name: 'Alice Johnson',    dept: 'Engineering',  salary: 125000, bonus: 0.15 },
  { name: 'Bob Smith',        dept: 'Marketing',    salary:  85000, bonus: 0.10 },
  { name: 'Carol Williams',   dept: 'Engineering',  salary: 135000, bonus: 0.18 },
  { name: 'David Brown',      dept: 'Sales',        salary:  72000, bonus: 0.12 },
  { name: 'Eva Martinez',     dept: 'Engineering',  salary: 115000, bonus: 0.14 },
  { name: 'Frank Lee',        dept: 'Marketing',    salary:  92000, bonus: 0.11 },
  { name: 'Grace Kim',        dept: 'Sales',        salary:  68000, bonus: 0.09 },
  { name: 'Henry Davis',      dept: 'Engineering',  salary: 145000, bonus: 0.20 },
  { name: 'Iris Wang',        dept: 'Finance',      salary:  98000, bonus: 0.13 },
  { name: 'Jack Wilson',      dept: 'Finance',      salary: 105000, bonus: 0.12 },
  { name: 'Karen Taylor',     dept: 'Sales',        salary:  76000, bonus: 0.10 },
  { name: 'Leo Anderson',     dept: 'Marketing',    salary:  88000, bonus: 0.11 },
];

// Department color palette
const DEPT_COLORS = {
  Engineering: { bg: 'DBEAFE', text: '1E40AF' },
  Marketing:   { bg: 'FEF3C7', text: '92400E' },
  Sales:       { bg: 'D1FAE5', text: '065F46' },
  Finance:     { bg: 'EDE9FE', text: '5B21B6' },
};

// ---------------------------------------------------------------------------
// 3. Create workbook
// ---------------------------------------------------------------------------
const wb = new Workbook();

// ---------------------------------------------------------------------------
// 4. Summary sheet (visible, partially protected)
// ---------------------------------------------------------------------------
const summary = wb.addSheet('Summary');

// Title
const titleStyle = new StyleBuilder()
  .font({ bold: true, size: 16, color: '1F2937' })
  .alignment({ horizontal: 'center' })
  .build(wb.styles);

summary.cell('A1').value = 'Payroll Summary - Q1 2025';
summary.cell('A1').styleIndex = titleStyle;
summary.addMergeCell('A1:D1');

// Confidentiality notice
const noticeStyle = new StyleBuilder()
  .font({ italic: true, size: 9, color: 'DC2626' })
  .alignment({ horizontal: 'center' })
  .build(wb.styles);

summary.cell('A2').value = 'CONFIDENTIAL - Authorized personnel only';
summary.cell('A2').styleIndex = noticeStyle;
summary.addMergeCell('A2:D2');

// Summary table headers
const sumHdrStyle = new StyleBuilder()
  .font({ bold: true, color: 'FFFFFF', size: 11 })
  .fill({ pattern: 'solid', fgColor: '374151' })
  .alignment({ horizontal: 'center' })
  .border({
    bottom: { style: 'medium', color: '1F2937' },
  })
  .build(wb.styles);

const sumHeaders = ['Department', 'Headcount', 'Total Salary', 'Avg Salary'];
sumHeaders.forEach((h, i) => {
  summary.cell(colLetter(i) + '4').value = h;
  summary.cell(colLetter(i) + '4').styleIndex = sumHdrStyle;
});

// Aggregate by department
const depts = [...new Set(employees.map((e) => e.dept))].sort();
const currStyle = new StyleBuilder()
  .numberFormat('$#,##0')
  .alignment({ horizontal: 'right' })
  .border({ bottom: { style: 'thin', color: 'D1D5DB' } })
  .build(wb.styles);

const centerStyle = new StyleBuilder()
  .alignment({ horizontal: 'center' })
  .border({ bottom: { style: 'thin', color: 'D1D5DB' } })
  .build(wb.styles);

const nameStyle = new StyleBuilder()
  .font({ bold: true })
  .border({ bottom: { style: 'thin', color: 'D1D5DB' } })
  .build(wb.styles);

depts.forEach((dept, i) => {
  const rowNum = 5 + i;
  const deptEmps = employees.filter((e) => e.dept === dept);
  const totalSalary = deptEmps.reduce((sum, e) => sum + e.salary, 0);
  const avgSalary = totalSalary / deptEmps.length;

  summary.cell('A' + rowNum).value = dept;
  summary.cell('A' + rowNum).styleIndex = nameStyle;
  summary.cell('B' + rowNum).value = deptEmps.length;
  summary.cell('B' + rowNum).styleIndex = centerStyle;
  summary.cell('C' + rowNum).value = totalSalary;
  summary.cell('C' + rowNum).styleIndex = currStyle;
  summary.cell('D' + rowNum).value = Math.round(avgSalary);
  summary.cell('D' + rowNum).styleIndex = currStyle;
});

// Totals
const totalRowNum = 5 + depts.length;
const totalLabelStyle = new StyleBuilder()
  .font({ bold: true, size: 11 })
  .fill({ pattern: 'solid', fgColor: 'F3F4F6' })
  .border({ top: { style: 'double', color: '374151' } })
  .build(wb.styles);

const totalNumStyle = new StyleBuilder()
  .font({ bold: true })
  .fill({ pattern: 'solid', fgColor: 'F3F4F6' })
  .numberFormat('$#,##0')
  .alignment({ horizontal: 'right' })
  .border({ top: { style: 'double', color: '374151' } })
  .build(wb.styles);

const totalCenterStyle = new StyleBuilder()
  .font({ bold: true })
  .fill({ pattern: 'solid', fgColor: 'F3F4F6' })
  .alignment({ horizontal: 'center' })
  .border({ top: { style: 'double', color: '374151' } })
  .build(wb.styles);

summary.cell('A' + totalRowNum).value = 'TOTAL';
summary.cell('A' + totalRowNum).styleIndex = totalLabelStyle;
summary.cell('B' + totalRowNum).value = employees.length;
summary.cell('B' + totalRowNum).styleIndex = totalCenterStyle;
summary.cell('C' + totalRowNum).formula = `SUM(C5:C${totalRowNum - 1})`;
summary.cell('C' + totalRowNum).styleIndex = totalNumStyle;
summary.cell('D' + totalRowNum).formula = `AVERAGE(D5:D${totalRowNum - 1})`;
summary.cell('D' + totalRowNum).styleIndex = totalNumStyle;

// Column widths
summary.setColumnWidth(1, 18);
summary.setColumnWidth(2, 14);
summary.setColumnWidth(3, 16);
summary.setColumnWidth(4, 16);

// Freeze header
summary.frozenPane = { rows: 3, cols: 0 };

// ---------------------------------------------------------------------------
// 5. Detail sheet (hidden formulas)
// ---------------------------------------------------------------------------
const detail = wb.addSheet('Payroll Detail');

// Headers
const detailHdrStyle = new StyleBuilder()
  .font({ bold: true, color: 'FFFFFF', size: 11 })
  .fill({ pattern: 'solid', fgColor: '1F2937' })
  .alignment({ horizontal: 'center' })
  .border({ bottom: { style: 'medium', color: '111827' } })
  .build(wb.styles);

const detHeaders = ['Employee', 'Department', 'Base Salary', 'Bonus %', 'Bonus Amount', 'Net Pay'];
detHeaders.forEach((h, i) => {
  detail.cell(colLetter(i) + '1').value = h;
  detail.cell(colLetter(i) + '1').styleIndex = detailHdrStyle;
});

// Build per-department styles
const deptStyles = {};
for (const [dept, colors] of Object.entries(DEPT_COLORS)) {
  deptStyles[dept] = new StyleBuilder()
    .font({ color: colors.text, bold: true, size: 10 })
    .fill({ pattern: 'solid', fgColor: colors.bg })
    .alignment({ horizontal: 'center' })
    .border({
      left:   { style: 'thin', color: 'E5E7EB' },
      right:  { style: 'thin', color: 'E5E7EB' },
      top:    { style: 'thin', color: 'E5E7EB' },
      bottom: { style: 'thin', color: 'E5E7EB' },
    })
    .build(wb.styles);
}

// Body styles
const salaryStyle = new StyleBuilder()
  .numberFormat('$#,##0')
  .alignment({ horizontal: 'right' })
  .border({
    left:   { style: 'thin', color: 'E5E7EB' },
    right:  { style: 'thin', color: 'E5E7EB' },
    top:    { style: 'thin', color: 'E5E7EB' },
    bottom: { style: 'thin', color: 'E5E7EB' },
  })
  .build(wb.styles);

const pctStyle = new StyleBuilder()
  .numberFormat('0%')
  .alignment({ horizontal: 'center' })
  .border({
    left:   { style: 'thin', color: 'E5E7EB' },
    right:  { style: 'thin', color: 'E5E7EB' },
    top:    { style: 'thin', color: 'E5E7EB' },
    bottom: { style: 'thin', color: 'E5E7EB' },
  })
  .build(wb.styles);

const bodyBorderStyle = new StyleBuilder()
  .border({
    left:   { style: 'thin', color: 'E5E7EB' },
    right:  { style: 'thin', color: 'E5E7EB' },
    top:    { style: 'thin', color: 'E5E7EB' },
    bottom: { style: 'thin', color: 'E5E7EB' },
  })
  .build(wb.styles);

// Write employee rows
employees.forEach((emp, i) => {
  const row = i + 2;

  // Name
  detail.cell('A' + row).value = emp.name;
  detail.cell('A' + row).styleIndex = bodyBorderStyle;

  // Department (color-coded)
  detail.cell('B' + row).value = emp.dept;
  detail.cell('B' + row).styleIndex = deptStyles[emp.dept] || bodyBorderStyle;

  // Base Salary
  detail.cell('C' + row).value = emp.salary;
  detail.cell('C' + row).styleIndex = salaryStyle;

  // Bonus %
  detail.cell('D' + row).value = emp.bonus;
  detail.cell('D' + row).styleIndex = pctStyle;

  // Bonus Amount (formula: salary * bonus %)
  const bonusCell = detail.cell('E' + row);
  bonusCell.formula = `C${row}*D${row}`;
  bonusCell.value = emp.salary * emp.bonus; // cached result
  bonusCell.styleIndex = salaryStyle;

  // Net Pay (formula: salary + bonus)
  const netCell = detail.cell('F' + row);
  netCell.formula = `C${row}+E${row}`;
  netCell.value = emp.salary + emp.salary * emp.bonus; // cached result
  netCell.styleIndex = salaryStyle;
});

// Column widths
detail.setColumnWidth(1, 20);
detail.setColumnWidth(2, 16);
detail.setColumnWidth(3, 16);
detail.setColumnWidth(4, 12);
detail.setColumnWidth(5, 16);
detail.setColumnWidth(6, 16);

// Freeze header
detail.frozenPane = { rows: 1, cols: 0 };
detail.autoFilter = `A1:F${employees.length + 1}`;

// ---------------------------------------------------------------------------
// 6. Sheet protection — hide formulas, prevent edits
// ---------------------------------------------------------------------------
detail.sheetProtection = {
  sheet: true,           // Enable sheet protection
  objects: true,         // Protect objects
  scenarios: true,       // Protect scenarios
  password: null,        // No separate sheet password (file-level encryption handles it)
  formatCells: false,    // Do not allow formatting
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  deleteColumns: false,
  deleteRows: false,
  sort: true,            // Allow sorting
  autoFilter: true,      // Allow filtering
};

// Also protect the summary sheet
summary.sheetProtection = {
  sheet: true,
  objects: true,
  scenarios: true,
  password: null,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  deleteColumns: false,
  deleteRows: false,
  sort: true,
  autoFilter: true,
};

// ---------------------------------------------------------------------------
// 7. Workbook protection — prevent structural changes
// ---------------------------------------------------------------------------
wb.protection = {
  lockStructure: true,   // Prevent adding/removing/renaming sheets
  lockWindows: true,     // Prevent window resizing
};

// ---------------------------------------------------------------------------
// 8. Document properties
// ---------------------------------------------------------------------------
wb.docProperties = {
  title: 'Payroll Q1 2025',
  creator: 'HR Department',
  description: 'Confidential payroll data. Password-protected.',
};

// ---------------------------------------------------------------------------
// 9. Write encrypted file
// ---------------------------------------------------------------------------
console.log(`Encrypting with password "${PASSWORD}"...`);
const buffer = await wb.toBuffer({ password: PASSWORD });
writeFileSync('payroll-encrypted.xlsx', buffer);

console.log('\nCreated payroll-encrypted.xlsx');
console.log(`  Password: ${PASSWORD}`);
console.log(`  Encryption: AES-256-CBC (Agile Encryption, SHA-512)`);
console.log(`  Sheets: ${wb.sheetCount} (Summary + Payroll Detail)`);
console.log(`  Employees: ${employees.length}`);
console.log(`  Sheet protection: formulas hidden, sort/filter allowed`);
console.log(`  Workbook protection: structure locked`);
console.log('\nTo open: use the password "payroll2025" in Excel, LibreOffice, or Google Sheets.');

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function colLetter(idx) {
  let result = '';
  let n = idx;
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}
