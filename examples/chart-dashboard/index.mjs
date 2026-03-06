/**
 * Chart Dashboard — modern-xlsx example
 *
 * Creates a multi-sheet workbook with:
 *   - "Raw Data" sheet — monthly metrics for 12 months
 *   - "Dashboard" sheet — 3 charts (bar, line, pie) + summary statistics
 *
 * Open in Excel or LibreOffice to see the charts rendered.
 */

import { writeFileSync } from 'node:fs';
import { Workbook, StyleBuilder, initWasm } from 'modern-xlsx';

// ---------------------------------------------------------------------------
// 1. Initialize WASM
// ---------------------------------------------------------------------------
await initWasm();

// ---------------------------------------------------------------------------
// 2. Sample data — monthly business metrics
// ---------------------------------------------------------------------------
const months = [
  'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
];

const revenue  = [42000, 48000, 51000, 55000, 62000, 58000, 64000, 71000, 68000, 75000, 82000, 91000];
const costs    = [31000, 33000, 35000, 36000, 38000, 37000, 39000, 42000, 40000, 43000, 45000, 48000];
const profit   = revenue.map((r, i) => r - costs[i]);
const customers = [1200, 1350, 1420, 1580, 1690, 1750, 1820, 1950, 2010, 2150, 2280, 2400];

// Product mix for pie chart
const products = [
  { name: 'SaaS Platform',    revenue: 340000 },
  { name: 'Consulting',       revenue: 180000 },
  { name: 'Support Contracts', revenue: 120000 },
  { name: 'Training',         revenue: 65000 },
  { name: 'Other',            revenue: 22000 },
];

// ---------------------------------------------------------------------------
// 3. Create workbook
// ---------------------------------------------------------------------------
const wb = new Workbook();

// ---------------------------------------------------------------------------
// 4. Raw Data sheet
// ---------------------------------------------------------------------------
const data = wb.addSheet('Raw Data');

// Header style
const hdrStyle = new StyleBuilder()
  .font({ bold: true, color: 'FFFFFF', size: 11 })
  .fill({ pattern: 'solid', fgColor: '2D3748' })
  .alignment({ horizontal: 'center' })
  .border({ bottom: { style: 'medium', color: '1A202C' } })
  .build(wb.styles);

// Currency style
const curStyle = new StyleBuilder()
  .numberFormat('$#,##0')
  .alignment({ horizontal: 'right' })
  .border({
    left:   { style: 'thin', color: 'E2E8F0' },
    right:  { style: 'thin', color: 'E2E8F0' },
    top:    { style: 'thin', color: 'E2E8F0' },
    bottom: { style: 'thin', color: 'E2E8F0' },
  })
  .build(wb.styles);

// Number style (customers)
const numStyle = new StyleBuilder()
  .numberFormat('#,##0')
  .alignment({ horizontal: 'right' })
  .border({
    left:   { style: 'thin', color: 'E2E8F0' },
    right:  { style: 'thin', color: 'E2E8F0' },
    top:    { style: 'thin', color: 'E2E8F0' },
    bottom: { style: 'thin', color: 'E2E8F0' },
  })
  .build(wb.styles);

const bodyBorder = new StyleBuilder()
  .border({
    left:   { style: 'thin', color: 'E2E8F0' },
    right:  { style: 'thin', color: 'E2E8F0' },
    top:    { style: 'thin', color: 'E2E8F0' },
    bottom: { style: 'thin', color: 'E2E8F0' },
  })
  .build(wb.styles);

// -- Monthly data table --
const dataHeaders = ['Month', 'Revenue', 'Costs', 'Profit', 'Customers'];
dataHeaders.forEach((h, i) => {
  data.cell(colLetter(i) + '1').value = h;
  data.cell(colLetter(i) + '1').styleIndex = hdrStyle;
});

months.forEach((m, i) => {
  const row = i + 2;
  data.cell('A' + row).value = m;
  data.cell('A' + row).styleIndex = bodyBorder;
  data.cell('B' + row).value = revenue[i];
  data.cell('B' + row).styleIndex = curStyle;
  data.cell('C' + row).value = costs[i];
  data.cell('C' + row).styleIndex = curStyle;
  data.cell('D' + row).value = profit[i];
  data.cell('D' + row).styleIndex = curStyle;
  data.cell('E' + row).value = customers[i];
  data.cell('E' + row).styleIndex = numStyle;
});

// Column widths
data.setColumnWidth(1, 12);
data.setColumnWidth(2, 14);
data.setColumnWidth(3, 14);
data.setColumnWidth(4, 14);
data.setColumnWidth(5, 14);
data.frozenPane = { rows: 1, cols: 0 };

// -- Product mix table (for pie chart) --
// Place it in columns G-H
data.cell('G1').value = 'Product';
data.cell('G1').styleIndex = hdrStyle;
data.cell('H1').value = 'Revenue';
data.cell('H1').styleIndex = hdrStyle;

products.forEach((p, i) => {
  const row = i + 2;
  data.cell('G' + row).value = p.name;
  data.cell('G' + row).styleIndex = bodyBorder;
  data.cell('H' + row).value = p.revenue;
  data.cell('H' + row).styleIndex = curStyle;
});

data.setColumnWidth(7, 20);
data.setColumnWidth(8, 14);

// ---------------------------------------------------------------------------
// 5. Dashboard sheet — charts + summary statistics
// ---------------------------------------------------------------------------
const dash = wb.addSheet('Dashboard');

// -- Title --
const titleStyle = new StyleBuilder()
  .font({ bold: true, size: 20, color: '1A202C' })
  .alignment({ horizontal: 'center', vertical: 'center' })
  .build(wb.styles);

dash.cell('A1').value = 'Business Dashboard 2025';
dash.cell('A1').styleIndex = titleStyle;
dash.addMergeCell('A1:L1');

// -- Summary statistics in row 3 --
const statLabelStyle = new StyleBuilder()
  .font({ bold: true, size: 9, color: '6B7280' })
  .alignment({ horizontal: 'center' })
  .build(wb.styles);

const statValueStyle = new StyleBuilder()
  .font({ bold: true, size: 16, color: '1F2937' })
  .alignment({ horizontal: 'center' })
  .numberFormat('$#,##0')
  .build(wb.styles);

const statValuePct = new StyleBuilder()
  .font({ bold: true, size: 16, color: '059669' })
  .alignment({ horizontal: 'center' })
  .numberFormat('0.0%')
  .build(wb.styles);

const statValueNum = new StyleBuilder()
  .font({ bold: true, size: 16, color: '1F2937' })
  .alignment({ horizontal: 'center' })
  .numberFormat('#,##0')
  .build(wb.styles);

// KPI cards
const totalRevenue = revenue.reduce((a, b) => a + b, 0);
const totalCosts = costs.reduce((a, b) => a + b, 0);
const totalProfit = totalRevenue - totalCosts;
const marginPct = totalProfit / totalRevenue;
const latestCustomers = customers[customers.length - 1];

const kpis = [
  { label: 'TOTAL REVENUE',  value: totalRevenue,    style: statValueStyle,  col: 0 },
  { label: 'TOTAL COSTS',    value: totalCosts,      style: statValueStyle,  col: 3 },
  { label: 'NET PROFIT',     value: totalProfit,     style: statValueStyle,  col: 6 },
  { label: 'PROFIT MARGIN',  value: marginPct,       style: statValuePct,    col: 9 },
];

kpis.forEach((kpi) => {
  const col = colLetter(kpi.col);
  const colEnd = colLetter(kpi.col + 2);

  // Label
  dash.cell(col + '3').value = kpi.label;
  dash.cell(col + '3').styleIndex = statLabelStyle;
  dash.addMergeCell(col + '3:' + colEnd + '3');

  // Value
  dash.cell(col + '4').value = kpi.value;
  dash.cell(col + '4').styleIndex = kpi.style;
  dash.addMergeCell(col + '4:' + colEnd + '4');
});

// ---------------------------------------------------------------------------
// 6. Bar chart — Revenue vs Costs by month
// ---------------------------------------------------------------------------
dash.addChart('bar', (b) => {
  b.title('Monthly Revenue vs Costs')
    .addSeries({
      name: 'Revenue',
      catRef: "'Raw Data'!$A$2:$A$13",
      valRef: "'Raw Data'!$B$2:$B$13",
      fillColor: '4472C4',
    })
    .addSeries({
      name: 'Costs',
      catRef: "'Raw Data'!$A$2:$A$13",
      valRef: "'Raw Data'!$C$2:$C$13",
      fillColor: 'ED7D31',
    })
    .grouping('clustered')
    .legend('bottom')
    .valAxis({ title: 'Amount ($)', numFmt: '$#,##0', majorGridlines: true })
    .catAxis({ title: 'Month' })
    .anchor({ col: 0, row: 5 }, { col: 6, row: 20 });
});

// ---------------------------------------------------------------------------
// 7. Line chart — Profit trend + customer growth
// ---------------------------------------------------------------------------
dash.addChart('line', (b) => {
  b.title('Profit Trend')
    .addSeries({
      name: 'Profit',
      catRef: "'Raw Data'!$A$2:$A$13",
      valRef: "'Raw Data'!$D$2:$D$13",
      lineColor: '34D399',
      lineWidth: 28575, // in EMU (2.25pt)
      marker: 'circle',
    })
    .legend('bottom')
    .valAxis({ title: 'Profit ($)', numFmt: '$#,##0', majorGridlines: true })
    .catAxis({ title: 'Month' })
    .anchor({ col: 6, row: 5 }, { col: 12, row: 20 });
});

// ---------------------------------------------------------------------------
// 8. Pie chart — Product revenue mix
// ---------------------------------------------------------------------------
dash.addChart('pie', (b) => {
  b.title('Revenue by Product')
    .addSeries({
      name: 'Revenue',
      catRef: "'Raw Data'!$G$2:$G$6",
      valRef: "'Raw Data'!$H$2:$H$6",
    })
    .legend('right')
    .dataLabels({ showPercent: true, showCatName: true })
    .anchor({ col: 0, row: 21 }, { col: 6, row: 36 });
});

// ---------------------------------------------------------------------------
// 9. Customer growth line chart
// ---------------------------------------------------------------------------
dash.addChart('line', (b) => {
  b.title('Customer Growth')
    .addSeries({
      name: 'Customers',
      catRef: "'Raw Data'!$A$2:$A$13",
      valRef: "'Raw Data'!$E$2:$E$13",
      lineColor: '6366F1',
      lineWidth: 28575,
      marker: 'diamond',
    })
    .legend('bottom')
    .valAxis({
      title: 'Customers',
      numFmt: '#,##0',
      majorGridlines: true,
      min: 1000,
    })
    .catAxis({ title: 'Month' })
    .anchor({ col: 6, row: 21 }, { col: 12, row: 36 });
});

// -- Column widths on dashboard --
for (let c = 1; c <= 12; c++) dash.setColumnWidth(c, 10);

// ---------------------------------------------------------------------------
// 10. Document properties
// ---------------------------------------------------------------------------
wb.docProperties = {
  title: 'Business Dashboard 2025',
  creator: 'modern-xlsx chart-dashboard example',
};

// ---------------------------------------------------------------------------
// 11. Write output
// ---------------------------------------------------------------------------
const buffer = await wb.toBuffer();
writeFileSync('chart-dashboard.xlsx', buffer);

console.log('Created chart-dashboard.xlsx');
console.log(`  Sheets: ${wb.sheetNames.join(', ')}`);
console.log(`  Charts: 4 (bar, line x2, pie)`);
console.log(`  Data: 12 months x 5 metrics + 5 products`);
console.log(`  KPIs: Total Revenue $${totalRevenue.toLocaleString()}, Margin ${(marginPct * 100).toFixed(1)}%`);
console.log('\nOpen in Excel to see the charts rendered on the Dashboard sheet.');

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
