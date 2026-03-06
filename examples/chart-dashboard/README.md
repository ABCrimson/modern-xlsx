# Chart Dashboard

Creates a multi-sheet workbook with raw data and a dashboard featuring four
charts -- open in Excel or LibreOffice to see the visualizations.

## What It Builds

### Raw Data sheet
- 12 months of business metrics: Revenue, Costs, Profit, Customers
- Product revenue mix (5 product lines)
- Styled headers, currency formatting, frozen header row

### Dashboard sheet
- **KPI row** -- Total Revenue, Total Costs, Net Profit, Profit Margin
- **Bar chart** -- Monthly Revenue vs Costs (clustered)
- **Line chart** -- Profit trend with circle markers
- **Pie chart** -- Revenue breakdown by product (with percentage labels)
- **Line chart** -- Customer growth with diamond markers

## Usage

```bash
npm install
node index.mjs
```

This produces `chart-dashboard.xlsx`. Open it in Excel to see all four charts
rendered on the Dashboard sheet.

## Key APIs Used

| API | Purpose |
|-----|---------|
| `ws.addChart(type, cb)` | Add charts via the fluent ChartBuilder |
| `.addSeries({ catRef, valRef })` | Define data series with cell references |
| `.grouping('clustered')` | Set bar chart grouping mode |
| `.legend('bottom')` | Position the legend |
| `.valAxis({ title, numFmt })` | Configure value axis with formatting |
| `.dataLabels({ showPercent })` | Show percentage labels on pie chart |
| `.anchor(from, to)` | Position chart on the worksheet grid |
| `StyleBuilder` | KPI card styling, headers, number formats |
| `ws.addMergeCell()` | Merge cells for KPI cards and title |

## Chart Types Available

modern-xlsx supports 10 chart types via the `ChartBuilder`:

| Type | Description |
|------|-------------|
| `bar` | Horizontal or vertical bars |
| `col` | Vertical columns |
| `line` | Line with optional markers |
| `pie` | Pie chart |
| `doughnut` | Doughnut (pie with hole) |
| `scatter` | XY scatter plot |
| `area` | Filled area chart |
| `radar` | Radar/spider chart |
| `bubble` | Bubble chart |
| `stock` | Stock (OHLC) chart |
