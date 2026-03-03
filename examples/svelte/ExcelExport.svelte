<!--
  Svelte 5 — Export data to XLSX using modern-xlsx.

  Usage:
    <ExcelExport data={myData} filename="report.xlsx" />
-->

<script lang="ts">
  import {
    Workbook,
    StyleBuilder,
    writeBlob,
    encodeCellRef,
  } from 'modern-xlsx';
  import { xlsx } from './xlsx.svelte';

  interface Props {
    data: Record<string, unknown>[];
    filename?: string;
    sheetName?: string;
  }

  let { data, filename = 'export.xlsx', sheetName = 'Sheet1' }: Props = $props();

  const { ready } = xlsx();
  let exporting = $state(false);

  async function handleExport() {
    if (!ready || data.length === 0) return;
    exporting = true;

    try {
      const wb = new Workbook();
      const ws = wb.addSheet(sheetName);

      const headers = Object.keys(data[0]);

      const headerStyle = new StyleBuilder()
        .font({ bold: true, color: 'FFFFFF' })
        .fill({ pattern: 'solid', fgColor: '2563EB' })
        .alignment({ horizontal: 'center' })
        .build(wb.styles);

      headers.forEach((header, col) => {
        const ref = encodeCellRef(0, col);
        ws.cell(ref).value = header;
        ws.cell(ref).styleIndex = headerStyle;
      });

      data.forEach((row, rowIdx) => {
        headers.forEach((header, col) => {
          const ref = encodeCellRef(rowIdx + 1, col);
          const value = row[header];
          if (value !== null && value !== undefined) {
            ws.cell(ref).value = value as string | number | boolean;
          }
        });
      });

      const blob = writeBlob(wb);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
    } finally {
      exporting = false;
    }
  }
</script>

<button onclick={handleExport} disabled={!ready || exporting}>
  {!ready ? 'Loading...' : exporting ? 'Exporting...' : `Export ${filename}`}
</button>
