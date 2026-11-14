<!--
  Vue 3 — Export data to XLSX using modern-xlsx.

  Usage:
    <ExcelExport :data="myData" filename="report.xlsx" />
-->

<script setup lang="ts">
import { encodeCellRef, StyleBuilder, Workbook } from 'modern-xlsx';
import { ref } from 'vue';
import { useXlsx } from './useXlsx';

interface Props {
  data: Record<string, unknown>[];
  filename?: string;
  sheetName?: string;
}

const { data, filename = 'export.xlsx', sheetName = 'Sheet1' } = defineProps<Props>();

// biome-ignore lint/correctness/noUnusedVariables: used in the template; Biome's Vue template analysis is partial
const { ready, error, downloadBlob } = useXlsx();
const exporting = ref(false);

// biome-ignore lint/correctness/noUnusedVariables: bound by @click in the template; Biome's Vue template analysis is partial
async function handleExport() {
  if (!ready.value || data.length === 0) return;
  exporting.value = true;

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

    downloadBlob(wb, filename);
  } finally {
    exporting.value = false;
  }
}
</script>

<template>
  <button @click="handleExport" :disabled="!ready || exporting || error !== null">
    {{ error ? `Failed to load: ${error.message}` : !ready ? 'Loading...' : exporting ? 'Exporting...' : `Export ${filename}` }}
  </button>
</template>
