<!--
  Vue 3 — Export data to XLSX using modern-xlsx.

  Usage:
    <ExcelExport :data="myData" filename="report.xlsx" />
-->

<script setup lang="ts">
import { ref } from 'vue';
import {
  Workbook,
  StyleBuilder,
  writeBlob,
  encodeCellRef,
} from 'modern-xlsx';
import { useXlsx } from './useXlsx';

interface Props {
  data: Record<string, unknown>[];
  filename?: string;
  sheetName?: string;
}

const props = withDefaults(defineProps<Props>(), {
  filename: 'export.xlsx',
  sheetName: 'Sheet1',
});

const { ready } = useXlsx();
const exporting = ref(false);

async function handleExport() {
  if (!ready.value || props.data.length === 0) return;
  exporting.value = true;

  try {
    const wb = new Workbook();
    const ws = wb.addSheet(props.sheetName);

    const headers = Object.keys(props.data[0]);

    const headerStyle = new StyleBuilder()
      .font({ bold: true, color: 'FFFFFF' })
      .fill({ pattern: 'solid', fgColor: '2563EB' })
      .alignment({ horizontal: 'center' })
      .build(wb.styles);

    headers.forEach((header, col) => {
      const ref = encodeCellRef({ row: 0, col });
      ws.cell(ref).value = header;
      ws.cell(ref).styleIndex = headerStyle;
    });

    props.data.forEach((row, rowIdx) => {
      headers.forEach((header, col) => {
        const ref = encodeCellRef({ row: rowIdx + 1, col });
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
    a.download = props.filename;
    a.click();
    URL.revokeObjectURL(url);
  } finally {
    exporting.value = false;
  }
}
</script>

<template>
  <button @click="handleExport" :disabled="!ready || exporting">
    {{ !ready ? 'Loading...' : exporting ? 'Exporting...' : `Export ${filename}` }}
  </button>
</template>
