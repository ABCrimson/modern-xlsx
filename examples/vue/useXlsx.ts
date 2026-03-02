/**
 * Vue 3 composable — Initialize modern-xlsx WASM and provide read/write helpers.
 *
 * Usage:
 *   const { ready, readFile, createWorkbook, downloadBlob } = useXlsx();
 */

import { ref, onMounted } from 'vue';
import {
  Workbook,
  initWasm,
  readBuffer,
  writeBlob,
} from 'modern-xlsx';

export function useXlsx() {
  const ready = ref(false);

  onMounted(async () => {
    await initWasm();
    ready.value = true;
  });

  async function readFile(file: File | Uint8Array): Promise<Workbook> {
    const data = file instanceof File
      ? new Uint8Array(await file.arrayBuffer())
      : file;
    return readBuffer(data);
  }

  function createWorkbook(): Workbook {
    return new Workbook();
  }

  function downloadBlob(wb: Workbook, filename = 'export.xlsx'): void {
    const blob = writeBlob(wb);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  return { ready, readFile, createWorkbook, downloadBlob };
}
