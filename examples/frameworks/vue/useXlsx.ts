/**
 * Vue 3 composable — Initialize modern-xlsx WASM and provide read/write helpers.
 *
 * Usage:
 *   const { ready, error, readFile, createWorkbook, downloadBlob } = useXlsx();
 */

import { initWasm, readBuffer, Workbook, writeBlob } from 'modern-xlsx';
import { onMounted, ref } from 'vue';

export function useXlsx() {
  const ready = ref(false);
  const error = ref<Error | null>(null);

  onMounted(async () => {
    try {
      await initWasm();
      ready.value = true;
    } catch (e) {
      error.value = e instanceof Error ? e : new Error(String(e));
    }
  });

  async function readFile(file: File | Uint8Array): Promise<Workbook> {
    const data = file instanceof File ? new Uint8Array(await file.arrayBuffer()) : file;
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

  return { ready, error, readFile, createWorkbook, downloadBlob };
}
