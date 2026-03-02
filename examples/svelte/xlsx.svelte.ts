/**
 * Svelte 5 rune — Initialize modern-xlsx WASM and provide helpers.
 *
 * Usage:
 *   import { xlsx } from './xlsx.svelte';
 *   const { ready, readFile, createWorkbook, downloadBlob } = xlsx();
 */

import {
  Workbook,
  initWasm,
  readBuffer,
  writeBlob,
} from 'modern-xlsx';

export function xlsx() {
  let ready = $state(false);

  $effect(() => {
    initWasm().then(() => {
      ready = true;
    });
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

  return {
    get ready() { return ready; },
    readFile,
    createWorkbook,
    downloadBlob,
  };
}
