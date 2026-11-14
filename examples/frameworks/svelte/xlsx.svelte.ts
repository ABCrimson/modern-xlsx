/**
 * Svelte 5 rune — Initialize modern-xlsx WASM and provide helpers.
 *
 * Usage:
 *   import { xlsx } from './xlsx.svelte';
 *   const x = xlsx(); // read x.ready and x.error reactively; do not destructure them
 *   const { readFile, createWorkbook, downloadBlob } = x;
 */

import { initWasm, readBuffer, Workbook, writeBlob } from 'modern-xlsx';

export function xlsx() {
  let ready = $state(false);
  let error = $state<Error | null>(null);

  $effect(() => {
    initWasm()
      .then(() => {
        ready = true;
      })
      .catch((e: unknown) => {
        error = e instanceof Error ? e : new Error(String(e));
      });
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

  return {
    get ready() {
      return ready;
    },
    get error() {
      return error;
    },
    readFile,
    createWorkbook,
    downloadBlob,
  };
}
