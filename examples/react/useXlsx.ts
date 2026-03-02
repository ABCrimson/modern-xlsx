/**
 * React hook — Initialize modern-xlsx WASM and provide read/write helpers.
 *
 * Usage:
 *   const { ready, readFile, createWorkbook, downloadBlob } = useXlsx();
 */

import { useCallback, useEffect, useRef, useState } from 'react';
import {
  Workbook,
  initWasm,
  readBuffer,
  writeBlob,
} from 'modern-xlsx';

interface UseXlsxReturn {
  /** Whether the WASM module is initialized and ready. */
  ready: boolean;
  /** Read an XLSX file (from drag-drop, file input, or fetch). */
  readFile: (file: File | Uint8Array) => Promise<Workbook>;
  /** Create a new empty Workbook. */
  createWorkbook: () => Workbook;
  /** Convert a Workbook to a Blob and trigger download. */
  downloadBlob: (wb: Workbook, filename?: string) => void;
}

export function useXlsx(): UseXlsxReturn {
  const [ready, setReady] = useState(false);
  const initRef = useRef(false);

  useEffect(() => {
    if (initRef.current) return;
    initRef.current = true;
    initWasm().then(() => setReady(true));
  }, []);

  const readFile = useCallback(async (file: File | Uint8Array): Promise<Workbook> => {
    const data = file instanceof File
      ? new Uint8Array(await file.arrayBuffer())
      : file;
    return readBuffer(data);
  }, []);

  const createWorkbook = useCallback(() => new Workbook(), []);

  const downloadBlob = useCallback((wb: Workbook, filename = 'export.xlsx') => {
    const blob = writeBlob(wb);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }, []);

  return { ready, readFile, createWorkbook, downloadBlob };
}
