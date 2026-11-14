/**
 * React hook — Initialize modern-xlsx WASM and provide read/write helpers.
 *
 * Usage:
 *   const { ready, error, readFile, createWorkbook, downloadBlob } = useXlsx();
 */

import { initWasm, readBuffer, Workbook, writeBlob } from 'modern-xlsx';
import { useCallback, useEffect, useRef, useState } from 'react';

interface UseXlsxReturn {
  /** Whether the WASM module is initialized and ready. */
  ready: boolean;
  /** The initialization failure, if `initWasm()` rejected; otherwise null. */
  error: Error | null;
  /** Read an XLSX file (from drag-drop, file input, or fetch). */
  readFile: (file: File | Uint8Array) => Promise<Workbook>;
  /** Create a new empty Workbook. */
  createWorkbook: () => Workbook;
  /** Convert a Workbook to a Blob and trigger download. */
  downloadBlob: (wb: Workbook, filename?: string) => void;
}

export function useXlsx(): UseXlsxReturn {
  const [ready, setReady] = useState(false);
  const [error, setError] = useState<Error | null>(null);
  const initRef = useRef(false);

  useEffect(() => {
    if (initRef.current) return;
    initRef.current = true;
    initWasm()
      .then(() => setReady(true))
      .catch((e: unknown) => setError(e instanceof Error ? e : new Error(String(e))));
  }, []);

  const readFile = useCallback(async (file: File | Uint8Array): Promise<Workbook> => {
    const data = file instanceof File ? new Uint8Array(await file.arrayBuffer()) : file;
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

  return { ready, error, readFile, createWorkbook, downloadBlob };
}
