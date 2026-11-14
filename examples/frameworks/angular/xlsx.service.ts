/**
 * Angular service — Initialize modern-xlsx WASM and provide helpers.
 *
 * Usage:
 *   private readonly xlsx = inject(XlsxService);
 *   await this.xlsx.waitForReady();
 *   const wb = this.xlsx.createWorkbook();
 */

import { Injectable, signal } from '@angular/core';
import { initWasm, readBuffer, Workbook, writeBlob } from 'modern-xlsx';

@Injectable({ providedIn: 'root' })
export class XlsxService {
  readonly ready = signal(false);
  readonly error = signal<Error | null>(null);
  private initPromise: Promise<void> | null = null;

  constructor() {
    this.initPromise = initWasm().then(
      () => {
        this.ready.set(true);
      },
      (e: unknown) => {
        this.error.set(e instanceof Error ? e : new Error(String(e)));
        // Rethrow so waitForReady() still rejects for callers that await it.
        throw e;
      },
    );
    // Mark handled: the error signal reports it; awaiting callers still see the rejection.
    this.initPromise.catch(() => {});
  }

  async waitForReady(): Promise<void> {
    await this.initPromise;
  }

  async readFile(file: File | Uint8Array): Promise<Workbook> {
    await this.waitForReady();
    const data = file instanceof File ? new Uint8Array(await file.arrayBuffer()) : file;
    return readBuffer(data);
  }

  createWorkbook(): Workbook {
    return new Workbook();
  }

  downloadBlob(wb: Workbook, filename = 'export.xlsx'): void {
    const blob = writeBlob(wb);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }
}
