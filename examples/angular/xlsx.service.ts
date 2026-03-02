/**
 * Angular service — Initialize modern-xlsx WASM and provide helpers.
 *
 * Usage:
 *   constructor(private xlsx: XlsxService) {}
 *   await this.xlsx.waitForReady();
 *   const wb = this.xlsx.createWorkbook();
 */

import { Injectable, signal } from '@angular/core';
import {
  Workbook,
  initWasm,
  readBuffer,
  writeBlob,
} from 'modern-xlsx';

@Injectable({ providedIn: 'root' })
export class XlsxService {
  readonly ready = signal(false);
  private initPromise: Promise<void> | null = null;

  constructor() {
    this.initPromise = initWasm().then(() => {
      this.ready.set(true);
    });
  }

  async waitForReady(): Promise<void> {
    await this.initPromise;
  }

  async readFile(file: File | Uint8Array): Promise<Workbook> {
    await this.waitForReady();
    const data = file instanceof File
      ? new Uint8Array(await file.arrayBuffer())
      : file;
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
