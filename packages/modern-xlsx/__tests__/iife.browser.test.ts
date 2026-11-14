import { describe, expect, it, vi } from 'vitest';

interface Cell {
  value: unknown;
}
interface Sheet {
  cell(ref: string): Cell;
}
interface ModernXlsxGlobal {
  initWasm(): Promise<void>;
  readBuffer(data: Uint8Array): Promise<{ getSheet(name: string): Sheet | undefined }>;
  Workbook: new () => { addSheet(name: string): Sheet; toBuffer(): Promise<Uint8Array> };
}

describe('IIFE bundle (dist/modern-xlsx.min.js)', () => {
  it('loads as a classic script, fetches its sibling WASM and round-trips', async () => {
    const warn = vi.spyOn(console, 'warn');
    try {
      const script = document.createElement('script');
      script.src = '/dist/modern-xlsx.min.js';
      const loaded = new Promise<void>((resolve, reject) => {
        script.addEventListener('load', () => resolve());
        script.addEventListener('error', () => reject(new Error('IIFE script failed to load')));
      });
      document.head.append(script);
      await loaded;

      const api = (globalThis as unknown as { ModernXlsx: ModernXlsxGlobal }).ModernXlsx;
      await expect(api.initWasm()).resolves.toBeUndefined();

      const fetched = performance
        .getEntriesByType('resource')
        .map((entry) => new URL(entry.name).pathname);
      expect(fetched.some((p) => p.endsWith('/dist/modern-xlsx.wasm'))).toBe(true);
      expect(warn).not.toHaveBeenCalled();

      const wb = new api.Workbook();
      wb.addSheet('Sheet1').cell('A1').value = 'iife';
      const back = await api.readBuffer(await wb.toBuffer());
      expect(back.getSheet('Sheet1')?.cell('A1').value).toBe('iife');
    } finally {
      warn.mockRestore();
    }
  });
});
