import { describe, expect, it } from 'vitest';
import XLSX from 'xlsx';
import { readBuffer, Workbook } from '../src/index.js';

describe('Edge Cases', () => {
  // ---------------------------------------------------------------------------
  // 3. Cell at high row number
  // ---------------------------------------------------------------------------
  describe('cell at high row number', () => {
    it('handles cells at high row numbers', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('HighRow');
      ws.cell('A10000').value = 'deep';
      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      expect(wb2.getSheet('HighRow')?.cell('A10000').value).toBe('deep');
    });
  });

  // ---------------------------------------------------------------------------
  // 4. Cell at high column
  // ---------------------------------------------------------------------------
  describe('cell at high column', () => {
    it('handles cells at column Z (26th column)', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('HighCol');
      ws.cell('Z1').value = 'far right';
      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      expect(wb2.getSheet('HighCol')?.cell('Z1').value).toBe('far right');
    });
  });

  // ---------------------------------------------------------------------------
  // 5. Unicode in cell values
  // ---------------------------------------------------------------------------
  describe('unicode in cell values', () => {
    it('roundtrips unicode cell values', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('Unicode');
      ws.cell('A1').value = '🎉 celebration';
      ws.cell('A2').value = '中文测试';
      ws.cell('A3').value = 'مرحبا';
      ws.cell('A4').value = 'café résumé naïve';
      ws.cell('A5').value = 'Ⅳ ∑ ∞ √ π';
      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      const ws2 = wb2.getSheet('Unicode');
      expect(ws2?.cell('A1').value).toBe('🎉 celebration');
      expect(ws2?.cell('A2').value).toBe('中文测试');
      expect(ws2?.cell('A3').value).toBe('مرحبا');
      expect(ws2?.cell('A4').value).toBe('café résumé naïve');
      expect(ws2?.cell('A5').value).toBe('Ⅳ ∑ ∞ √ π');
    });
  });

  // ---------------------------------------------------------------------------
  // 6. Unicode in sheet names
  // ---------------------------------------------------------------------------
  describe('unicode in sheet names', () => {
    it('roundtrips unicode sheet names', async () => {
      const wb = new Workbook();
      wb.addSheet('日本語シート').cell('A1').value = 1;
      wb.addSheet('Données').cell('A1').value = 2;
      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      expect(wb2.sheetNames).toContain('日本語シート');
      expect(wb2.sheetNames).toContain('Données');
      expect(wb2.getSheet('日本語シート')?.cell('A1').value).toBe(1);
    });
  });

  // ---------------------------------------------------------------------------
  // 7. Special characters in string values (XML entities)
  // ---------------------------------------------------------------------------
  describe('special characters in string values', () => {
    it('roundtrips XML special characters in values', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('Escaping');
      ws.cell('A1').value = 'Tom & Jerry';
      ws.cell('A2').value = '2 < 3';
      ws.cell('A3').value = '5 > 4';
      ws.cell('A4').value = 'He said "hello"';
      ws.cell('A5').value = "It's fine";
      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      const ws2 = wb2.getSheet('Escaping');
      expect(ws2?.cell('A1').value).toBe('Tom & Jerry');
      expect(ws2?.cell('A2').value).toBe('2 < 3');
      expect(ws2?.cell('A3').value).toBe('5 > 4');
      expect(ws2?.cell('A4').value).toBe('He said "hello"');
      expect(ws2?.cell('A5').value).toBe("It's fine");
    });
  });

  // ---------------------------------------------------------------------------
  // 8. Empty string value
  // ---------------------------------------------------------------------------
  describe('empty string value', () => {
    it('roundtrips an empty string cell value', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('Empty');
      ws.cell('A1').value = '';
      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      const ws2 = wb2.getSheet('Empty');
      expect(ws2?.cell('A1').value).toBe('');
    });
  });

  // ---------------------------------------------------------------------------
  // 9. Very long string
  // ---------------------------------------------------------------------------
  describe('very long string', () => {
    it('roundtrips a 10,000 character string', async () => {
      const longStr = 'x'.repeat(10_000);
      const wb = new Workbook();
      const ws = wb.addSheet('Long');
      ws.cell('A1').value = longStr;
      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      const ws2 = wb2.getSheet('Long');
      const result = ws2?.cell('A1').value;
      expect(typeof result).toBe('string');
      expect((result as string).length).toBe(10_000);
      expect(result).toBe(longStr);
    });
  });

  // ---------------------------------------------------------------------------
  // 10. Numeric edge values
  // ---------------------------------------------------------------------------
  describe('numeric edge values', () => {
    it('roundtrips numeric edge values', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('Numbers');
      ws.cell('A1').value = 0;
      ws.cell('A2').value = -0;
      ws.cell('A3').value = 1e15;
      ws.cell('A4').value = 1e-15;
      ws.cell('A5').value = Number.MAX_SAFE_INTEGER;
      ws.cell('A6').value = -Number.MAX_SAFE_INTEGER;
      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      const ws2 = wb2.getSheet('Numbers');
      expect(ws2?.cell('A1').value).toBe(0);
      expect(ws2?.cell('A3').value).toBe(1e15);
      expect(ws2?.cell('A4').value).toBe(1e-15);
      expect(ws2?.cell('A5').value).toBe(Number.MAX_SAFE_INTEGER);
      expect(ws2?.cell('A6').value).toBe(-Number.MAX_SAFE_INTEGER);
    });
  });

  // ---------------------------------------------------------------------------
  // 11. Boolean values
  // ---------------------------------------------------------------------------
  describe('boolean values', () => {
    it('roundtrips true and false', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('Booleans');
      ws.cell('A1').value = true;
      ws.cell('A2').value = false;
      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      const ws2 = wb2.getSheet('Booleans');
      expect(ws2?.cell('A1').value).toBe(true);
      expect(ws2?.cell('A2').value).toBe(false);
    });
  });

  // ---------------------------------------------------------------------------
  // 12. Sparse cells
  // ---------------------------------------------------------------------------
  describe('sparse cells', () => {
    it('roundtrips sparse cells without filling gaps', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('Sparse');
      ws.cell('A1').value = 'first';
      ws.cell('Z100').value = 'middle';
      ws.cell('AA500').value = 'last';
      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      const ws2 = wb2.getSheet('Sparse');
      expect(ws2?.cell('A1').value).toBe('first');
      expect(ws2?.cell('Z100').value).toBe('middle');
      expect(ws2?.cell('AA500').value).toBe('last');

      // Verify only 3 rows contain data
      const populatedRows = ws2?.rows.filter((r) => r.cells.length > 0) ?? [];
      expect(populatedRows).toHaveLength(3);
    });
  });

  // ---------------------------------------------------------------------------
  // 13. Large SST (shared string table)
  // ---------------------------------------------------------------------------
  describe('large SST', () => {
    it('roundtrips 1000 unique strings', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('ManyStrings');
      const strings: string[] = [];
      for (let i = 0; i < 1000; i++) {
        const str = `unique_string_${i.toString().padStart(4, '0')}`;
        strings.push(str);
        ws.cell(`A${i + 1}`).value = str;
      }
      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      const ws2 = wb2.getSheet('ManyStrings');
      for (let i = 0; i < 1000; i++) {
        expect(ws2?.cell(`A${i + 1}`).value).toBe(strings[i]);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // 14. Date serial edge cases
  // ---------------------------------------------------------------------------
  describe('date serial edge cases', () => {
    it('writes date serial numbers with date format and roundtrips', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('Dates');

      // Create a date number format style
      const dateStyleIndex = wb.createStyle().numberFormat('yyyy-mm-dd').build(wb.styles);

      // Serial 0 = Jan 0, 1900 (Excel quirk)
      ws.cell('A1').value = 0;
      ws.cell('A1').styleIndex = dateStyleIndex;

      // Serial 1 = Jan 1, 1900
      ws.cell('A2').value = 1;
      ws.cell('A2').styleIndex = dateStyleIndex;

      // Serial 60 = Feb 29, 1900 (Lotus 1-2-3 bug — this date does not exist)
      ws.cell('A3').value = 60;
      ws.cell('A3').styleIndex = dateStyleIndex;

      // Serial 61 = Mar 1, 1900
      ws.cell('A4').value = 61;
      ws.cell('A4').styleIndex = dateStyleIndex;

      // Serial 59 = Feb 28, 1900
      ws.cell('A5').value = 59;
      ws.cell('A5').styleIndex = dateStyleIndex;

      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      const ws2 = wb2.getSheet('Dates');

      expect(ws2?.cell('A1').value).toBe(0);
      expect(ws2?.cell('A2').value).toBe(1);
      expect(ws2?.cell('A3').value).toBe(60);
      expect(ws2?.cell('A4').value).toBe(61);
      expect(ws2?.cell('A5').value).toBe(59);

      // dateValue maps serials to calendar dates; serial 60 (the non-existent
      // 1900-02-29) maps to the same day as serial 59.
      const iso = (ref: string) => ws2?.cell(ref).dateValue?.toISOString().slice(0, 10);
      expect(iso('A2')).toBe('1900-01-01');
      expect(iso('A5')).toBe('1900-02-28');
      expect(iso('A3')).toBe('1900-02-28');
      expect(iso('A4')).toBe('1900-03-01');
    });
  });

  // ---------------------------------------------------------------------------
  // 15. Newlines in cell values
  // ---------------------------------------------------------------------------
  describe('newlines in cell values', () => {
    it('roundtrips strings with newlines and tabs', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('Whitespace');
      ws.cell('A1').value = 'line1\nline2';
      ws.cell('A2').value = 'line1\r\nline2';
      ws.cell('A3').value = 'col1\tcol2';
      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      const ws2 = wb2.getSheet('Whitespace');
      expect(ws2?.cell('A1').value).toBe('line1\nline2');
      expect(ws2?.cell('A2').value).toBe('line1\r\nline2');
      expect(ws2?.cell('A3').value).toBe('col1\tcol2');
    });
  });

  // ---------------------------------------------------------------------------
  // 16. Reference, whitespace and length limits
  // ---------------------------------------------------------------------------
  describe('limits and whitespace', () => {
    it('roundtrips the last cell of the sheet', async () => {
      const wb = new Workbook();
      wb.addSheet('Max').cell('XFD1048576').value = 'max';
      const wb2 = await readBuffer(await wb.toBuffer());
      expect(wb2.getSheet('Max')?.cell('XFD1048576').value).toBe('max');
    });

    it('preserves leading and trailing spaces with xml:space="preserve"', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('Spaces');
      ws.cell('A1').value = '  lead';
      ws.cell('A2').value = 'trail  ';
      const buffer = await wb.toBuffer();
      const ws2 = (await readBuffer(buffer)).getSheet('Spaces');
      expect(ws2?.cell('A1').value).toBe('  lead');
      expect(ws2?.cell('A2').value).toBe('trail  ');

      const cfb = XLSX.CFB.read(buffer, { type: 'buffer' });
      const sst = XLSX.CFB.find(cfb, '/xl/sharedStrings.xml');
      expect(new TextDecoder().decode(sst?.content as Uint8Array)).toContain(
        '<t xml:space="preserve">',
      );
    });

    it('roundtrips a 32767-character string exactly', async () => {
      // Seeded LCG text without <, > or &. Varied on purpose: a single repeated
      // character would trip the 100:1 zip-bomb ratio guard.
      const alphabet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ';
      let state = 12345;
      let text = '';
      for (let i = 0; i < 32767; i++) {
        state = (Math.imul(state, 1103515245) + 12345) >>> 0;
        text += alphabet[(state >>> 16) % alphabet.length];
      }
      expect(text).toHaveLength(32767);
      const wb = new Workbook();
      wb.addSheet('Long').cell('A1').value = text;
      const wb2 = await readBuffer(await wb.toBuffer());
      expect(wb2.getSheet('Long')?.cell('A1').value).toBe(text);
    });

    // Open write defect: toBuffer() rejects a lone surrogate (code JSON, "JSON error: unexpected end of hex escape at line 1 column N") instead of writing U+FFFD; the target below is the fixed behaviour.
    it.fails('lone surrogate is written as U+FFFD', async () => {
      const wb = new Workbook();
      wb.addSheet('Surrogate').cell('A1').value = 'a\uD800b';
      const wb2 = await readBuffer(await wb.toBuffer());
      expect(wb2.getSheet('Surrogate')?.cell('A1').value).toBe('a\uFFFDb');
    });
  });
});
