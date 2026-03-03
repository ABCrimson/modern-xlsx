import { describe, expect, it } from 'vitest';
import {
  columnToLetter,
  decodeCellRef,
  letterToColumn,
  readBuffer,
  Workbook,
} from '../src/index.js';

describe('0.1.6 — Edge Cases', () => {
  // ---------------------------------------------------------------------------
  // 1. Column letter boundaries
  // ---------------------------------------------------------------------------
  describe('column letter boundaries', () => {
    it('converts boundary column indices to letters', () => {
      expect(columnToLetter(0)).toBe('A');
      expect(columnToLetter(25)).toBe('Z');
      expect(columnToLetter(26)).toBe('AA');
      expect(columnToLetter(51)).toBe('AZ');
      expect(columnToLetter(52)).toBe('BA');
      expect(columnToLetter(701)).toBe('ZZ');
      expect(columnToLetter(702)).toBe('AAA');
      expect(columnToLetter(16383)).toBe('XFD');
    });

    it('converts boundary letters back to column indices', () => {
      expect(letterToColumn('A')).toBe(0);
      expect(letterToColumn('Z')).toBe(25);
      expect(letterToColumn('AA')).toBe(26);
      expect(letterToColumn('AZ')).toBe(51);
      expect(letterToColumn('BA')).toBe(52);
      expect(letterToColumn('ZZ')).toBe(701);
      expect(letterToColumn('AAA')).toBe(702);
      expect(letterToColumn('XFD')).toBe(16383);
    });

    it('roundtrips all boundary values', () => {
      const boundaries = [0, 25, 26, 51, 52, 701, 702, 16383];
      for (const col of boundaries) {
        expect(letterToColumn(columnToLetter(col))).toBe(col);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // 2. Cell reference at max Excel dimensions
  // ---------------------------------------------------------------------------
  describe('cell reference at max Excel dimensions', () => {
    it('decodes XFD1048576 (max Excel cell)', () => {
      const ref = decodeCellRef('XFD1048576');
      expect(ref.col).toBe(16383);
      expect(ref.row).toBe(1048575);
    });

    it('decodes A1 (min Excel cell)', () => {
      const ref = decodeCellRef('A1');
      expect(ref.col).toBe(0);
      expect(ref.row).toBe(0);
    });
  });

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
      expect(ws2?.cell('A4').value).toBeCloseTo(1e-15);
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

      const buffer = await wb.toBuffer();
      const wb2 = await readBuffer(buffer);
      const ws2 = wb2.getSheet('Dates');

      expect(ws2?.cell('A1').value).toBe(0);
      expect(ws2?.cell('A2').value).toBe(1);
      expect(ws2?.cell('A3').value).toBe(60);
      expect(ws2?.cell('A4').value).toBe(61);
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
      // CRLF may be normalized to LF by the XML parser — accept either
      const a2 = ws2?.cell('A2').value;
      expect(a2 === 'line1\r\nline2' || a2 === 'line1\nline2').toBe(true);
      expect(ws2?.cell('A3').value).toBe('col1\tcol2');
    });
  });
});
