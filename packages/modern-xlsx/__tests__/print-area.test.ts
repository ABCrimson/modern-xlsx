import { describe, expect, it } from 'vitest';
import { INVALID_ARGUMENT, Workbook } from '../src/index.js';

function twoSheets(): Workbook {
  const wb = new Workbook();
  wb.addSheet('Sheet1');
  wb.addSheet('My Sheet');
  wb.addSheet("O'Brien");
  return wb;
}

describe('print area and print titles are stored sheet-qualified', () => {
  it('qualifies an unqualified print area with a plain and a quoted sheet name', () => {
    const wb = twoSheets();
    wb.getSheet('Sheet1')?.setPrintArea('A1:B2');
    wb.getSheet('My Sheet')?.setPrintArea('$A$1:$B$2');
    wb.getSheet("O'Brien")?.setPrintArea('A1');
    expect(wb.getSheet('Sheet1')?.getPrintArea()).toBe('Sheet1!A1:B2');
    expect(wb.getSheet('My Sheet')?.getPrintArea()).toBe("'My Sheet'!$A$1:$B$2");
    expect(wb.getSheet("O'Brien")?.getPrintArea()).toBe("'O''Brien'!A1");
  });

  it('qualifies every unqualified area of a list and keeps a qualified one', () => {
    const wb = twoSheets();
    wb.setPrintArea(0, 'A1:B2,Sheet1!$D$1:$E$2');
    expect(wb.getPrintArea(0)).toBe('Sheet1!A1:B2,Sheet1!$D$1:$E$2');
    wb.setPrintArea(1, 'A:B,$3:$4');
    expect(wb.getPrintArea(1)).toBe("'My Sheet'!A:B,'My Sheet'!$3:$4");
  });

  it('stores a qualified value as it is, whichever sheet it names', () => {
    const wb = twoSheets();
    wb.setPrintArea(0, 'Sheet1!$A$1:$D$50');
    wb.setPrintArea(1, "'My Sheet'!$A$1:$B$2");
    wb.getSheet('Sheet1')?.setPrintArea("'My Sheet'!A1:B2");
    expect(wb.getPrintArea(0)).toBe("'My Sheet'!A1:B2");
    expect(wb.getPrintArea(1)).toBe("'My Sheet'!$A$1:$B$2");
  });

  it('qualifies print titles', () => {
    const wb = twoSheets();
    wb.setPrintTitles(0, '$1:$2');
    wb.setPrintTitles(1, '$A:$B,$1:$2');
    expect(wb.getPrintTitles(0)).toBe('Sheet1!$1:$2');
    expect(wb.getPrintTitles(1)).toBe("'My Sheet'!$A:$B,'My Sheet'!$1:$2");
  });

  it('clears with null', () => {
    const wb = twoSheets();
    wb.setPrintArea(0, 'A1:B2');
    wb.setPrintArea(0, null);
    expect(wb.getPrintArea(0)).toBeNull();
  });

  it.each([
    '',
    'garbage',
    'A1:',
    'A1,,B2',
    'A1:B2,',
    '=A1:B2',
    'MyName',
    '1+2',
    'Sheet1!A1:B2 extra',
  ])('throws INVALID_ARGUMENT for %j', (value) => {
    const wb = twoSheets();
    const thrown = (fn: () => void) => {
      try {
        fn();
      } catch (error) {
        return error;
      }
      return undefined;
    };
    for (const fn of [
      () => wb.setPrintArea(0, value),
      () => wb.setPrintTitles(0, value),
      () => wb.getSheet('Sheet1')?.setPrintArea(value),
    ]) {
      expect(thrown(fn)).toMatchObject({ name: 'ModernXlsxError', code: INVALID_ARGUMENT });
    }
    expect(wb.getPrintArea(0)).toBeNull();
    expect(wb.getPrintTitles(0)).toBeNull();
  });

  it('writes the qualified name into the workbook part', async () => {
    const { readZipText } = await import('./helpers/zip.js');
    const wb = twoSheets();
    wb.setPrintArea(1, 'A1:B2');
    const xml = readZipText(await wb.toBuffer())['xl/workbook.xml'];
    expect(xml).toContain('>&apos;My Sheet&apos;!A1:B2</definedName>');
  });
});

describe('addNamedRange with a built-in print name', () => {
  it('stores the value sheet-qualified', () => {
    const wb = twoSheets();
    wb.addNamedRange('_xlnm.Print_Area', 'A1:B2', 0);
    wb.addNamedRange('_xlnm.Print_Titles', '$1:$2', 1);
    expect(wb.getNamedRange('_xlnm.Print_Area')?.value).toBe('Sheet1!A1:B2');
    expect(wb.namedRanges.find((d) => d.name === '_xlnm.Print_Titles')?.value).toBe(
      "'My Sheet'!$1:$2",
    );
  });

  it('throws INVALID_ARGUMENT for a value that is not a reference list', () => {
    const wb = twoSheets();
    expect(() => wb.addNamedRange('_xlnm.Print_Area', 'garbage!!', 0)).toThrow(
      expect.objectContaining({ code: INVALID_ARGUMENT }),
    );
  });
});

describe('built-in print names in any letter case', () => {
  it.each(['_xlnm.print_area', '_XLNM.PRINT_AREA', '_xlnm.Print_area'])(
    'addNamedRange(%j) stores _xlnm.Print_Area qualified and readable',
    (name) => {
      const wb = twoSheets();
      wb.addNamedRange(name, 'A1:B2', 1);
      expect(wb.namedRanges).toEqual([
        { name: '_xlnm.Print_Area', value: "'My Sheet'!A1:B2", sheetId: 1 },
      ]);
      expect(wb.getPrintArea(1)).toBe("'My Sheet'!A1:B2");
      expect(wb.getSheet('My Sheet')?.getPrintArea()).toBe("'My Sheet'!A1:B2");
      expect(wb.validate().isValid).toBe(true);
    },
  );

  it.each(['_xlnm.print_titles', '_XLNM.PRINT_TITLES', '_xlnm.Print_titles'])(
    'addNamedRange(%j) stores _xlnm.Print_Titles qualified and readable',
    (name) => {
      const wb = twoSheets();
      wb.addNamedRange(name, '$1:$2', 0);
      expect(wb.namedRanges).toEqual([
        { name: '_xlnm.Print_Titles', value: 'Sheet1!$1:$2', sheetId: 0 },
      ]);
      expect(wb.getPrintTitles(0)).toBe('Sheet1!$1:$2');
    },
  );

  it('a variant spelling in the model is read by the getters and updated by the setters', () => {
    // A raw model, as a file or JSON gives it, holds the variant spellings.
    const wb = twoSheets();
    wb.toJSON().definedNames = [
      { name: '_xlnm.print_area', value: 'Sheet1!$A$1:$B$2', sheetId: 0 },
      { name: '_XLNM.PRINT_TITLES', value: 'Sheet1!$1:$1', sheetId: 0 },
      { name: '_xlnm.Print_area', value: "'My Sheet'!$C$3", sheetId: 1 },
    ];
    expect(wb.getPrintArea(0)).toBe('Sheet1!$A$1:$B$2');
    expect(wb.getSheet('Sheet1')?.getPrintArea()).toBe('Sheet1!$A$1:$B$2');
    expect(wb.getPrintTitles(0)).toBe('Sheet1!$1:$1');
    expect(wb.getPrintArea(1)).toBe("'My Sheet'!$C$3");

    wb.setPrintArea(0, 'C1:D2');
    wb.setPrintTitles(0, '$1:$2');
    wb.getSheet('My Sheet')?.setPrintArea('E5');
    // One entry each, in the canonical spelling.
    expect(wb.namedRanges).toEqual([
      { name: '_xlnm.Print_Area', value: 'Sheet1!C1:D2', sheetId: 0 },
      { name: '_xlnm.Print_Titles', value: 'Sheet1!$1:$2', sheetId: 0 },
      { name: '_xlnm.Print_Area', value: "'My Sheet'!E5", sheetId: 1 },
    ]);

    wb.setPrintArea(0, null);
    wb.getSheet('My Sheet')?.setPrintArea(null);
    expect(wb.namedRanges).toEqual([
      { name: '_xlnm.Print_Titles', value: 'Sheet1!$1:$2', sheetId: 0 },
    ]);
  });

  it('with two spellings on one sheet the last is read, and a set leaves one entry', () => {
    // 1.2.1 writes this shape (a variant addNamedRange, then setPrintArea);
    // Excel prints the last entry.
    const twoAreas = () => {
      const wb = twoSheets();
      wb.toJSON().definedNames = [
        { name: '_xlnm.print_area', value: 'Sheet1!$A$1', sheetId: 0 },
        { name: '_xlnm.Print_Area', value: 'Sheet1!$B$2:$C$3', sheetId: 0 },
      ];
      return wb;
    };
    const read = twoAreas();
    expect(read.getPrintArea(0)).toBe('Sheet1!$B$2:$C$3');
    expect(read.getSheet('Sheet1')?.getPrintArea()).toBe('Sheet1!$B$2:$C$3');

    const set = twoAreas();
    set.setPrintArea(0, 'D4');
    expect(set.namedRanges).toEqual([{ name: '_xlnm.Print_Area', value: 'Sheet1!D4', sheetId: 0 }]);

    const cleared = twoAreas();
    cleared.setPrintArea(0, null);
    expect(cleared.namedRanges).toEqual([]);
  });

  it('a second print area in another spelling on the same sheet is a duplicate', () => {
    const wb = twoSheets();
    wb.setPrintArea(0, 'A1:B2');
    expect(() => wb.addNamedRange('_xlnm.print_area', 'C1:D2', 0)).toThrow(
      expect.objectContaining({ code: INVALID_ARGUMENT }),
    );
    expect(wb.namedRanges).toHaveLength(1);
  });

  it('writes the canonical spelling into the workbook part', async () => {
    const { readZipText } = await import('./helpers/zip.js');
    const wb = twoSheets();
    wb.addNamedRange('_xlnm.print_area', 'A1:B2', 0);
    const xml = readZipText(await wb.toBuffer())['xl/workbook.xml'];
    expect(xml).toContain(
      '<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!A1:B2</definedName>',
    );
  });
});

describe('validate() on an unqualified built-in name', () => {
  it('reports an error for a raw-JSON workbook', () => {
    const wb = new Workbook({
      sheets: [
        {
          name: 'Thr',
          worksheet: {
            dimension: null,
            rows: [],
            mergeCells: [],
            autoFilter: null,
            frozenPane: null,
            columns: [],
          },
        },
      ],
      definedNames: [
        { name: '_xlnm.Print_Area', value: 'A1:B2', sheetId: 0 },
        { name: '_xlnm.Print_Titles', value: 'Thr!$1:$2', sheetId: 0 },
      ],
    });
    const report = wb.validate();
    const issues = report.issues.filter((i) => i.category === 'definedName');
    expect(report.isValid).toBe(false);
    expect(issues).toHaveLength(1);
    expect(issues[0]).toMatchObject({ severity: 'error', location: 'definedNames[0]' });
    expect(issues[0]?.message).toContain('Print_Area');
  });

  it.each(['_xlnm.print_area', '_XLNM.PRINT_TITLES', '_xlnm.Print_area'])(
    'reports an error for the variant spelling %j',
    (name) => {
      const wb = twoSheets();
      wb.toJSON().definedNames = [{ name, value: 'A1:B2', sheetId: 0 }];
      const report = wb.validate();
      expect(report.isValid).toBe(false);
      expect(
        report.issues
          .filter((i) => i.category === 'definedName')
          .map((i) => [i.severity, i.location]),
      ).toEqual([['error', 'definedNames[0]']]);
    },
  );
});
