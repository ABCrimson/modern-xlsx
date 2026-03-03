import { beforeAll, describe, expect, it } from 'vitest';
import { initWasm, Workbook } from '../src/index';

beforeAll(async () => {
  await initWasm();
});

describe('OOXML Validation & Compliance', () => {
  it('validates a minimal workbook as clean', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    const report = wb.validate();
    expect(report.isValid).toBe(true);
    expect(report.errorCount).toBe(0);
  });

  it('detects dangling style index on a cell', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'test';
    ws.cell('A1').styleIndex = 999;
    const report = wb.validate();
    expect(report.isValid).toBe(false);
    expect(report.errorCount).toBeGreaterThan(0);
    const issue = report.issues.find(
      (i) => i.category === 'styleIndex' && i.message.includes('999'),
    );
    expect(issue).toBeDefined();
    expect(issue?.autoFixable).toBe(true);
  });

  it('detects overlapping merge regions', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addMergeCell('A1:B2');
    ws.addMergeCell('B2:C3');
    const report = wb.validate();
    expect(report.isValid).toBe(false);
    expect(report.issues.some((i) => i.category === 'mergeCell')).toBe(true);
  });

  it('warns about missing theme colors (info level)', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    const report = wb.validate();
    // New workbooks don't have theme colors — should be info, not error
    const themeIssue = report.issues.find((i) => i.category === 'theme');
    if (themeIssue) {
      expect(themeIssue.severity).toBe('info');
    }
  });

  it('detects duplicate sheet names (case-insensitive)', () => {
    const wb = new Workbook({
      sheets: [
        { name: 'Data', worksheet: emptyWorksheet() },
        { name: 'data', worksheet: emptyWorksheet() },
      ],
      dateSystem: 'date1900',
      styles: defaultStyles(),
    });
    const report = wb.validate();
    expect(report.isValid).toBe(false);
    expect(report.issues.some((i) => i.message.includes('Duplicate'))).toBe(true);
  });

  it('repairs dangling style indices', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 42;
    ws.cell('A1').styleIndex = 999;

    const { workbook, report, repairCount } = wb.repair();
    expect(repairCount).toBeGreaterThan(0);
    expect(report.isValid).toBe(true);
    expect(workbook.getSheet('Sheet1')?.cell('A1').styleIndex).toBe(0);
  });

  it('repairs missing default styles', () => {
    const wb = new Workbook({
      sheets: [{ name: 'Sheet1', worksheet: emptyWorksheet() }],
      dateSystem: 'date1900',
      styles: {
        numFmts: [],
        fonts: [],
        fills: [],
        borders: [],
        cellXfs: [],
      },
    });

    const reportBefore = wb.validate();
    expect(reportBefore.isValid).toBe(false);

    const { workbook, report, repairCount } = wb.repair();
    expect(repairCount).toBeGreaterThan(0);
    expect(report.isValid).toBe(true);
    expect(workbook.styles.fonts.length).toBeGreaterThan(0);
    expect(workbook.styles.fills.length).toBeGreaterThanOrEqual(2);
    expect(workbook.styles.cellXfs.length).toBeGreaterThan(0);
  });

  it('repairs bad metadata dates', () => {
    const wb = new Workbook({
      sheets: [{ name: 'Sheet1', worksheet: emptyWorksheet() }],
      dateSystem: 'date1900',
      styles: defaultStyles(),
      docProperties: {
        created: 'not-a-date',
        title: '   ', // whitespace-only
      },
    });

    const { workbook, repairCount } = wb.repair();
    expect(repairCount).toBeGreaterThan(0);
    // serde skips None fields (undefined in JS), not null
    expect(workbook.docProperties?.created).toBeUndefined();
    expect(workbook.docProperties?.title).toBeUndefined();
  });

  it('repaired workbook can be written and read back', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Hello';
    ws.cell('A1').styleIndex = 999;

    const { workbook } = wb.repair();
    const buffer = await workbook.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(0);
  });

  it('validation report has correct structure', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    const report = wb.validate();

    expect(typeof report.isValid).toBe('boolean');
    expect(typeof report.errorCount).toBe('number');
    expect(typeof report.warningCount).toBe('number');
    expect(typeof report.infoCount).toBe('number');
    expect(Array.isArray(report.issues)).toBe(true);

    for (const issue of report.issues) {
      expect(typeof issue.severity).toBe('string');
      expect(typeof issue.category).toBe('string');
      expect(typeof issue.message).toBe('string');
      expect(typeof issue.location).toBe('string');
      expect(typeof issue.suggestion).toBe('string');
      expect(typeof issue.autoFixable).toBe('boolean');
    }
  });
});

// Helpers

function emptyWorksheet() {
  return {
    dimension: null,
    rows: [],
    mergeCells: [],
    autoFilter: null,
    frozenPane: null,
    columns: [],
  };
}

function defaultStyles() {
  return {
    numFmts: [],
    fonts: [
      {
        name: 'Aptos',
        size: 11,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        color: null,
      },
    ],
    fills: [
      { patternType: 'none', fgColor: null, bgColor: null },
      { patternType: 'gray125', fgColor: null, bgColor: null },
    ],
    borders: [{ left: null, right: null, top: null, bottom: null }],
    cellXfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }],
  };
}
