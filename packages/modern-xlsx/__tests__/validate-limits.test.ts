import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/index.js';

/** A workbook of one sheet `S` whose cells A1:A{cells} each have a dangling style index. */
function danglingStyles(cells: number): Workbook {
  const wb = new Workbook();
  const ws = wb.addSheet('S');
  for (let r = 1; r <= cells; r++) {
    const cell = ws.cell(`A${r}`);
    cell.value = r;
    cell.styleIndex = 999;
  }
  return wb;
}

/** Bytes of `message` and `location` text (UTF-8) the report stores. */
function storedText(report: ReturnType<Workbook['validate']>): number {
  const encoder = new TextEncoder();
  let bytes = 0;
  for (const issue of report.issues) {
    bytes += encoder.encode(issue.message).length + encoder.encode(issue.location).length;
  }
  return bytes;
}

describe('validate() report bounds', () => {
  it('stores at most 10,000 issues and counts every one', () => {
    const report = danglingStyles(50_000).validate();
    expect(report.errorCount).toBe(50_000);
    expect(report.warningCount + report.infoCount).toBe(1);
    expect(report.issues).toHaveLength(10_000);
    expect(report.omittedIssueCount).toBe(40_001);
    expect(report.isValid).toBe(false);
    expect(report.issues[0]?.location).toBe('S!A1');
  });

  it('repair() fixes every issue, stored or not', () => {
    const { report, repairCount } = danglingStyles(50_000).repair();
    // 50,000 style indices and the missing theme.
    expect(repairCount).toBe(50_001);
    expect(report.isValid).toBe(true);
    expect(report.omittedIssueCount).toBeUndefined();
  });

  it('stops storing at 16 MiB of text', () => {
    const data = danglingStyles(1_000).toJSON();
    const sheet = data.sheets[0];
    if (!sheet) throw new Error('no sheet');
    // A sheet name every cell issue repeats.
    sheet.name = 'x'.repeat(100_000);
    const report = new Workbook(data).validate();
    // The sheet name and the 1,000 style indices.
    expect(report.errorCount).toBe(1_001);
    const stored = report.issues.length;
    expect(stored).toBeLessThan(1_000);
    expect(stored + (report.omittedIssueCount ?? 0)).toBe(
      report.errorCount + report.warningCount + report.infoCount,
    );
    const text = storedText(report);
    expect(text).toBeLessThanOrEqual(16 * 1024 * 1024);
    // One more cell issue would not fit.
    expect(text + 100_000).toBeGreaterThan(16 * 1024 * 1024);
    expect(report.isValid).toBe(false);
  });

  it('keeps the report shape of a workbook under the bounds', () => {
    const report = danglingStyles(3).validate();
    expect(Object.keys(report).sort()).toEqual([
      'errorCount',
      'infoCount',
      'isValid',
      'issues',
      'warningCount',
    ]);
    expect(report.issues).toHaveLength(report.errorCount + report.warningCount + report.infoCount);
    expect(report.omittedIssueCount).toBeUndefined();
  });
});
