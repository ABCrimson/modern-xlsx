import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';
import type { Worksheet } from '../src/workbook.js';

/**
 * Read-fidelity suite (1.2.x): build a workbook exercising the advanced OOXML
 * features, write it to bytes, read it back, and assert each feature survives a
 * full write -> read round-trip through the WASM pipeline.
 */

function mustGetSheet(wb: Workbook, name: string): Worksheet {
  const ws = wb.getSheet(name);
  if (!ws) throw new Error(`expected sheet "${name}"`);
  return ws;
}

describe('read fidelity — advanced features survive round-trip', () => {
  it('preserves hyperlinks, merges, autofilter, tab color, comments, defined names, and visibility', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'Name';
    ws.cell('B1').value = 'Link';
    ws.cell('A2').value = 'Example';
    ws.cell('B2').value = 'site';
    ws.cell('B3').value = 'mail';
    ws.cell('B4').value = 'internal';

    // Hyperlinks: external URL (+ display/tooltip), email, and internal ref.
    ws.addHyperlink('B2', 'https://example.com/', { display: 'Example', tooltip: 'Open example' });
    ws.addHyperlink('B3', 'mailto:hi@example.com');
    ws.addHyperlink('B4', 'Data!A1', { display: 'jump' });

    ws.addMergeCell('A1:B1');
    ws.autoFilter = 'A1:B4';
    ws.tabColor = '#FF8800';
    ws.addComment('A2', 'Reviewer', 'Check this row');
    ws.groupRows(2, 4, 1);
    ws.groupColumns(2, 2, 1);

    const hidden = wb.addSheet('Secret');
    hidden.cell('A1').value = 'hidden value';
    hidden.state = 'hidden';

    wb.addNamedRange('MyRange', 'Data!$A$1:$B$4');

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = mustGetSheet(wb2, 'Data');

    // Hyperlinks
    const links = ws2.hyperlinks;
    const b2 = links.find((l) => l.cellRef === 'B2');
    expect(b2, 'B2 hyperlink should round-trip').toBeDefined();
    expect(b2?.location).toBe('https://example.com/');
    expect(b2?.tooltip).toBe('Open example');
    expect(links.find((l) => l.cellRef === 'B3')?.location).toBe('mailto:hi@example.com');
    expect(links.find((l) => l.cellRef === 'B4')?.location).toBe('Data!A1');

    // Merged cells
    expect(ws2.mergeCells).toContain('A1:B1');

    // Auto filter
    expect(ws2.autoFilter?.range).toBe('A1:B4');

    // Tab color (stored as ARGB/RGB hex — compare case-insensitively, ignoring '#'/alpha)
    expect(ws2.tabColor?.toUpperCase().replace('#', '')).toContain('FF8800');

    // Comment
    expect(ws2.comments.some((c) => c.cellRef === 'A2' && c.text.includes('Check this row'))).toBe(
      true,
    );

    // Defined name (workbook-scoped)
    expect(wb2.getNamedRange('MyRange')?.value).toContain('$A$1');

    // Sheet visibility
    expect(mustGetSheet(wb2, 'Secret').state).toBe('hidden');
  });

  it('preserves row/column outline grouping levels', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Grouped');
    for (let r = 1; r <= 5; r++) ws.cell(`A${r}`).value = `row ${r}`;
    ws.groupRows(2, 4, 1);

    const wb2 = await readBuffer(await wb.toBuffer());
    const ws2 = mustGetSheet(wb2, 'Grouped');
    const grouped = ws2.data.worksheet.rows.filter((row) => (row.outlineLevel ?? 0) > 0);
    expect(grouped.length).toBeGreaterThanOrEqual(3);
  });
});
