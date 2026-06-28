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

    // Hyperlinks. NOTE on fidelity scope:
    //  - Internal links (B4 -> `Data!A1`) are emitted as the OOXML-standard
    //    `<hyperlink ref location>` and are fully Excel-interoperable.
    //  - External URL (B2) and email (B3) links currently round-trip through the
    //    library's own `location` representation, NOT the OOXML-standard external
    //    relationship form (`r:id` -> sheetN.xml.rels `TargetMode="External"`).
    //    So these assertions verify the value survives *our* write->read pipeline;
    //    full external-relationship interop with files authored by Excel is a
    //    known limitation tracked for a later 1.2.x release. See the comment at
    //    the `<hyperlinks>` writer in crates/.../worksheet/writer.rs.
    ws.addHyperlink('B2', 'https://example.com/', { display: 'Example', tooltip: 'Open example' });
    ws.addHyperlink('B3', 'mailto:hi@example.com');
    ws.addHyperlink('B4', 'Data!A1', { display: 'jump' });

    ws.addMergeCell('A1:B1');
    ws.autoFilter = 'A1:B4';
    // Colors are written verbatim (the library's documented contract); use the
    // canonical ARGB-hex-without-'#' form so the output is valid for Excel.
    ws.tabColor = 'FFFF8800';
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

    // Hyperlinks. B4 (internal) is a true OOXML round-trip; B2/B3 (external/email)
    // verify the library-convention `location` round-trip only — see the note at
    // the write site above for why standard external-relationship interop differs.
    const links = ws2.hyperlinks;
    expect(links.find((l) => l.cellRef === 'B4')?.location).toBe('Data!A1');
    const b2 = links.find((l) => l.cellRef === 'B2');
    expect(b2, 'B2 hyperlink should round-trip through our pipeline').toBeDefined();
    expect(b2?.location).toBe('https://example.com/');
    expect(b2?.tooltip).toBe('Open example');
    expect(links.find((l) => l.cellRef === 'B3')?.location).toBe('mailto:hi@example.com');

    // Merged cells
    expect(ws2.mergeCells).toContain('A1:B1');

    // Auto filter
    expect(ws2.autoFilter?.range).toBe('A1:B4');

    // Tab color: written verbatim, so it round-trips exactly.
    expect(ws2.tabColor).toBe('FFFF8800');

    // Comment
    expect(ws2.comments.some((c) => c.cellRef === 'A2' && c.text.includes('Check this row'))).toBe(
      true,
    );

    // Defined name (workbook-scoped)
    expect(wb2.getNamedRange('MyRange')?.value).toContain('$A$1');

    // Sheet visibility
    expect(mustGetSheet(wb2, 'Secret').state).toBe('hidden');
  });

  it('sanitizes XML-1.0-illegal characters in cell text on the default write path', async () => {
    // The default toBuffer() path (non-streaming writer) must never emit a
    // character illegal in XML 1.0, or the produced file is corrupt and Excel
    // rejects it. C0 control chars other than tab/LF/CR, plus the U+FFFE/U+FFFF
    // noncharacters, are dropped; legal whitespace and ordinary text survive.
    //
    // The illegal chars are built via String.fromCharCode so no literal
    // noncharacter bytes live in this source file (they trip some tooling).
    const C0 = String.fromCharCode(0x01);
    const FFFE = String.fromCharCode(0xfffe);
    const FFFF = String.fromCharCode(0xffff);
    const wb = new Workbook();
    const ws = wb.addSheet('Dirty');
    ws.cell('A1').value = `abc${C0}de${FFFE}f${FFFF}g\th`;
    ws.cell('A2').value = 'clean text';

    const buf = await wb.toBuffer();
    // The serialized parts must be parseable XML — no raw illegal bytes leaked,
    // so the read-back below succeeds and the illegal chars are gone.
    const wb2 = await readBuffer(buf);
    const ws2 = mustGetSheet(wb2, 'Dirty');
    expect(ws2.cell('A1').value).toBe('abcdefg\th');
    expect(ws2.cell('A2').value).toBe('clean text');
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
