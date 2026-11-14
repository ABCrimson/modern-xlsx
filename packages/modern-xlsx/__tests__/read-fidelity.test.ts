import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';
import type { Worksheet } from '../src/model/worksheet.js';
import { readZipText } from './helpers/zip.js';

/**
 * Read-fidelity suite: build a workbook exercising the advanced OOXML
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

    // Hyperlinks, both OOXML forms:
    //  - Internal links (B4 -> `Data!A1`) are written as `<hyperlink ref location>`.
    //  - External URL (B2) and email (B3) links go to `target` and are written as
    //    `<hyperlink ref r:id>` plus a sheet1.xml.rels relationship with
    //    `TargetMode="External"`, the form Excel writes and reads.
    ws.addHyperlink('B2', 'https://example.com/', { display: 'Example', tooltip: 'Open example' });
    ws.addHyperlink('B3', 'mailto:hi@example.com');
    ws.addHyperlink('B4', 'Data!A1', { display: 'jump' });

    ws.addMergeCell('A1:B1');
    ws.autoFilter = 'A1:B4';
    // Colours are written as uppercase ARGB (RRGGBB gets an FF alpha); this one is
    // already in that canonical form, so it reads back unchanged.
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

    // Hyperlinks: external targets read back from the sheet relationship into
    // `target`; the internal link keeps `location` and gets no relationship.
    const links = ws2.hyperlinks;
    const b4 = links.find((l) => l.cellRef === 'B4');
    expect(b4?.location).toBe('Data!A1');
    expect(b4?.target).toBeUndefined();
    const b2 = links.find((l) => l.cellRef === 'B2');
    expect(b2, 'B2 hyperlink should round-trip').toBeDefined();
    expect(b2?.target).toBe('https://example.com/');
    expect(b2?.location).toBeUndefined();
    expect(b2?.display).toBe('Example');
    expect(b2?.tooltip).toBe('Open example');
    const b3 = links.find((l) => l.cellRef === 'B3');
    expect(b3?.target).toBe('mailto:hi@example.com');
    expect(b3?.location).toBeUndefined();

    // The sheet rels are kept verbatim in preservedEntries: one external
    // relationship per external link, none for the internal one.
    const relsBytes = wb2.toJSON().preservedEntries?.['xl/worksheets/_rels/sheet1.xml.rels'];
    expect(relsBytes, 'sheet1 rels should be preserved').toBeDefined();
    const rels = new TextDecoder().decode(Uint8Array.from(relsBytes ?? []));
    expect(rels).toContain('TargetMode="External"');
    expect(rels).toContain('Target="https://example.com/"');
    expect(rels).toContain('Target="mailto:hi@example.com"');
    expect(rels.match(/relationships\/hyperlink"/g)).toHaveLength(2);
    expect(rels).not.toContain('Data!A1');

    // sheet1.xml references the B2 relationship by r:id and has no URL location.
    const sheetXml = readZipText(buf)['xl/worksheets/sheet1.xml'] ?? '';
    expect(sheetXml).toMatch(/<hyperlink ref="B2" r:id="rId[0-9]+"/);
    expect(sheetXml).not.toContain('location="https://');
    expect(sheetXml).toContain('<hyperlink ref="B4" location="Data!A1"');

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

  it('keeps a query-string URL exact through the external relationship', async () => {
    const wb = new Workbook();
    wb.addSheet('Data').addHyperlink('A1', 'https://e.com/?a=1&b=2');
    const wb2 = await readBuffer(await wb.toBuffer());
    expect(mustGetSheet(wb2, 'Data').hyperlinks[0]?.target).toBe('https://e.com/?a=1&b=2');
    const rels = readZipText(await wb2.toBuffer())['xl/worksheets/_rels/sheet1.xml.rels'] ?? '';
    expect(rels).toContain('Target="https://e.com/?a=1&amp;b=2"');
    expect(rels).not.toContain('&amp;amp;');
  });

  it('keeps same-sheet ranges in location and sends schemes to target', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    const ranges = ['A1:B2', 'A1:$B$2', 'B:D'];
    for (const [i, range] of ranges.entries()) ws.addHyperlink(`C${i + 1}`, range);
    ws.addHyperlink('D1', 'tel:911');
    ws.addHyperlink('D2', 'C:\\a.xlsx');
    for (const [i, range] of ranges.entries()) {
      expect(ws.hyperlinks[i]?.location).toBe(range);
      expect(ws.hyperlinks[i]?.target).toBeUndefined();
    }
    expect(ws.hyperlinks[3]?.target).toBe('tel:911');
    expect(ws.hyperlinks[4]?.target).toBe('C:\\a.xlsx');

    // Read -> write -> read keeps the ranges internal.
    const wb2 = await readBuffer(await (await readBuffer(await wb.toBuffer())).toBuffer());
    const links = mustGetSheet(wb2, 'Data').hyperlinks;
    expect(links.slice(0, 3).map((hl) => hl.location)).toEqual(ranges);
    expect(links.slice(0, 3).every((hl) => hl.target == null)).toBe(true);
    expect(links.slice(3).map((hl) => hl.target)).toEqual(['tel:911', 'C:\\a.xlsx']);
  });

  it('writes legacy JSON with a URL kept in location as an external relationship', async () => {
    // Legacy JSON keeps an external URL in `location`; the writer heals it into an external relationship.
    const base = new Workbook();
    base.addSheet('Data');
    const data = base.toJSON();
    const sheet = data.sheets[0];
    if (!sheet) throw new Error('expected sheet "Data"');
    sheet.worksheet.hyperlinks = [{ cellRef: 'A1', location: 'https://x.y/' }];
    const buf = await new Workbook(data).toBuffer();
    const parts = readZipText(buf);
    const rels = parts['xl/worksheets/_rels/sheet1.xml.rels'] ?? '';
    expect(rels).toContain('Target="https://x.y/" TargetMode="External"');
    expect(parts['xl/worksheets/sheet1.xml']).toMatch(/<hyperlink ref="A1" r:id="rId[0-9]+"\/>/);
    const link = mustGetSheet(await readBuffer(buf), 'Data').hyperlinks[0];
    expect(link?.target).toBe('https://x.y/');
    expect(link?.location).toBeUndefined();
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
    ws.groupColumns(2, 2, 1);

    const wb2 = await readBuffer(await wb.toBuffer());
    const ws2 = mustGetSheet(wb2, 'Grouped');
    const grouped = ws2.rows.filter((row) => (row.outlineLevel ?? 0) > 0);
    expect(grouped).toHaveLength(3);
    expect(ws2.rows.map((r) => [r.index, r.outlineLevel ?? null])).toEqual([
      [1, null],
      [2, 1],
      [3, 1],
      [4, 1],
      [5, null],
    ]);
    expect(ws2.columns.find((c) => c.min === 2 && c.max === 2)?.outlineLevel).toBe(1);
  });
});
