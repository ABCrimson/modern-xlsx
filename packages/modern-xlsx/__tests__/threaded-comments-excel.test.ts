import { assert, describe, expect, it } from 'vitest';
import type { Worksheet } from '../src/index.js';
import { readBuffer, Workbook } from '../src/index.js';
import { buildZip, readZipText } from './helpers/zip.js';

/** The id form Excel writes and requires: an upper-case GUID in braces. */
const EXCEL_GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/;

/** The parts of a written sheet 1 these tests read. */
const PARTS = [
  'xl/threadedComments/threadedComment1.xml',
  'xl/persons/person.xml',
  'xl/comments1.xml',
  'xl/drawings/vmlDrawing1.vml',
  'xl/worksheets/_rels/sheet1.xml.rels',
] as const;

/** The sheet named `name`, narrowed with `assert.isDefined`. */
function sheet(wb: Workbook, name: string): Worksheet {
  const ws = wb.getSheet(name);
  assert.isDefined(ws);
  return ws;
}

/** The text of `path` in `parts`, failing when it is missing. */
function part(parts: Record<string, string>, path: string): string {
  const text = parts[path];
  assert.isDefined(text, `${path} is missing`);
  return text;
}

/** Every value of attribute `name` in `xml`, in order. */
function attrs(xml: string, name: string): string[] {
  return [...xml.matchAll(new RegExp(` ${name}="([^"]*)"`, 'g'))].map((m) => m[1] ?? '');
}

/** A workbook with a thread on B2 (Alice) and a reply to it (Bob). */
function threadedWorkbook(): { wb: Workbook; rootId: string; replyId: string } {
  const wb = new Workbook();
  const ws = wb.addSheet('Thr');
  ws.cell('A1').value = 'x';
  ws.cell('B2').value = 'y';
  const rootId = ws.addThreadedComment('B2', 'Please check', 'Alice');
  const replyId = ws.replyToComment(rootId, 'Done', 'Bob');
  return { wb, rootId, replyId };
}

describe('threaded comments open in Excel', () => {
  it('addThreadedComment and replyToComment return ids in the braced upper-case form Excel needs', () => {
    const { wb, rootId, replyId } = threadedWorkbook();
    expect(rootId).toMatch(EXCEL_GUID);
    expect(replyId).toMatch(EXCEL_GUID);
    const ws = sheet(wb, 'Thr');
    expect(ws.threadedComments.map((c) => c.parentId)).toEqual([undefined, rootId]);
    for (const c of ws.threadedComments) expect(c.personId).toMatch(EXCEL_GUID);
    for (const p of wb.toJSON().persons ?? []) expect(p.id).toMatch(EXCEL_GUID);
    // replyToComment finds the parent by the id it returned.
    expect(ws.replyToComment(rootId, 'Again', 'Alice')).toMatch(EXCEL_GUID);
  });

  it('a written thread has Excel-form ids, its person, a tc= placeholder note and one note shape', async () => {
    const { wb, rootId, replyId } = threadedWorkbook();
    const parts = readZipText(await wb.toBuffer());

    const threads = part(parts, 'xl/threadedComments/threadedComment1.xml');
    expect(attrs(threads, 'id')).toEqual([rootId, replyId]);
    for (const id of [...attrs(threads, 'id'), ...attrs(threads, 'personId')]) {
      expect(id).toMatch(EXCEL_GUID);
    }
    expect(attrs(threads, 'parentId')).toEqual([rootId]);

    const persons = part(parts, 'xl/persons/person.xml');
    expect(new Set(attrs(persons, 'id'))).toEqual(new Set(attrs(threads, 'personId')));

    const comments = part(parts, 'xl/comments1.xml');
    expect(comments).toContain(`<author>tc=${rootId}</author>`);
    expect(comments).toContain('<comment ref="B2"');
    expect(comments).toContain('[Threaded comment]');
    expect(comments).toMatch(/Comment:\n {4}Please check\nReply:\n {4}Done/);

    const vml = part(parts, 'xl/drawings/vmlDrawing1.vml');
    expect(vml.match(/<x:ClientData ObjectType="Note">/g)).toHaveLength(1);
    expect(part(parts, 'xl/worksheets/sheet1.xml')).toMatch(/<legacyDrawing r:id="rId\d+"\/>/);

    // The model is not changed by the write: the placeholder is only in the file.
    expect(sheet(wb, 'Thr').comments).toEqual([]);
  });

  it('a reply to a reply joins the thread root, and the placeholder lists it', async () => {
    const { wb, rootId, replyId } = threadedWorkbook();
    const ws = sheet(wb, 'Thr');
    const nestedId = ws.replyToComment(replyId, 'Checked again', 'Alice');
    expect(ws.threadedComments.find((c) => c.id === nestedId)?.parentId).toBe(rootId);

    const parts = readZipText(await wb.toBuffer());
    expect(attrs(part(parts, 'xl/threadedComments/threadedComment1.xml'), 'parentId')).toEqual([
      rootId,
      rootId,
    ]);
    expect(part(parts, 'xl/comments1.xml')).toMatch(
      /Comment:\n {4}Please check\nReply:\n {4}Done\nReply:\n {4}Checked again/,
    );
  });

  it('addThreadedComment stores the cell reference without $ and in upper case', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Refs');
    ws.addThreadedComment('$D$4', 'Absolute', 'Alice');
    ws.addThreadedComment('b2', 'Lower case', 'Alice');
    expect(ws.threadedComments.map((c) => c.refCell)).toEqual(['D4', 'B2']);
  });

  it('a read of the written file returns the thread, and a second write is stable', async () => {
    const { wb } = threadedWorkbook();
    const first = await wb.toBuffer();
    const read = await readBuffer(first);
    const ws = sheet(read, 'Thr');
    // Each timestamp is written in Excel's dT form: UTC, hundredths, no `Z`.
    expect(ws.threadedComments).toEqual(
      sheet(wb, 'Thr').threadedComments.map((c) => ({
        ...c,
        timestamp: c.timestamp.replace(/^(.*\.\d{2})\d*Z$/, '$1'),
      })),
    );
    expect(read.toJSON().persons).toEqual(wb.toJSON().persons);
    expect(ws.comments.map((c) => c.author)).toEqual([`tc=${ws.threadedComments[0]?.id}`]);

    const [a, b] = [readZipText(first), readZipText(await read.toBuffer())];
    for (const path of PARTS) expect(part(b, path), path).toBe(part(a, path));
  });

  it('validate() reports a raw model with bare ids as errors', () => {
    const { wb } = threadedWorkbook();
    expect(wb.validate().isValid).toBe(true);
    const data = wb.toJSON();
    const [root] = sheet(wb, 'Thr').threadedComments;
    assert.isDefined(root);
    const bare = root.id.slice(1, -1).toLowerCase();
    // Raw-JSON callers can put any string here.
    Object.assign(root, { id: bare });
    const report = wb.validate();
    expect(report.isValid).toBe(false);
    expect(
      report.issues.filter((i) => i.severity === 'error').map((i) => [i.category, i.location]),
    ).toEqual([['structure', 'Thr!threadedComments[0]']]);
    expect(data.sheets[0]?.worksheet.threadedComments?.[0]?.id).toBe(bare);
  });
});

/**
 * The threaded comments part Excel 16.0 (build 20527) saved on 2026-11-12 for
 * threads on A1 (two replies), B2 (one reply) and C3, added through COM with
 * `Range.AddCommentThreaded` and `CommentThreaded.AddReply`; verbatim.
 */
const EXCEL_THREADS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><threadedComment ref="A1" dT="2026-11-12T20:23:00.33" personId="{2636B023-1FE4-4E47-B301-7CE1AA501A89}" id="{9C5C4D94-0E96-47F6-9B5B-7B357C40E46C}"><text>Root on A1</text></threadedComment><threadedComment ref="A1" dT="2026-11-12T20:23:00.34" personId="{2636B023-1FE4-4E47-B301-7CE1AA501A89}" id="{5289CABE-EBD7-4600-AFF7-B467D7880758}" parentId="{9C5C4D94-0E96-47F6-9B5B-7B357C40E46C}"><text>Reply one on A1</text></threadedComment><threadedComment ref="A1" dT="2026-11-12T20:23:00.34" personId="{2636B023-1FE4-4E47-B301-7CE1AA501A89}" id="{5289CABE-EBD7-4601-AFF7-B467D7880758}" parentId="{9C5C4D94-0E96-47F6-9B5B-7B357C40E46C}"><text>Reply two on A1</text></threadedComment><threadedComment ref="B2" dT="2026-11-12T20:23:00.34" personId="{2636B023-1FE4-4E47-B301-7CE1AA501A89}" id="{DAC092D4-2AE3-4CB8-ADB3-03B050961FF4}"><text>Root on B2</text></threadedComment><threadedComment ref="B2" dT="2026-11-12T20:23:00.34" personId="{2636B023-1FE4-4E47-B301-7CE1AA501A89}" id="{59F25975-E87B-4FBB-91D1-8FEBDDD9BD86}" parentId="{DAC092D4-2AE3-4CB8-ADB3-03B050961FF4}"><text>Reply on B2</text></threadedComment><threadedComment ref="C3" dT="2026-11-12T20:23:00.34" personId="{2636B023-1FE4-4E47-B301-7CE1AA501A89}" id="{D865BFE2-48F4-4C06-8925-ADE288C14C93}"><text>Root on C3</text></threadedComment></ThreadedComments>';

/** The person part of that file, with the display name and user id replaced. */
const EXCEL_PERSONS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><person displayName="Excel User" id="{2636B023-1FE4-4E47-B301-7CE1AA501A89}" userId="0000000000000000" providerId="Windows Live"/></personList>';

/** The thread root ids of {@link EXCEL_THREADS}, in cell order. */
const EXCEL_ROOT_IDS = [
  '{9C5C4D94-0E96-47F6-9B5B-7B357C40E46C}',
  '{DAC092D4-2AE3-4CB8-ADB3-03B050961FF4}',
  '{D865BFE2-48F4-4C06-8925-ADE288C14C93}',
];

/** A workbook with one thread per cell on A1 (two replies), B2 (one reply) and C3. */
function threePerCell(): Workbook {
  const wb = new Workbook();
  const ws = wb.addSheet('Sheet1');
  ws.cell('A1').value = 'x';
  const a1 = ws.addThreadedComment('A1', 'Root on A1', 'Alice');
  ws.replyToComment(a1, 'Reply one on A1', 'Bob');
  ws.replyToComment(a1, 'Reply two on A1', 'Alice');
  const b2 = ws.addThreadedComment('B2', 'Root on B2', 'Bob');
  ws.replyToComment(b2, 'Reply on B2', 'Alice');
  ws.addThreadedComment('C3', 'Root on C3', 'Alice');
  return wb;
}

/**
 * A file of the {@link threePerCell} shape with Excel's threads and person put
 * in, and each placeholder's tc= author naming Excel's root id.
 */
async function excelThreadsFile(): Promise<Uint8Array> {
  const wb = threePerCell();
  const parts = readZipText(await wb.toBuffer());
  const ours = sheet(wb, 'Sheet1').threadedComments.filter((c) => c.parentId === undefined);
  let comments = part(parts, 'xl/comments1.xml');
  ours.forEach((root, i) => {
    comments = comments.replace(`tc=${root.id}`, `tc=${EXCEL_ROOT_IDS[i]}`);
  });
  const swapped: Record<string, string> = {
    ...parts,
    'xl/threadedComments/threadedComment1.xml': EXCEL_THREADS,
    'xl/persons/person.xml': EXCEL_PERSONS,
    'xl/comments1.xml': comments,
  };
  const encoder = new TextEncoder();
  return buildZip(
    Object.entries(swapped).map(([name, text]) => [name, encoder.encode(text)] as const),
  );
}

describe('one thread per cell', () => {
  it('threads Excel made on several cells read back and validate clean', async () => {
    const read = await readBuffer(await excelThreadsFile());
    const threads = sheet(read, 'Sheet1').threadedComments;
    expect(threads.map((c) => [c.refCell, c.parentId === undefined])).toEqual([
      ['A1', true],
      ['A1', false],
      ['A1', false],
      ['B2', true],
      ['B2', false],
      ['C3', true],
    ]);
    const report = read.validate();
    expect(report.issues.filter((i) => i.severity === 'error')).toEqual([]);
    expect(report.isValid).toBe(true);
  });

  it("Excel's threads read and written back keep their order, ids and dT values", async () => {
    const read = await readBuffer(await excelThreadsFile());
    // A read gives dT back as the file holds it.
    expect(sheet(read, 'Sheet1').threadedComments.map((c) => c.timestamp)).toEqual(
      attrs(EXCEL_THREADS, 'dT'),
    );
    const written = part(
      readZipText(await read.toBuffer()),
      'xl/threadedComments/threadedComment1.xml',
    );
    for (const name of ['dT', 'id', 'parentId', 'ref']) {
      expect(attrs(written, name), name).toEqual(attrs(EXCEL_THREADS, name));
    }
  });

  it('three cells with replies write three thread roots and three placeholders', async () => {
    const wb = threePerCell();
    expect(wb.validate().isValid).toBe(true);
    const parts = readZipText(await wb.toBuffer());
    const threads = part(parts, 'xl/threadedComments/threadedComment1.xml');
    expect(attrs(threads, 'ref')).toEqual(['A1', 'A1', 'A1', 'B2', 'B2', 'C3']);
    expect(attrs(threads, 'parentId')).toHaveLength(3);
    expect(attrs(part(parts, 'xl/comments1.xml'), 'ref')).toEqual(['A1', 'B2', 'C3']);
  });
});
