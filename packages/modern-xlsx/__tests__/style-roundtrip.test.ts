import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('0.1.1 — Style Roundtrip Tests', () => {
  // -----------------------------------------------------------------------
  // 1. Font roundtrip
  // -----------------------------------------------------------------------
  it('font roundtrip — bold, italic, underline, strike, color, size, name', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('FontTest');
    const styleIndex = wb
      .createStyle()
      .font({
        bold: true,
        italic: true,
        underline: true,
        strike: true,
        color: 'FF0000',
        size: 14,
        name: 'Arial',
      })
      .build(wb.styles);

    ws.cell('A1').value = 'Styled text';
    ws.cell('A1').styleIndex = styleIndex;

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    const xf = wb2.styles.cellXfs[styleIndex];
    expect(xf).toBeDefined();

    const font = wb2.styles.fonts[xf?.fontId];
    expect(font).toBeDefined();
    expect(font?.bold).toBe(true);
    expect(font?.italic).toBe(true);
    expect(font?.underline).toBe(true);
    expect(font?.strike).toBe(true);
    expect(font?.color).toBe('FF0000');
    expect(font?.size).toBe(14);
    expect(font?.name).toBe('Arial');
  });

  // -----------------------------------------------------------------------
  // 2. Fill roundtrip (solid)
  // -----------------------------------------------------------------------
  it('fill roundtrip — solid fill with fgColor', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('FillTest');
    const styleIndex = wb
      .createStyle()
      .fill({ pattern: 'solid', fgColor: 'FFFF00' })
      .build(wb.styles);

    ws.cell('A1').value = 'Yellow fill';
    ws.cell('A1').styleIndex = styleIndex;

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    const xf = wb2.styles.cellXfs[styleIndex];
    expect(xf).toBeDefined();

    const fill = wb2.styles.fills[xf?.fillId];
    expect(fill).toBeDefined();
    expect(fill?.patternType).toBe('solid');
    expect(fill?.fgColor).toBe('FFFF00');
  });

  // -----------------------------------------------------------------------
  // 3. Border roundtrip
  // -----------------------------------------------------------------------
  it('border roundtrip — all 4 sides with different styles and colors', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('BorderTest');
    const styleIndex = wb
      .createStyle()
      .border({
        left: { style: 'thin', color: 'FF0000' },
        right: { style: 'medium', color: '00FF00' },
        top: { style: 'thick', color: '0000FF' },
        bottom: { style: 'dashed', color: 'FFFF00' },
      })
      .build(wb.styles);

    ws.cell('A1').value = 'Bordered';
    ws.cell('A1').styleIndex = styleIndex;

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    const xf = wb2.styles.cellXfs[styleIndex];
    expect(xf).toBeDefined();

    const border = wb2.styles.borders[xf?.borderId];
    expect(border).toBeDefined();

    expect(border?.left).toBeDefined();
    expect(border?.left?.style).toBe('thin');
    expect(border?.left?.color).toBe('FF0000');

    expect(border?.right).toBeDefined();
    expect(border?.right?.style).toBe('medium');
    expect(border?.right?.color).toBe('00FF00');

    expect(border?.top).toBeDefined();
    expect(border?.top?.style).toBe('thick');
    expect(border?.top?.color).toBe('0000FF');

    expect(border?.bottom).toBeDefined();
    expect(border?.bottom?.style).toBe('dashed');
    expect(border?.bottom?.color).toBe('FFFF00');
  });

  // -----------------------------------------------------------------------
  // 4. Alignment roundtrip
  // -----------------------------------------------------------------------
  it('alignment roundtrip — horizontal, vertical, wrapText, textRotation, indent, shrinkToFit', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('AlignTest');
    const styleIndex = wb
      .createStyle()
      .alignment({
        horizontal: 'center',
        vertical: 'top',
        wrapText: true,
        textRotation: 45,
        indent: 2,
        shrinkToFit: false,
      })
      .build(wb.styles);

    ws.cell('A1').value = 'Aligned';
    ws.cell('A1').styleIndex = styleIndex;

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    const xf = wb2.styles.cellXfs[styleIndex];
    expect(xf).toBeDefined();
    expect(xf?.alignment).toBeDefined();
    expect(xf?.alignment?.horizontal).toBe('center');
    expect(xf?.alignment?.vertical).toBe('top');
    expect(xf?.alignment?.wrapText).toBe(true);
    expect(xf?.alignment?.textRotation).toBe(45);
    expect(xf?.alignment?.indent).toBe(2);
    // shrinkToFit=false may be omitted (undefined) by the reader; both are falsy
    expect(xf?.alignment?.shrinkToFit).toBeFalsy();
  });

  // -----------------------------------------------------------------------
  // 5. Protection roundtrip
  // -----------------------------------------------------------------------
  it('protection roundtrip — locked and hidden', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('ProtTest');
    const styleIndex = wb.createStyle().protection({ locked: true, hidden: true }).build(wb.styles);

    ws.cell('A1').value = 'Protected';
    ws.cell('A1').styleIndex = styleIndex;

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    const xf = wb2.styles.cellXfs[styleIndex];
    expect(xf).toBeDefined();
    expect(xf?.protection).toBeDefined();
    expect(xf?.protection?.locked).toBe(true);
    expect(xf?.protection?.hidden).toBe(true);
  });

  // -----------------------------------------------------------------------
  // 6. Number format roundtrip
  // -----------------------------------------------------------------------
  it('number format roundtrip — custom format 0.00%', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('NumFmtTest');
    const styleIndex = wb.createStyle().numberFormat('0.00%').build(wb.styles);

    ws.cell('A1').value = 0.1234;
    ws.cell('A1').styleIndex = styleIndex;

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    // Verify the custom numFmt exists
    const customFmts = wb2.styles.numFmts.filter((f) => f.formatCode === '0.00%');
    expect(customFmts).toHaveLength(1);
    expect(customFmts[0]?.id).toBeGreaterThanOrEqual(164);

    // Verify the cellXf references the correct numFmtId
    const xf = wb2.styles.cellXfs[styleIndex];
    expect(xf).toBeDefined();
    expect(xf?.numFmtId).toBe(customFmts[0]?.id);
  });

  // -----------------------------------------------------------------------
  // 7. Gradient fill roundtrip (raw data manipulation)
  // -----------------------------------------------------------------------
  it('gradient fill roundtrip — linear gradient with two stops', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('GradientTest');
    ws.cell('A1').value = 42;

    const data = wb.toJSON();

    // Add a gradient fill
    const gradientFillIndex = data.styles.fills.length;
    data.styles.fills.push({
      patternType: 'none',
      fgColor: null,
      bgColor: null,
      gradientFill: {
        gradientType: 'linear',
        degree: 90,
        stops: [
          { position: 0, color: 'FF0000' },
          { position: 1, color: '0000FF' },
        ],
      },
    });

    // Add cellXf referencing the gradient fill
    const xfIndex = data.styles.cellXfs.length;
    data.styles.cellXfs.push({
      numFmtId: 0,
      fontId: 0,
      fillId: gradientFillIndex,
      borderId: 0,
      applyFill: true,
    });

    // Apply style to the cell
    const targetCell = data.sheets[0]?.worksheet.rows[0]?.cells[0];
    if (targetCell) targetCell.styleIndex = xfIndex;

    const wb1 = new Workbook(data);
    const buffer = await wb1.toBuffer();
    const wb2 = await readBuffer(buffer);

    const xf2 = wb2.styles.cellXfs[xfIndex];
    expect(xf2).toBeDefined();

    const fill2 = wb2.styles.fills[xf2?.fillId];
    expect(fill2).toBeDefined();
    expect(fill2?.gradientFill).toBeDefined();
    expect(fill2?.gradientFill?.gradientType).toBe('linear');
    expect(fill2?.gradientFill?.degree).toBe(90);
    expect(fill2?.gradientFill?.stops).toBeDefined();
    expect(fill2?.gradientFill?.stops).toHaveLength(2);
    expect(fill2?.gradientFill?.stops?.[0]?.position).toBe(0);
    expect(fill2?.gradientFill?.stops?.[0]?.color).toBe('FF0000');
    expect(fill2?.gradientFill?.stops?.[1]?.position).toBe(1);
    expect(fill2?.gradientFill?.stops?.[1]?.color).toBe('0000FF');
  });

  // -----------------------------------------------------------------------
  // 8. Diagonal border roundtrip (raw data manipulation)
  // -----------------------------------------------------------------------
  it('diagonal border roundtrip — diagonal with diagonalUp/diagonalDown', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('DiagTest');
    ws.cell('A1').value = 42;

    const data = wb.toJSON();

    // Add a border with diagonal properties
    const borderIndex = data.styles.borders.length;
    data.styles.borders.push({
      left: null,
      right: null,
      top: null,
      bottom: null,
      diagonal: { style: 'thin', color: 'FF0000' },
      diagonalUp: true,
      diagonalDown: false,
    });

    // Add cellXf referencing the border
    const xfIndex = data.styles.cellXfs.length;
    data.styles.cellXfs.push({
      numFmtId: 0,
      fontId: 0,
      fillId: 0,
      borderId: borderIndex,
      applyBorder: true,
    });

    // Apply style to the cell
    const targetCell = data.sheets[0]?.worksheet.rows[0]?.cells[0];
    if (targetCell) targetCell.styleIndex = xfIndex;

    const wb1 = new Workbook(data);
    const buffer = await wb1.toBuffer();
    const wb2 = await readBuffer(buffer);

    const xf2 = wb2.styles.cellXfs[xfIndex];
    expect(xf2).toBeDefined();

    const border2 = wb2.styles.borders[xf2?.borderId];
    expect(border2).toBeDefined();
    expect(border2?.diagonal).toBeDefined();
    expect(border2?.diagonal?.style).toBe('thin');
    expect(border2?.diagonal?.color).toBe('FF0000');
    expect(border2?.diagonalUp).toBe(true);
    // diagonalDown=false may be omitted (undefined) by the reader; both are falsy
    expect(border2?.diagonalDown).toBeFalsy();
  });

  // -----------------------------------------------------------------------
  // 9. DXF styles roundtrip (raw data manipulation)
  // -----------------------------------------------------------------------
  it('DXF styles roundtrip — font, fill, and border overrides', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('DxfTest');
    ws.cell('A1').value = 42;

    const data = wb.toJSON();

    // Add DXF entries
    if (!data.styles.dxfs) {
      data.styles.dxfs = [];
    }
    data.styles.dxfs.push({
      font: {
        name: 'Courier New',
        size: 12,
        bold: true,
        italic: false,
        underline: false,
        strike: false,
        color: 'CC0000',
      },
      fill: {
        patternType: 'solid',
        fgColor: 'E6E6E6',
        bgColor: null,
      },
      border: {
        left: { style: 'thin', color: '333333' },
        right: { style: 'thin', color: '333333' },
        top: { style: 'thin', color: '333333' },
        bottom: { style: 'thin', color: '333333' },
      },
      numFmt: null,
    });

    const wb1 = new Workbook(data);
    const buffer = await wb1.toBuffer();
    const wb2 = await readBuffer(buffer);

    expect(wb2.styles.dxfs).toBeDefined();
    expect(wb2.styles.dxfs!).toHaveLength(1);

    const dxf = wb2.styles.dxfs?.[0];

    // Verify font overrides
    expect(dxf.font).toBeDefined();
    expect(dxf.font?.name).toBe('Courier New');
    expect(dxf.font?.size).toBe(12);
    expect(dxf.font?.bold).toBe(true);
    expect(dxf.font?.italic).toBe(false);
    expect(dxf.font?.color).toBe('CC0000');

    // Verify fill overrides
    expect(dxf.fill).toBeDefined();
    expect(dxf.fill?.patternType).toBe('solid');
    expect(dxf.fill?.fgColor).toBe('E6E6E6');

    // Verify border overrides
    expect(dxf.border).toBeDefined();
    expect(dxf.border?.left).toBeDefined();
    expect(dxf.border?.left?.style).toBe('thin');
    expect(dxf.border?.left?.color).toBe('333333');
    expect(dxf.border?.right?.style).toBe('thin');
    expect(dxf.border?.top?.style).toBe('thin');
    expect(dxf.border?.bottom?.style).toBe('thin');
  });

  // -----------------------------------------------------------------------
  // 10. Cell styles roundtrip (raw data manipulation)
  // -----------------------------------------------------------------------
  it('cell styles roundtrip — name, xfId, builtinId', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('CellStyleTest');
    ws.cell('A1').value = 42;

    const data = wb.toJSON();

    // Add cell style entries
    if (!data.styles.cellStyles) {
      data.styles.cellStyles = [];
    }
    data.styles.cellStyles.push({
      name: 'Heading',
      xfId: 0,
      builtinId: 1,
    });

    const wb1 = new Workbook(data);
    const buffer = await wb1.toBuffer();
    const wb2 = await readBuffer(buffer);

    expect(wb2.styles.cellStyles).toBeDefined();
    expect(wb2.styles.cellStyles?.length).toBeGreaterThanOrEqual(1);

    const headingStyle = wb2.styles.cellStyles?.find((cs) => cs.name === 'Heading');
    expect(headingStyle).toBeDefined();
    expect(headingStyle?.name).toBe('Heading');
    expect(headingStyle?.xfId).toBe(0);
    expect(headingStyle?.builtinId).toBe(1);
  });

  // -----------------------------------------------------------------------
  // 11. Multiple styles on different cells
  // -----------------------------------------------------------------------
  it('multiple styles on different cells — A1 font+fill, B1 border+alignment, C1 numfmt+protection', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('MultiStyle');

    // Style 1: font + fill for A1
    const style1 = wb
      .createStyle()
      .font({ bold: true, color: 'FF0000', size: 16, name: 'Verdana' })
      .fill({ pattern: 'solid', fgColor: '00FF00' })
      .build(wb.styles);

    // Style 2: border + alignment for B1
    const style2 = wb
      .createStyle()
      .border({
        left: { style: 'medium', color: '0000FF' },
        right: { style: 'medium', color: '0000FF' },
        top: { style: 'medium', color: '0000FF' },
        bottom: { style: 'medium', color: '0000FF' },
      })
      .alignment({ horizontal: 'right', vertical: 'bottom', wrapText: false })
      .build(wb.styles);

    // Style 3: number format + protection for C1
    const style3 = wb
      .createStyle()
      .numberFormat('#,##0.00')
      .protection({ locked: false, hidden: true })
      .build(wb.styles);

    ws.cell('A1').value = 'Bold Red';
    ws.cell('A1').styleIndex = style1;

    ws.cell('B1').value = 'Bordered';
    ws.cell('B1').styleIndex = style2;

    ws.cell('C1').value = 12345.678;
    ws.cell('C1').styleIndex = style3;

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('MultiStyle');
    expect(ws2).toBeDefined();

    // Verify A1 — font + fill
    const a1Style = ws2?.cell('A1').styleIndex;
    expect(a1Style).toBe(style1);
    const xf1 = wb2.styles.cellXfs[a1Style!];
    expect(xf1).toBeDefined();
    const font1 = wb2.styles.fonts[xf1?.fontId];
    expect(font1?.bold).toBe(true);
    expect(font1?.color).toBe('FF0000');
    expect(font1?.size).toBe(16);
    expect(font1?.name).toBe('Verdana');
    const fill1 = wb2.styles.fills[xf1?.fillId];
    expect(fill1?.patternType).toBe('solid');
    expect(fill1?.fgColor).toBe('00FF00');

    // Verify B1 — border + alignment
    const b1Style = ws2?.cell('B1').styleIndex;
    expect(b1Style).toBe(style2);
    const xf2 = wb2.styles.cellXfs[b1Style!];
    expect(xf2).toBeDefined();
    const border2 = wb2.styles.borders[xf2?.borderId];
    expect(border2?.left?.style).toBe('medium');
    expect(border2?.left?.color).toBe('0000FF');
    expect(border2?.right?.style).toBe('medium');
    expect(border2?.top?.style).toBe('medium');
    expect(border2?.bottom?.style).toBe('medium');
    expect(xf2?.alignment).toBeDefined();
    expect(xf2?.alignment?.horizontal).toBe('right');
    expect(xf2?.alignment?.vertical).toBe('bottom');

    // Verify C1 — number format + protection
    const c1Style = ws2?.cell('C1').styleIndex;
    expect(c1Style).toBe(style3);
    const xf3 = wb2.styles.cellXfs[c1Style!];
    expect(xf3).toBeDefined();
    const numFmt = wb2.styles.numFmts.find((f) => f.id === xf3?.numFmtId);
    expect(numFmt).toBeDefined();
    expect(numFmt?.formatCode).toBe('#,##0.00');
    expect(xf3?.protection).toBeDefined();
    expect(xf3?.protection?.locked).toBe(false);
    expect(xf3?.protection?.hidden).toBe(true);
  });

  // -----------------------------------------------------------------------
  // 12. Combined style (all properties)
  // -----------------------------------------------------------------------
  it('combined style roundtrip — font + fill + border + alignment + protection + numberFormat', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('CombinedTest');
    const styleIndex = wb
      .createStyle()
      .font({
        bold: true,
        italic: true,
        underline: true,
        strike: false,
        color: '336699',
        size: 12,
        name: 'Calibri',
      })
      .fill({ pattern: 'solid', fgColor: 'FFFFCC' })
      .border({
        left: { style: 'thin', color: '000000' },
        right: { style: 'thin', color: '000000' },
        top: { style: 'double', color: '000000' },
        bottom: { style: 'double', color: '000000' },
      })
      .alignment({
        horizontal: 'left',
        vertical: 'center',
        wrapText: true,
        textRotation: 90,
        indent: 1,
        shrinkToFit: false,
      })
      .protection({ locked: true, hidden: false })
      .numberFormat('$#,##0.00')
      .build(wb.styles);

    ws.cell('A1').value = 99999.99;
    ws.cell('A1').styleIndex = styleIndex;

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    const xf = wb2.styles.cellXfs[styleIndex];
    expect(xf).toBeDefined();

    // Verify font
    const font = wb2.styles.fonts[xf?.fontId];
    expect(font).toBeDefined();
    expect(font?.bold).toBe(true);
    expect(font?.italic).toBe(true);
    expect(font?.underline).toBe(true);
    expect(font?.strike).toBe(false);
    expect(font?.color).toBe('336699');
    expect(font?.size).toBe(12);
    expect(font?.name).toBe('Calibri');

    // Verify fill
    const fill = wb2.styles.fills[xf?.fillId];
    expect(fill).toBeDefined();
    expect(fill?.patternType).toBe('solid');
    expect(fill?.fgColor).toBe('FFFFCC');

    // Verify border
    const border = wb2.styles.borders[xf?.borderId];
    expect(border).toBeDefined();
    expect(border?.left?.style).toBe('thin');
    expect(border?.left?.color).toBe('000000');
    expect(border?.right?.style).toBe('thin');
    expect(border?.right?.color).toBe('000000');
    expect(border?.top?.style).toBe('double');
    expect(border?.top?.color).toBe('000000');
    expect(border?.bottom?.style).toBe('double');
    expect(border?.bottom?.color).toBe('000000');

    // Verify alignment
    expect(xf?.alignment).toBeDefined();
    expect(xf?.alignment?.horizontal).toBe('left');
    expect(xf?.alignment?.vertical).toBe('center');
    expect(xf?.alignment?.wrapText).toBe(true);
    expect(xf?.alignment?.textRotation).toBe(90);
    expect(xf?.alignment?.indent).toBe(1);
    // shrinkToFit=false may be omitted (undefined) by the reader; both are falsy
    expect(xf?.alignment?.shrinkToFit).toBeFalsy();

    // Verify protection
    expect(xf?.protection).toBeDefined();
    expect(xf?.protection?.locked).toBe(true);
    expect(xf?.protection?.hidden).toBe(false);

    // Verify number format
    const numFmt = wb2.styles.numFmts.find((f) => f.id === xf?.numFmtId);
    expect(numFmt).toBeDefined();
    expect(numFmt?.formatCode).toBe('$#,##0.00');
  });
});
