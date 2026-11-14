use super::*;
use pretty_assertions::assert_eq;

const MINIMAL_STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <fonts count="2">
<font><sz val="11"/><name val="Aptos"/></font>
<font><b/><sz val="11"/><name val="Aptos"/></font>
  </fonts>
  <fills count="2">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
<border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"#;

#[test]
fn test_parse_styles() {
    let styles = Styles::parse(MINIMAL_STYLES.as_bytes()).unwrap();

    // numFmts
    assert_eq!(styles.num_fmts.len(), 1);
    assert_eq!(styles.num_fmts[0].id, 164);
    assert_eq!(styles.num_fmts[0].format_code, "yyyy-mm-dd");

    // fonts
    assert_eq!(styles.fonts.len(), 2);
    assert!(!styles.fonts[0].bold);
    assert_eq!(styles.fonts[0].name.as_deref(), Some("Aptos"));
    assert_eq!(styles.fonts[0].size, Some(11.0));
    assert!(styles.fonts[1].bold);
    assert_eq!(styles.fonts[1].name.as_deref(), Some("Aptos"));
    assert_eq!(styles.fonts[1].size, Some(11.0));

    // fills
    assert_eq!(styles.fills.len(), 2);
    assert_eq!(styles.fills[0].pattern_type, "none");
    assert_eq!(styles.fills[1].pattern_type, "gray125");

    // borders
    assert_eq!(styles.borders.len(), 1);
    assert!(styles.borders[0].left.is_none());
    assert!(styles.borders[0].right.is_none());
    assert!(styles.borders[0].top.is_none());
    assert!(styles.borders[0].bottom.is_none());

    // cellXfs
    assert_eq!(styles.cell_xfs.len(), 3);
}

#[test]
fn test_style_get_num_fmt_id() {
    let styles = Styles::parse(MINIMAL_STYLES.as_bytes()).unwrap();
    assert_eq!(styles.cell_xfs[0].num_fmt_id, 0);
    assert_eq!(styles.cell_xfs[2].num_fmt_id, 164);
}

#[test]
fn test_styles_write_roundtrip() {
    let styles1 = Styles::parse(MINIMAL_STYLES.as_bytes()).unwrap();
    let xml = styles1.to_xml().unwrap();
    let styles2 = Styles::parse(&xml).unwrap();

    assert_eq!(styles2.num_fmts.len(), styles1.num_fmts.len());
    assert_eq!(styles2.fonts.len(), styles1.fonts.len());
    assert_eq!(styles2.fills.len(), styles1.fills.len());
    assert_eq!(styles2.borders.len(), styles1.borders.len());
    assert_eq!(styles2.cell_xfs.len(), styles1.cell_xfs.len());

    // Spot-check values survived the roundtrip.
    assert_eq!(styles2.num_fmts[0].id, 164);
    assert_eq!(styles2.num_fmts[0].format_code, "yyyy-mm-dd");
    assert!(styles2.fonts[1].bold);
    assert_eq!(styles2.fills[1].pattern_type, "gray125");
    assert_eq!(styles2.cell_xfs[2].num_fmt_id, 164);
}

#[test]
fn test_default_styles() {
    let styles = Styles::default_styles();

    // Verify shape.
    assert_eq!(styles.num_fmts.len(), 0);
    assert_eq!(styles.fonts.len(), 1);
    assert_eq!(styles.fills.len(), 2);
    assert_eq!(styles.borders.len(), 1);
    assert_eq!(styles.cell_xfs.len(), 1);

    // Verify default font.
    assert_eq!(styles.fonts[0].name.as_deref(), Some("Calibri"));
    assert_eq!(styles.fonts[0].size, Some(11.0));

    // Verify fills.
    assert_eq!(styles.fills[0].pattern_type, "none");
    assert_eq!(styles.fills[1].pattern_type, "gray125");

    // Roundtrip through XML.
    let xml = styles.to_xml().unwrap();
    let styles2 = Styles::parse(&xml).unwrap();

    assert_eq!(styles2.num_fmts.len(), 0);
    assert_eq!(styles2.fonts.len(), 1);
    assert_eq!(styles2.fills.len(), 2);
    assert_eq!(styles2.borders.len(), 1);
    assert_eq!(styles2.cell_xfs.len(), 1);
    assert_eq!(styles2.fonts[0].name.as_deref(), Some("Calibri"));
    assert_eq!(styles2.fonts[0].size, Some(11.0));
}

#[test]
fn test_parse_apply_flags() {
    let styles = Styles::parse(MINIMAL_STYLES.as_bytes()).unwrap();

    // xf[0] has no apply flags
    assert!(!styles.cell_xfs[0].apply_font);
    assert!(!styles.cell_xfs[0].apply_fill);
    assert!(!styles.cell_xfs[0].apply_number_format);

    // xf[1] has applyFont="1"
    assert!(styles.cell_xfs[1].apply_font);
    assert!(!styles.cell_xfs[1].apply_fill);

    // xf[2] has applyNumberFormat="1"
    assert!(styles.cell_xfs[2].apply_number_format);
    assert!(!styles.cell_xfs[2].apply_font);
}

#[test]
fn test_parse_alignment() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1">
  <alignment horizontal="center" vertical="bottom" wrapText="1" textRotation="90" indent="2" shrinkToFit="1"/>
</xf>
  </cellXfs>
</styleSheet>"#;

    let styles = Styles::parse(xml.as_bytes()).unwrap();
    assert_eq!(styles.cell_xfs.len(), 1);

    let xf = &styles.cell_xfs[0];
    assert!(xf.apply_alignment);
    let align = xf.alignment.as_ref().unwrap();
    assert_eq!(align.horizontal.as_deref(), Some("center"));
    assert_eq!(align.vertical.as_deref(), Some("bottom"));
    assert!(align.wrap_text);
    assert_eq!(align.text_rotation, Some(90));
    assert_eq!(align.indent, Some(2));
    assert!(align.shrink_to_fit);

    // Roundtrip
    let xml_out = styles.to_xml().unwrap();
    let styles2 = Styles::parse(&xml_out).unwrap();
    let xf2 = &styles2.cell_xfs[0];
    assert!(xf2.apply_alignment);
    let align2 = xf2.alignment.as_ref().unwrap();
    assert_eq!(align2.horizontal.as_deref(), Some("center"));
    assert_eq!(align2.vertical.as_deref(), Some("bottom"));
    assert!(align2.wrap_text);
    assert_eq!(align2.text_rotation, Some(90));
    assert_eq!(align2.indent, Some(2));
    assert!(align2.shrink_to_fit);
}

#[test]
fn test_parse_protection() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1">
  <protection locked="0" hidden="1"/>
</xf>
  </cellXfs>
</styleSheet>"#;

    let styles = Styles::parse(xml.as_bytes()).unwrap();
    let xf = &styles.cell_xfs[0];
    assert!(xf.apply_protection);
    let prot = xf.protection.as_ref().unwrap();
    assert!(!prot.locked);
    assert!(prot.hidden);

    // Roundtrip
    let xml_out = styles.to_xml().unwrap();
    let styles2 = Styles::parse(&xml_out).unwrap();
    let prot2 = styles2.cell_xfs[0].protection.as_ref().unwrap();
    assert!(!prot2.locked);
    assert!(prot2.hidden);
}

#[test]
fn test_parse_diagonal_border() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1">
<border diagonalUp="1" diagonalDown="1">
  <left/>
  <right/>
  <top/>
  <bottom/>
  <diagonal style="thin">
    <color rgb="FFFF0000"/>
  </diagonal>
</border>
  </borders>
  <cellXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>"#;

    let styles = Styles::parse(xml.as_bytes()).unwrap();
    let border = &styles.borders[0];
    assert!(border.diagonal_up);
    assert!(border.diagonal_down);
    let diag = border.diagonal.as_ref().unwrap();
    assert_eq!(diag.style, "thin");
    assert_eq!(diag.color.as_deref(), Some("FFFF0000"));

    // Roundtrip
    let xml_out = styles.to_xml().unwrap();
    let styles2 = Styles::parse(&xml_out).unwrap();
    let border2 = &styles2.borders[0];
    assert!(border2.diagonal_up);
    assert!(border2.diagonal_down);
    let diag2 = border2.diagonal.as_ref().unwrap();
    assert_eq!(diag2.style, "thin");
    assert_eq!(diag2.color.as_deref(), Some("FFFF0000"));
}

#[test]
fn test_parse_gradient_fill() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="2">
<fill><patternFill patternType="none"/></fill>
<fill>
  <gradientFill type="linear" degree="90">
    <stop position="0">
      <color rgb="FF000000"/>
    </stop>
    <stop position="1">
      <color rgb="FFFFFFFF"/>
    </stop>
  </gradientFill>
</fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>"#;

    let styles = Styles::parse(xml.as_bytes()).unwrap();
    assert_eq!(styles.fills.len(), 2);

    let grad_fill = &styles.fills[1];
    let gradient = grad_fill.gradient_fill.as_ref().unwrap();
    assert_eq!(gradient.gradient_type.as_deref(), Some("linear"));
    assert_eq!(gradient.degree, Some(90.0));
    assert_eq!(gradient.stops.len(), 2);
    assert_eq!(gradient.stops[0].position, 0.0);
    assert_eq!(gradient.stops[0].color, "FF000000");
    assert_eq!(gradient.stops[1].position, 1.0);
    assert_eq!(gradient.stops[1].color, "FFFFFFFF");

    // Roundtrip
    let xml_out = styles.to_xml().unwrap();
    let styles2 = Styles::parse(&xml_out).unwrap();
    let gradient2 = styles2.fills[1].gradient_fill.as_ref().unwrap();
    assert_eq!(gradient2.gradient_type.as_deref(), Some("linear"));
    assert_eq!(gradient2.degree, Some(90.0));
    assert_eq!(gradient2.stops.len(), 2);
    assert_eq!(gradient2.stops[0].color, "FF000000");
    assert_eq!(gradient2.stops[1].color, "FFFFFFFF");
}

#[test]
fn test_parse_font_extended_properties() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
<font>
  <b/>
  <sz val="11"/>
  <name val="Aptos"/>
  <color rgb="FF000000"/>
  <vertAlign val="superscript"/>
  <family val="2"/>
  <charset val="1"/>
  <scheme val="minor"/>
  <condense/>
  <extend/>
</font>
  </fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>"#;

    let styles = Styles::parse(xml.as_bytes()).unwrap();
    let font = &styles.fonts[0];
    assert!(font.bold);
    assert_eq!(font.vert_align.as_deref(), Some("superscript"));
    assert_eq!(font.family, Some(2));
    assert_eq!(font.charset, Some(1));
    assert_eq!(font.scheme.as_deref(), Some("minor"));
    assert!(font.condense);
    assert!(font.extend);

    // Roundtrip
    let xml_out = styles.to_xml().unwrap();
    let styles2 = Styles::parse(&xml_out).unwrap();
    let font2 = &styles2.fonts[0];
    assert_eq!(font2.vert_align.as_deref(), Some("superscript"));
    assert_eq!(font2.family, Some(2));
    assert_eq!(font2.charset, Some(1));
    assert_eq!(font2.scheme.as_deref(), Some("minor"));
    assert!(font2.condense);
    assert!(font2.extend);
}

#[test]
fn test_parse_dxf_styles() {
    let xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
  <dxfs count="2">
<dxf>
  <font><b/><color rgb="FFFF0000"/></font>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
</dxf>
<dxf>
  <numFmt numFmtId="164" formatCode="#,##0.00"/>
  <border>
    <left style="thin"><color rgb="FF000000"/></left>
    <right/><top/><bottom/><diagonal/>
  </border>
</dxf>
  </dxfs>
</styleSheet>"##;

    let styles = Styles::parse(xml.as_bytes()).unwrap();
    assert_eq!(styles.dxfs.len(), 2);

    // First DXF: font + fill
    let dxf0 = &styles.dxfs[0];
    let font = dxf0.font.as_ref().unwrap();
    assert!(font.bold);
    assert_eq!(font.color.as_deref(), Some("FFFF0000"));
    let fill = dxf0.fill.as_ref().unwrap();
    assert_eq!(fill.pattern_type, "solid");
    assert_eq!(fill.fg_color.as_deref(), Some("FFFFFF00"));

    // Second DXF: numFmt + border
    let dxf1 = &styles.dxfs[1];
    let nf = dxf1.num_fmt.as_ref().unwrap();
    assert_eq!(nf.id, 164);
    assert_eq!(nf.format_code, "#,##0.00");
    let border = dxf1.border.as_ref().unwrap();
    let left = border.left.as_ref().unwrap();
    assert_eq!(left.style, "thin");
    assert_eq!(left.color.as_deref(), Some("FF000000"));

    // Roundtrip
    let xml_out = styles.to_xml().unwrap();
    let styles2 = Styles::parse(&xml_out).unwrap();
    assert_eq!(styles2.dxfs.len(), 2);
    let dxf0_rt = &styles2.dxfs[0];
    assert!(dxf0_rt.font.as_ref().unwrap().bold);
    assert_eq!(
        dxf0_rt.fill.as_ref().unwrap().fg_color.as_deref(),
        Some("FFFFFF00")
    );
    let dxf1_rt = &styles2.dxfs[1];
    assert_eq!(dxf1_rt.num_fmt.as_ref().unwrap().format_code, "#,##0.00");
}

#[test]
fn test_parse_cell_styles() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
  <cellStyles count="2">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
<cellStyle name="Custom" xfId="1"/>
  </cellStyles>
</styleSheet>"#;

    let styles = Styles::parse(xml.as_bytes()).unwrap();
    assert_eq!(styles.cell_styles.len(), 2);
    assert_eq!(styles.cell_styles[0].name, "Normal");
    assert_eq!(styles.cell_styles[0].xf_id, 0);
    assert_eq!(styles.cell_styles[0].builtin_id, Some(0));
    assert_eq!(styles.cell_styles[1].name, "Custom");
    assert_eq!(styles.cell_styles[1].xf_id, 1);
    assert_eq!(styles.cell_styles[1].builtin_id, None);

    // Roundtrip
    let xml_out = styles.to_xml().unwrap();
    let styles2 = Styles::parse(&xml_out).unwrap();
    assert_eq!(styles2.cell_styles.len(), 2);
    assert_eq!(styles2.cell_styles[0].name, "Normal");
    assert_eq!(styles2.cell_styles[0].builtin_id, Some(0));
    assert_eq!(styles2.cell_styles[1].name, "Custom");
    assert_eq!(styles2.cell_styles[1].builtin_id, None);
}

#[test]
fn test_apply_flags_roundtrip() {
    let styles1 = Styles::parse(MINIMAL_STYLES.as_bytes()).unwrap();

    // xf[1] has applyFont="1"
    assert!(styles1.cell_xfs[1].apply_font);
    // xf[2] has applyNumberFormat="1"
    assert!(styles1.cell_xfs[2].apply_number_format);

    let xml = styles1.to_xml().unwrap();
    let styles2 = Styles::parse(&xml).unwrap();

    assert!(styles2.cell_xfs[1].apply_font);
    assert!(!styles2.cell_xfs[1].apply_fill);
    assert!(styles2.cell_xfs[2].apply_number_format);
    assert!(!styles2.cell_xfs[2].apply_font);
}

#[test]
fn test_protection_default_values() {
    let prot = Protection::default();
    assert!(prot.locked);
    assert!(!prot.hidden);
}

#[test]
fn test_dxf_with_gradient_fill_roundtrip() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
  <dxfs count="1">
<dxf>
  <fill>
    <gradientFill type="linear" degree="45">
      <stop position="0">
        <color rgb="FFFF0000"/>
      </stop>
      <stop position="1">
        <color rgb="FF00FF00"/>
      </stop>
    </gradientFill>
  </fill>
</dxf>
  </dxfs>
</styleSheet>"#;

    let styles = Styles::parse(xml.as_bytes()).unwrap();
    assert_eq!(styles.dxfs.len(), 1);
    let fill = styles.dxfs[0].fill.as_ref().unwrap();
    let gradient = fill.gradient_fill.as_ref().unwrap();
    assert_eq!(gradient.gradient_type.as_deref(), Some("linear"));
    assert_eq!(gradient.degree, Some(45.0));
    assert_eq!(gradient.stops.len(), 2);
    assert_eq!(gradient.stops[0].color, "FFFF0000");
    assert_eq!(gradient.stops[1].color, "FF00FF00");

    // Roundtrip
    let xml_out = styles.to_xml().unwrap();
    let styles2 = Styles::parse(&xml_out).unwrap();
    let fill2 = styles2.dxfs[0].fill.as_ref().unwrap();
    let gradient2 = fill2.gradient_fill.as_ref().unwrap();
    assert_eq!(gradient2.stops.len(), 2);
    assert_eq!(gradient2.stops[0].color, "FFFF0000");
}

/// Wrap `body` in a `<styleSheet>` root with the SpreadsheetML namespace.
fn sheet(body: &str) -> Vec<u8> {
    format!(r#"<styleSheet xmlns="{SPREADSHEET_NS}">{body}</styleSheet>"#).into_bytes()
}

/// `to_xml` of `styles` as a string.
fn xml_of(styles: &Styles) -> String {
    String::from_utf8(styles.to_xml().unwrap()).unwrap()
}

#[test]
fn boolean_properties_read_val() {
    let styles = Styles::parse(&sheet(
        r#"<fonts count="2"><font><b val="0"/><i val="false"/><strike val="0"/><condense val="0"/><extend val="false"/><u val="double"/></font><font><u val="none"/><b val="1"/><i/></font></fonts>"#,
    ))
    .unwrap();
    let f = &styles.fonts[0];
    assert!(!f.bold && !f.italic && !f.strike && !f.condense && !f.extend);
    assert!(f.underline);
    assert_eq!(f.underline_style.as_deref(), Some("double"));
    assert!(!styles.fonts[1].underline);
    assert_eq!(styles.fonts[1].underline_style, None);
    assert!(styles.fonts[1].bold && styles.fonts[1].italic);
    let out = xml_of(&styles);
    // Only the second font is bold; the first writes no <b/>.
    assert_eq!(out.matches("<b/>").count(), 1, "{out}");
    assert!(out.contains(r#"<font><u val="double"/>"#), "{out}");
    assert!(!out.contains("<strike/>"), "{out}");
    let again = Styles::parse(out.as_bytes()).unwrap();
    assert!(!again.fonts[0].bold && again.fonts[0].underline);
    assert_eq!(again.fonts[0].underline_style.as_deref(), Some("double"));
    assert!(!again.fonts[1].underline);
}

#[test]
fn underline_single_and_bare_write_plain_u() {
    let styles = Styles::parse(&sheet(
        r#"<fonts count="2"><font><u/></font><font><u val="single"/></font></fonts>"#,
    ))
    .unwrap();
    for f in &styles.fonts {
        assert!(f.underline);
        assert_eq!(f.underline_style, None);
    }
    assert_eq!(xml_of(&styles).matches("<u/>").count(), 2);
}

#[test]
fn theme_font_colour_round_trips() {
    let styles = Styles::parse(&sheet(
        r#"<fonts count="1"><font><color theme="1" tint="-0.25"/></font></fonts>"#,
    ))
    .unwrap();
    let again = Styles::parse(&styles.to_xml().unwrap()).unwrap();
    let font = &again.fonts[0];
    assert_eq!(
        font.color_spec,
        Some(Color {
            theme: Some(1),
            tint: Some(-0.25),
            ..Color::default()
        })
    );
    assert_eq!(font.color, None);
    assert!(xml_of(&again).contains(r#"<color theme="1" tint="-0.25"/>"#));
}

#[test]
fn fill_and_border_colour_specs_round_trip() {
    let styles = Styles::parse(&sheet(concat!(
        r#"<fills count="1"><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.4"/><bgColor indexed="64"/></patternFill></fill></fills>"#,
        r#"<borders count="1"><border><left style="thin"><color auto="1"/></left><right/><top/><bottom/><diagonal/></border></borders>"#,
    )))
    .unwrap();
    let out = xml_of(&styles);
    assert!(
        out.contains(r#"<fgColor theme="4" tint="0.4"/><bgColor indexed="64"/>"#),
        "{out}"
    );
    assert!(
        out.contains(r#"<left style="thin"><color auto="1"/></left>"#),
        "{out}"
    );
    assert!(!out.contains(r#"rgb="""#), "{out}");
    let again = Styles::parse(out.as_bytes()).unwrap();
    assert_eq!(again.fills[0].fg_color_spec, styles.fills[0].fg_color_spec);
    assert_eq!(
        again.fills[0].bg_color_spec.as_ref().unwrap().indexed,
        Some(64)
    );
    let left = again.borders[0].left.as_ref().unwrap();
    assert_eq!(left.color_spec.as_ref().unwrap().auto, Some(true));
    assert_eq!(left.color, None);
}

#[test]
fn colour_spec_yields_to_a_changed_legacy_string() {
    let mut styles = Styles::parse(&sheet(
        r#"<fonts count="1"><font><color rgb="FF000000" theme="1"/></font></fonts>"#,
    ))
    .unwrap();
    assert!(xml_of(&styles).contains(r#"<color rgb="FF000000" theme="1"/>"#));
    // An API caller edits the legacy string: the stale spec is not written.
    styles.fonts[0].color = Some("FFFF0000".into());
    let out = xml_of(&styles);
    assert!(out.contains(r#"<color rgb="FFFF0000"/>"#), "{out}");
    // An empty legacy string counts as absent and matches a spec without rgb.
    styles.fonts[0].color = Some(String::new());
    styles.fonts[0].color_spec = Some(Color {
        theme: Some(3),
        ..Color::default()
    });
    assert!(xml_of(&styles).contains(r#"<color theme="3"/>"#));
}

#[test]
fn alignment_new_attributes_round_trip() {
    let styles = Styles::parse(&sheet(concat!(
        r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0">"#,
        r#"<alignment horizontal="distributed" readingOrder="2" justifyLastLine="1" relativeIndent="1" indent="1"/>"#,
        r#"</xf></cellXfs>"#,
    )))
    .unwrap();
    let again = Styles::parse(&styles.to_xml().unwrap()).unwrap();
    let a = again.cell_xfs[0].alignment.as_ref().unwrap();
    assert_eq!(a.reading_order, Some(2));
    assert!(a.justify_last_line);
    assert_eq!(a.relative_indent, Some(1));
    assert_eq!(a.indent, Some(1));
}

#[test]
fn dxf_alignment_and_protection_keep_schema_order() {
    let styles = Styles::parse(&sheet(concat!(
        r#"<dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill>"#,
        r#"<alignment horizontal="center" readingOrder="2"/>"#,
        r#"<border><left style="thin"/></border>"#,
        r#"<protection locked="0" hidden="1"/></dxf></dxfs>"#,
    )))
    .unwrap();
    let dxf = &styles.dxfs[0];
    assert_eq!(
        dxf.alignment.as_ref().unwrap().horizontal.as_deref(),
        Some("center")
    );
    let out = xml_of(&styles);
    let alignment = out.find("</fill><alignment").expect("alignment after fill");
    let border = out.find("<border>").expect("border");
    assert!(alignment < border, "{out}");
    assert!(
        out.contains(r#"</border><protection locked="0" hidden="1"/></dxf>"#),
        "{out}"
    );
    let again = Styles::parse(out.as_bytes()).unwrap();
    let dxf = &again.dxfs[0];
    assert_eq!(dxf.alignment.as_ref().unwrap().reading_order, Some(2));
    let p = dxf.protection.as_ref().unwrap();
    assert!(!p.locked && p.hidden);
    assert!(dxf.border.as_ref().unwrap().left.is_some());
}

#[test]
fn cell_style_xfs_are_always_written() {
    let mut styles = Styles::default_styles();
    styles.cell_styles.push(CellStyle {
        name: "Normal".into(),
        xf_id: 0,
        builtin_id: Some(0),
    });
    let out = xml_of(&styles);
    assert!(
        out.contains(r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs"#),
        "{out}"
    );
    assert!(out.contains(r#"<cellStyles count="1">"#), "{out}");
}

#[test]
fn cell_style_xfs_and_xf_id_round_trip() {
    let styles = Styles::parse(&sheet(concat!(
        r#"<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>"#,
        r#"<xf numFmtId="9" fontId="0" fillId="0" borderId="0"><alignment vertical="top"/></xf></cellStyleXfs>"#,
        r#"<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        r#"<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="1" applyNumberFormat="1"/></cellXfs>"#,
    )))
    .unwrap();
    assert_eq!(styles.cell_style_xfs.len(), 2);
    assert_eq!(styles.cell_xfs.len(), 2);
    assert_eq!(
        styles.cell_style_xfs[1]
            .alignment
            .as_ref()
            .unwrap()
            .vertical
            .as_deref(),
        Some("top")
    );
    let out = xml_of(&styles);
    assert!(out.contains(r#"<cellStyleXfs count="2">"#), "{out}");
    assert!(
        out.contains(r#"<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="1" applyNumberFormat="1"/>"#),
        "{out}"
    );
    let again = Styles::parse(out.as_bytes()).unwrap();
    assert_eq!(
        serde_json::to_value(&again).unwrap(),
        serde_json::to_value(&styles).unwrap()
    );
}

#[test]
fn verbatim_children_declare_inherited_prefixes() {
    let xml = concat!(
        "\u{feff}",
        r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="urn:x14" xmlns:mc="urn:mc">"#,
        r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>"#,
        "\n  ",
        r#"<colors><mruColors><color rgb="FF00B050"/></mruColors></colors>"#,
        r#"<extLst><ext uri="{A}"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst>"#,
        "</styleSheet>",
    );
    let styles = Styles::parse(xml.as_bytes()).unwrap();
    assert_eq!(
        styles.colors_xml.as_deref(),
        Some(r#"<colors><mruColors><color rgb="FF00B050"/></mruColors></colors>"#)
    );
    assert_eq!(
        styles.ext_lst_xml.as_deref(),
        Some(
            r#"<extLst xmlns:x14="urn:x14"><ext uri="{A}"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst>"#
        )
    );
    // The mruColors colour is not taken for a style colour.
    assert!(styles.fonts.is_empty());
    let out = xml_of(&styles);
    assert!(
        out.ends_with(r#"</cellXfs><colors><mruColors><color rgb="FF00B050"/></mruColors></colors><extLst xmlns:x14="urn:x14"><ext uri="{A}"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst></styleSheet>"#),
        "{out}"
    );
    let again = Styles::parse(out.as_bytes()).unwrap();
    assert_eq!(again.ext_lst_xml, styles.ext_lst_xml);
    assert_eq!(again.colors_xml, styles.colors_xml);
}

#[test]
fn nested_ext_lst_is_skipped() {
    let styles = Styles::parse(&sheet(concat!(
        r#"<dxfs count="1"><dxf><font><b/></font>"#,
        r#"<extLst><ext uri="{X}"><font><i/></font></ext></extLst></dxf></dxfs>"#,
    )))
    .unwrap();
    let font = styles.dxfs[0].font.as_ref().unwrap();
    assert!(font.bold && !font.italic);
    assert_eq!(styles.ext_lst_xml, None);
}

/// `xl/styles.xml` of the Excel-authored fixture (tests/fixtures/msoffcrypto/README.md).
fn excel_styles() -> Vec<u8> {
    use crate::zip::reader::{ZipSecurityLimits, read_zip_entries};
    let pkg = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/msoffcrypto/example.xlsx"
    ));
    let mut entries = read_zip_entries(pkg, &ZipSecurityLimits::default()).unwrap();
    entries.remove("xl/styles.xml").unwrap()
}

#[test]
fn excel_styles_round_trip() {
    let source = excel_styles();
    let first = Styles::parse(&source).unwrap();
    assert_eq!(first.fonts.len(), 2);
    assert_eq!(
        first.fonts[0].color_spec,
        Some(Color {
            theme: Some(1),
            ..Color::default()
        })
    );
    assert_eq!(first.cell_style_xfs.len(), 1);
    assert_eq!(first.cell_xfs[1].xf_id, Some(0));

    let out = String::from_utf8(first.to_xml().unwrap()).unwrap();
    let second = Styles::parse(out.as_bytes()).unwrap();
    assert_eq!(
        serde_json::to_value(&second).unwrap(),
        serde_json::to_value(&first).unwrap()
    );
    assert!(out.contains(r#"<cellStyleXfs count="1">"#), "{out}");
    assert_eq!(out.matches(r#" xfId="0""#).count(), 3, "{out}");
    assert!(
        out.contains(r#"defaultTableStyle="TableStyleMedium2""#),
        "{out}"
    );
    assert!(
        out.contains(r#"<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}""#),
        "{out}"
    );
    assert!(out.contains(r#"<color theme="1"/>"#), "{out}");

    // The verbatim extLst declares every prefix it uses.
    let ext = first.ext_lst_xml.as_deref().unwrap();
    for prefix in ["x14", "x15"] {
        assert!(ext.contains(&format!("<{prefix}:")), "{ext}");
        assert!(ext.contains(&format!("xmlns:{prefix}=")), "{ext}");
    }
    let decoys = [
        ("x14".to_owned(), "urn:decoy".to_owned()),
        ("x15".to_owned(), "urn:decoy".to_owned()),
    ]
    .into_iter()
    .collect::<xml::RootNamespaces>();
    assert_eq!(xml::self_contained_fragment(ext, &decoys), ext);
}

#[test]
fn self_closing_entries_keep_indices() {
    let styles = Styles::parse(&sheet(concat!(
        r#"<fonts count="2"><font/><font><b/></font></fonts>"#,
        r#"<fills count="2"><fill/><fill><patternFill patternType="gray125"/></fill></fills>"#,
        r#"<borders count="2"><border/><border><left style="thin"/></border></borders>"#,
        r#"<cellXfs count="1"><xf numFmtId="0" fontId="1" fillId="1" borderId="1"/></cellXfs>"#,
        r#"<dxfs count="2"><dxf/><dxf><font><b/></font></dxf></dxfs>"#,
    )))
    .unwrap();
    assert_eq!(
        [
            styles.fonts.len(),
            styles.fills.len(),
            styles.borders.len(),
            styles.dxfs.len()
        ],
        [2, 2, 2, 2]
    );
    let xf = &styles.cell_xfs[0];
    assert!(styles.fonts[xf.font_id as usize].bold);
    assert_eq!(styles.fills[xf.fill_id as usize].pattern_type, "gray125");
    assert!(styles.borders[xf.border_id as usize].left.is_some());
    assert!(styles.dxfs[1].font.as_ref().unwrap().bold);
    let out = xml_of(&styles);
    assert!(!out.contains(r#"patternType="""#), "{out}");
    assert!(out.contains("<fill><patternFill/></fill>"), "{out}");
}

#[test]
fn empty_pattern_fill_writes_no_pattern_type() {
    let styles = Styles::parse(&sheet(
        r#"<fills count="1"><fill><patternFill/></fill></fills>"#,
    ))
    .unwrap();
    assert_eq!(styles.fills.len(), 1);
    let out = xml_of(&styles);
    assert!(!out.contains(r#"patternType="""#), "{out}");
}

#[test]
fn gradient_keeps_theme_colour_stops() {
    let styles = Styles::parse(&sheet(concat!(
        r#"<fills count="1"><fill><gradientFill degree="90">"#,
        r#"<stop position="0"><color theme="0"/></stop>"#,
        r#"<stop position="1"><color rgb="FF4472C4"/></stop>"#,
        r#"</gradientFill></fill></fills>"#,
    )))
    .unwrap();
    let g = styles.fills[0].gradient_fill.as_ref().unwrap();
    assert_eq!(g.stops.len(), 2);
    assert_eq!(g.stops[0].color, "");
    assert_eq!(g.stops[1].color, "FF4472C4");
    let out = xml_of(&styles);
    assert!(!out.contains(r#"rgb="""#), "{out}");
    let again = Styles::parse(out.as_bytes()).unwrap();
    assert_eq!(
        again.fills[0].gradient_fill.as_ref().unwrap().stops.len(),
        2
    );
}

#[test]
fn num_fmt_format_code_is_unescaped_once() {
    let styles = Styles::parse(&sheet(
        r#"<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0"/></numFmts>"#,
    ))
    .unwrap();
    assert_eq!(styles.num_fmts[0].format_code, r##""$"#,##0"##);
    let out = xml_of(&styles);
    assert!(out.contains(r#"formatCode="&quot;$&quot;#,##0""#), "{out}");
    assert!(!out.contains("&amp;quot;"), "{out}");
}

/// A one-sheet package written by `write_xlsx` with `xl/styles.xml` replaced.
fn package_with_styles(styles_xml: &[u8]) -> Vec<u8> {
    use crate::zip::reader::{ZipSecurityLimits, read_zip_entries};
    use crate::zip::writer::{ZipEntry, write_zip};
    let wb = crate::WorkbookData {
        sheets: vec![crate::SheetData {
            name: "S".into(),
            ..Default::default()
        }],
        ..Default::default()
    };
    let bytes = crate::writer::write_xlsx(&wb).unwrap();
    let mut entries: Vec<ZipEntry> = read_zip_entries(&bytes, &ZipSecurityLimits::default())
        .unwrap()
        .into_iter()
        .map(|(name, data)| ZipEntry { name, data })
        .collect();
    entries.sort_by(|a, b| a.name.cmp(&b.name));
    let styles = entries
        .iter_mut()
        .find(|e| e.name == "xl/styles.xml")
        .unwrap();
    styles.data = styles_xml.to_vec();
    write_zip(&entries).unwrap()
}

#[test]
fn malformed_style_ids_fail_the_read() {
    let bad = package_with_styles(&sheet(
        r#"<numFmts count="1"><numFmt numFmtId="x" formatCode="0.00"/></numFmts>"#,
    ));
    let err = crate::reader::read_xlsx(&bad).unwrap_err();
    assert_eq!(err.code(), "XML_PARSE", "{err}");

    let spaced = package_with_styles(&sheet(
        r#"<numFmts count="1"><numFmt numFmtId=" 164 " formatCode="0.00"/></numFmts>"#,
    ));
    let wb = crate::reader::read_xlsx(&spaced).unwrap();
    assert_eq!(wb.styles.num_fmts[0].id, 164);

    let e = Styles::parse(&sheet(
        r#"<cellXfs count="1"><xf numFmtId="0" fontId="one" fillId="0" borderId="0"/></cellXfs>"#,
    ))
    .unwrap_err();
    assert!(matches!(e, ModernXlsxError::XmlParse(_)), "{e:?}");
}

#[test]
fn malformed_xml_is_xml_parse() {
    assert!(matches!(
        Styles::parse(b"<styleSheet><a></b></styleSheet>"),
        Err(ModernXlsxError::XmlParse(_))
    ));
}

/// A styles.xml whose colours are all rgb-only.
const RGB_ONLY: &str = concat!(
    r#"<fonts count="1"><font><b/><sz val="10.5"/><color rgb="FF112233"/><name val="Arial"/></font></fonts>"#,
    r#"<fills count="2"><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor rgb="FF000000"/></patternFill></fill>"#,
    r#"<fill><gradientFill degree="45"><stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color rgb="FF00FF00"/></stop></gradientFill></fill></fills>"#,
    r#"<borders count="1"><border><left style="thin"><color rgb="FF0000FF"/></left><right/><top/><bottom/><diagonal/></border></borders>"#,
    r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>"#,
    r#"<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>"#,
);

#[test]
fn rgb_only_json_is_unchanged() {
    let styles = Styles::parse(&sheet(RGB_ONLY)).unwrap();
    // Snapshot of the 1.2.1 JSON for the same input.
    let expected = concat!(
        r#"{"numFmts":[],"fonts":[{"name":"Arial","size":10.5,"bold":true,"italic":false,"underline":false,"strike":false,"color":"FF112233"}],"#,
        r#""fills":[{"patternType":"solid","fgColor":"FFFFFF00","bgColor":"FF000000"},{"patternType":"","fgColor":null,"bgColor":null,"gradientFill":{"degree":45.0,"stops":[{"position":0.0,"color":"FFFF0000"},{"position":1.0,"color":"FF00FF00"}]}}],"#,
        r#""borders":[{"left":{"style":"thin","color":"FF0000FF"},"right":null,"top":null,"bottom":null}],"#,
        r#""cellXfs":[{"numFmtId":0,"fontId":0,"fillId":0,"borderId":0}],"#,
        r#""dxfs":[{"font":{"name":null,"size":null,"bold":false,"italic":false,"underline":false,"strike":false,"color":"FF9C0006"},"fill":{"patternType":"","fgColor":null,"bgColor":"FFFFC7CE"}}]}"#,
    );
    assert_eq!(serde_json::to_string(&styles).unwrap(), expected);
}

/// A dxf font's explicit "off" (`<b val="0"/>`, `<u val="none"/>`) overrides
/// the cell's value, so it must survive a round trip instead of becoming
/// "inherit" (an absent element).
#[test]
fn dxf_explicit_off_properties_round_trip() {
    let styles = Styles::parse(&sheet(
        r#"<dxfs count="1"><dxf><font><b val="0"/><i val="0"/><u val="none"/></font></dxf></dxfs>"#,
    ))
    .unwrap();
    let font = styles.dxfs[0].font.as_ref().unwrap();
    assert!(!font.bold && !font.italic && !font.underline);
    let out = xml_of(&styles);
    assert!(
        out.contains(r#"<dxf><font><b val="0"/><i val="0"/><u val="none"/></font></dxf>"#),
        "{out}"
    );
    let again = Styles::parse(out.as_bytes()).unwrap();
    let out2 = xml_of(&again);
    assert!(
        out2.contains(r#"<dxf><font><b val="0"/><i val="0"/><u val="none"/></font></dxf>"#),
        "{out2}"
    );
}

/// CT_Xf `quotePrefix` and `pivotButton` survive a read-then-write; an xf
/// without them writes neither attribute.
#[test]
fn xf_quote_prefix_and_pivot_button_round_trip() {
    let styles = Styles::parse(&sheet(
        r#"<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" quotePrefix="1" pivotButton="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>"#,
    ))
    .unwrap();
    let out = xml_of(&styles);
    let again = Styles::parse(out.as_bytes()).unwrap();
    assert!(again.cell_xfs[0].quote_prefix, "{out}");
    assert!(again.cell_xfs[0].pivot_button, "{out}");
    assert!(!again.cell_xfs[1].quote_prefix && !again.cell_xfs[1].pivot_button);
    assert_eq!(out.matches("quotePrefix").count(), 1, "{out}");
    assert_eq!(out.matches("pivotButton").count(), 1, "{out}");
}

/// A C0 control character in a user style string is removed on write, so
/// `styles.xml` stays well-formed and reads back clean.
#[test]
fn style_strings_drop_xml_illegal_characters() {
    let styles = Styles {
        num_fmts: vec![NumFmt {
            id: 164,
            format_code: "0\u{0001}".to_owned(),
        }],
        fonts: vec![Font {
            name: Some("Ar\u{0001}ial".to_owned()),
            ..Font::default()
        }],
        ..Styles::default()
    };
    let out = xml_of(&styles);
    assert!(!out.contains('\u{0001}'), "{out:?}");
    let again = Styles::parse(out.as_bytes()).unwrap();
    assert_eq!(again.num_fmts[0].format_code, "0");
    assert_eq!(again.fonts[0].name.as_deref(), Some("Arial"));
}

#[test]
fn malformed_xml_error_names_the_part() {
    match Styles::parse(b"<styleSheet><fonts></styleSheet>") {
        Err(ModernXlsxError::XmlParse(msg)) => assert!(msg.contains("styles.xml"), "{msg}"),
        other => panic!("expected XmlParse, got {other:?}"),
    }
}
