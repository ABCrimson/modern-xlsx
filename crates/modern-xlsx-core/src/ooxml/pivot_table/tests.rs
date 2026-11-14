use super::*;

fn sample_pivot_table_xml() -> &'static [u8] {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1" cacheId="0" dataCaption="Values">
  <location ref="A3:C11" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields count="3">
<pivotField axis="axisRow" compact="0" outline="0">
  <items count="2">
    <item x="0"/>
    <item t="default"/>
  </items>
</pivotField>
<pivotField axis="axisCol">
  <items count="2">
    <item x="0"/>
    <item t="default"/>
  </items>
</pivotField>
<pivotField dataField="1"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <colFields count="1"><field x="1"/></colFields>
  <dataFields count="1">
<dataField name="Sum of Amount" fld="2" subtotal="sum"/>
  </dataFields>
  <pageFields count="1">
<pageField fld="0" item="0" name="Filter"/>
  </pageFields>
</pivotTableDefinition>"#
}

#[test]
fn pivot_table_roundtrip() {
    let pt = PivotTableData::parse(sample_pivot_table_xml()).unwrap();
    let xml = pt.to_xml().unwrap();
    let pt2 = PivotTableData::parse(&xml).unwrap();
    assert_eq!(pt2.name, pt.name);
    assert_eq!(pt2.cache_id, pt.cache_id);
    assert_eq!(pt2.data_caption, pt.data_caption);
    assert_eq!(pt2.location.ref_range, pt.location.ref_range);
    assert_eq!(pt2.location.first_header_row, pt.location.first_header_row);
    assert_eq!(pt2.location.first_data_row, pt.location.first_data_row);
    assert_eq!(pt2.location.first_data_col, pt.location.first_data_col);
    assert_eq!(pt2.pivot_fields.len(), pt.pivot_fields.len());
    for (a, b) in pt2.pivot_fields.iter().zip(&pt.pivot_fields) {
        assert_eq!(a.axis, b.axis);
        assert_eq!(a.compact, b.compact);
        assert_eq!(a.outline, b.outline);
        assert_eq!(a.items.len(), b.items.len());
    }
    assert_eq!(pt2.row_fields.len(), pt.row_fields.len());
    assert_eq!(pt2.col_fields.len(), pt.col_fields.len());
    assert_eq!(pt2.data_fields.len(), pt.data_fields.len());
    assert_eq!(pt2.data_fields[0].subtotal, pt.data_fields[0].subtotal);
}

#[test]
fn parse_pivot_cache_definition() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
recordCount="4">
  <cacheSource type="worksheet">
<worksheetSource ref="A1:C5" sheet="Data"/>
  </cacheSource>
  <cacheFields count="3">
<cacheField name="Region">
  <sharedItems count="2">
    <s v="East"/>
    <s v="West"/>
  </sharedItems>
</cacheField>
<cacheField name="Product">
  <sharedItems count="2">
    <s v="Widget"/>
    <s v="Gadget"/>
  </sharedItems>
</cacheField>
<cacheField name="Amount">
  <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1">
    <n v="100"/>
    <n v="200"/>
  </sharedItems>
</cacheField>
  </cacheFields>
</pivotCacheDefinition>"#;
    let cache = PivotCacheDefinitionData::parse(xml).unwrap();
    assert_eq!(cache.source.sheet, "Data");
    assert_eq!(cache.source.ref_range, "A1:C5");
    assert_eq!(cache.fields.len(), 3);
    assert_eq!(cache.fields[0].name, "Region");
    assert_eq!(cache.fields[0].shared_items.len(), 2);
    assert_eq!(cache.record_count, Some(4));
}

#[test]
fn parse_pivot_cache_records() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4">
  <r><x v="0"/><x v="0"/><n v="100"/></r>
  <r><x v="0"/><x v="1"/><n v="200"/></r>
  <r><x v="1"/><x v="0"/><n v="150"/></r>
  <r><x v="1"/><x v="1"/><n v="250"/></r>
</pivotCacheRecords>"#;
    let records = PivotCacheRecordsData::parse(xml).unwrap();
    assert_eq!(records.records.len(), 4);
    assert_eq!(records.records[0].len(), 3);
}

#[test]
fn pivot_cache_roundtrip() {
    // Roundtrip cache definition.
    let def_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
recordCount="4">
  <cacheSource type="worksheet">
<worksheetSource ref="A1:C5" sheet="Data"/>
  </cacheSource>
  <cacheFields count="3">
<cacheField name="Region">
  <sharedItems count="2">
    <s v="East"/>
    <s v="West"/>
  </sharedItems>
</cacheField>
<cacheField name="Product">
  <sharedItems count="2">
    <s v="Widget"/>
    <s v="Gadget"/>
  </sharedItems>
</cacheField>
<cacheField name="Amount">
  <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1">
    <n v="100"/>
    <n v="200"/>
  </sharedItems>
</cacheField>
  </cacheFields>
</pivotCacheDefinition>"#;
    let def1 = PivotCacheDefinitionData::parse(def_xml).unwrap();
    let out1 = def1.to_xml().unwrap();
    let def2 = PivotCacheDefinitionData::parse(&out1).unwrap();
    assert_eq!(def2.source.sheet, def1.source.sheet);
    assert_eq!(def2.source.ref_range, def1.source.ref_range);
    assert_eq!(def2.record_count, def1.record_count);
    assert_eq!(def2.fields.len(), def1.fields.len());
    for (a, b) in def2.fields.iter().zip(&def1.fields) {
        assert_eq!(a.name, b.name);
        assert_eq!(a.shared_items.len(), b.shared_items.len());
    }

    // Roundtrip cache records.
    let rec_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4">
  <r><x v="0"/><x v="0"/><n v="100"/></r>
  <r><x v="0"/><x v="1"/><n v="200"/></r>
  <r><x v="1"/><x v="0"/><n v="150"/></r>
  <r><x v="1"/><x v="1"/><n v="250"/></r>
</pivotCacheRecords>"#;
    let rec1 = PivotCacheRecordsData::parse(rec_xml).unwrap();
    let out2 = rec1.to_xml().unwrap();
    let rec2 = PivotCacheRecordsData::parse(&out2).unwrap();
    assert_eq!(rec2.records.len(), rec1.records.len());
    for (a, b) in rec2.records.iter().zip(&rec1.records) {
        assert_eq!(a.len(), b.len());
    }
}

#[test]
fn parse_pivot_table_definition() {
    let pt = PivotTableData::parse(sample_pivot_table_xml()).unwrap();
    assert_eq!(pt.name, "PivotTable1");
    assert_eq!(pt.cache_id, 0);
    assert_eq!(pt.data_caption.as_deref(), Some("Values"));
    assert_eq!(pt.location.ref_range, "A3:C11");
    assert_eq!(pt.location.first_header_row, Some(1));
    assert_eq!(pt.location.first_data_row, Some(2));
    assert_eq!(pt.location.first_data_col, Some(1));
    assert_eq!(pt.pivot_fields.len(), 3);
    assert_eq!(pt.pivot_fields[0].axis, Some(PivotAxis::AxisRow));
    assert!(!pt.pivot_fields[0].compact);
    assert!(!pt.pivot_fields[0].outline);
    assert_eq!(pt.pivot_fields[0].items.len(), 2);
    assert_eq!(pt.pivot_fields[0].items[0].x, Some(0));
    assert_eq!(pt.pivot_fields[0].items[1].t.as_deref(), Some("default"));
    assert_eq!(pt.pivot_fields[1].axis, Some(PivotAxis::AxisCol));
    assert_eq!(pt.pivot_fields[2].axis, None); // dataField only, no axis
    assert_eq!(pt.row_fields.len(), 1);
    assert_eq!(pt.row_fields[0].x, 0);
    assert_eq!(pt.col_fields.len(), 1);
    assert_eq!(pt.col_fields[0].x, 1);
    assert_eq!(pt.data_fields.len(), 1);
    assert_eq!(pt.data_fields[0].name.as_deref(), Some("Sum of Amount"));
    assert_eq!(pt.data_fields[0].fld, 2);
    assert_eq!(pt.data_fields[0].subtotal, SubtotalFunction::Sum);
}

fn pivot_wrap(body: &str) -> Vec<u8> {
    format!(
        r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="0"><location ref="A1:B2"/>{body}</pivotTableDefinition>"#
    )
    .into_bytes()
}

fn cache_wrap(fields: &str) -> Vec<u8> {
    format!(
        r#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource ref="A1:B2" sheet="S"/></cacheSource><cacheFields count="1">{fields}</cacheFields></pivotCacheDefinition>"#
    )
    .into_bytes()
}

#[test]
fn attributes_are_unescaped_on_read_and_escaped_once_on_write() {
    let xml = br#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet"><worksheetSource ref="A1:B2" sheet="Q&amp;A"/></cacheSource>
  <cacheFields count="1"><cacheField name="F &amp; G"><sharedItems count="1"><s v="A &amp; B"/></sharedItems></cacheField></cacheFields>
</pivotCacheDefinition>"#;
    let def = PivotCacheDefinitionData::parse(xml).unwrap();
    assert_eq!(def.source.sheet, "Q&A");
    assert_eq!(def.fields[0].name, "F & G");
    assert_eq!(
        def.fields[0].shared_items[0],
        CacheValue::String { v: "A & B".into() }
    );
    let out = String::from_utf8(def.to_xml().unwrap()).unwrap();
    assert!(out.contains(r#"sheet="Q&amp;A""#), "{out}");
    assert!(out.contains(r#"name="F &amp; G""#), "{out}");
    assert!(out.contains(r#"<s v="A &amp; B"/>"#), "{out}");
    assert!(!out.contains("&amp;amp;"), "{out}");
    assert_eq!(
        PivotCacheDefinitionData::parse(out.as_bytes()).unwrap(),
        def
    );

    let pt = PivotTableData::parse(&pivot_wrap("")).unwrap();
    let pt = PivotTableData {
        name: "P & L".into(),
        ..pt
    };
    let out = String::from_utf8(pt.to_xml().unwrap()).unwrap();
    assert!(out.contains(r#"name="P &amp; L""#), "{out}");
    assert_eq!(PivotTableData::parse(out.as_bytes()).unwrap().name, "P & L");
}

#[test]
fn data_field_and_subtotal_flags_round_trip() {
    let xml = pivot_wrap(
        r#"<pivotFields count="1"><pivotField dataField="1" defaultSubtotal="0" sumSubtotal="1" countASubtotal="1"/></pivotFields>"#,
    );
    let pt = PivotTableData::parse(&xml).unwrap();
    let f = &pt.pivot_fields[0];
    assert!(f.data_field);
    assert_eq!(f.default_subtotal, Some(false));
    assert_eq!(
        f.subtotals,
        [SubtotalFunction::Sum, SubtotalFunction::Count]
    );
    let out = String::from_utf8(pt.to_xml().unwrap()).unwrap();
    for attr in [
        r#"dataField="1""#,
        r#"defaultSubtotal="0""#,
        r#"sumSubtotal="1""#,
        r#"countASubtotal="1""#,
    ] {
        assert!(out.contains(attr), "{attr} missing in {out}");
    }
    let back = PivotTableData::parse(out.as_bytes()).unwrap();
    assert_eq!(back.pivot_fields, pt.pivot_fields);
}

#[test]
fn every_subtotal_function_has_its_own_attribute() {
    let all = [
        SubtotalFunction::Sum,
        SubtotalFunction::Count,
        SubtotalFunction::Average,
        SubtotalFunction::Max,
        SubtotalFunction::Min,
        SubtotalFunction::Product,
        SubtotalFunction::CountNums,
        SubtotalFunction::StdDev,
        SubtotalFunction::StdDevP,
        SubtotalFunction::Var,
        SubtotalFunction::VarP,
    ];
    for v in all {
        assert_eq!(SubtotalFunction::from_xml(v.xml_val()), Some(v));
    }
    let pairs = [
        (SubtotalFunction::Sum, "sumSubtotal"),
        (SubtotalFunction::Count, "countASubtotal"),
        (SubtotalFunction::Average, "avgSubtotal"),
        (SubtotalFunction::Max, "maxSubtotal"),
        (SubtotalFunction::Min, "minSubtotal"),
        (SubtotalFunction::Product, "productSubtotal"),
        (SubtotalFunction::CountNums, "countSubtotal"),
        (SubtotalFunction::StdDev, "stdDevSubtotal"),
        (SubtotalFunction::StdDevP, "stdDevPSubtotal"),
        (SubtotalFunction::Var, "varSubtotal"),
        (SubtotalFunction::VarP, "varPSubtotal"),
    ];
    for (v, name) in pairs {
        assert_eq!(v.subtotal_attr(), name);
        assert_eq!(SubtotalFunction::from_subtotal_attr(name), Some(v));
    }
    assert_eq!(SubtotalFunction::from_subtotal_attr("bogusSubtotal"), None);
    assert_eq!(SubtotalFunction::from_xml("bogus"), None);
    for v in [
        PivotAxis::AxisRow,
        PivotAxis::AxisCol,
        PivotAxis::AxisPage,
        PivotAxis::AxisValues,
    ] {
        assert_eq!(PivotAxis::from_xml(v.xml_val()), Some(v));
    }
    assert_eq!(PivotAxis::from_xml("bogus"), None);
}

#[test]
fn caption_offsets_and_field_order_are_always_written() {
    let pt = PivotTableData {
        name: "P".into(),
        data_caption: None,
        location: PivotLocation {
            ref_range: "A1:C5".into(),
            first_header_row: None,
            first_data_row: None,
            first_data_col: None,
        },
        pivot_fields: Vec::new(),
        row_fields: vec![PivotFieldRef { x: 0 }],
        col_fields: vec![PivotFieldRef { x: 1 }],
        data_fields: vec![PivotDataFieldData {
            name: None,
            fld: 2,
            subtotal: SubtotalFunction::Sum,
            num_fmt_id: None,
        }],
        page_fields: vec![PivotPageFieldData {
            fld: 3,
            item: None,
            name: None,
        }],
        cache_id: 0,
    };
    let out = String::from_utf8(pt.to_xml().unwrap()).unwrap();
    assert!(out.contains(r#"dataCaption="Values""#), "{out}");
    assert!(
        out.contains(r#"firstHeaderRow="1" firstDataRow="2" firstDataCol="1""#),
        "{out}"
    );
    assert!(out.find("<pageFields").unwrap() < out.find("<dataFields").unwrap());

    let flat = PivotTableData {
        row_fields: Vec::new(),
        col_fields: Vec::new(),
        ..pt
    };
    let out = String::from_utf8(flat.to_xml().unwrap()).unwrap();
    assert!(
        out.contains(r#"firstHeaderRow="1" firstDataRow="1" firstDataCol="0""#),
        "{out}"
    );
    // None stays None on read when the file lacks the attributes.
    let back = PivotTableData::parse(&pivot_wrap("")).unwrap();
    assert_eq!(back.data_caption, None);
    assert_eq!(back.location.first_data_row, None);
}

#[test]
fn shared_items_attributes_round_trip() {
    let xml = cache_wrap(
        r#"<cacheField name="Amt" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="1" maxValue="9"/></cacheField>"#,
    );
    let def = PivotCacheDefinitionData::parse(&xml).unwrap();
    let f = &def.fields[0];
    assert_eq!(f.num_fmt_id, Some(0));
    let attrs = f.shared_items_attrs.as_ref().unwrap();
    assert_eq!(attrs.contains_semi_mixed_types, Some(false));
    assert_eq!(attrs.contains_string, Some(false));
    assert_eq!(attrs.contains_number, Some(true));
    assert_eq!(attrs.min_value, Some(1.0));
    assert_eq!(attrs.max_value, Some(9.0));
    let out = String::from_utf8(def.to_xml().unwrap()).unwrap();
    for attr in [
        r#"numFmtId="0""#,
        r#"containsSemiMixedTypes="0""#,
        r#"containsString="0""#,
        r#"containsNumber="1""#,
        r#"minValue="1""#,
        r#"maxValue="9""#,
    ] {
        assert!(out.contains(attr), "{attr} missing in {out}");
    }
    assert!(!out.contains("count=\"0\""), "{out}");
    assert_eq!(
        PivotCacheDefinitionData::parse(out.as_bytes()).unwrap(),
        def
    );

    // Dates and long text survive too.
    let xml = cache_wrap(
        r#"<cacheField name="D"><sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0" minDate="2020-01-01T00:00:00" maxDate="2021-01-01T00:00:00" longText="1" count="1"><d v="2020-01-01T00:00:00"/></sharedItems></cacheField>"#,
    );
    let def = PivotCacheDefinitionData::parse(&xml).unwrap();
    let out = def.to_xml().unwrap();
    assert_eq!(PivotCacheDefinitionData::parse(&out).unwrap(), def);
}

#[test]
fn compact_and_outline_are_always_serialised() {
    let xml = pivot_wrap(r#"<pivotFields count="1"><pivotField compact="0"/></pivotFields>"#);
    let pt = PivotTableData::parse(&xml).unwrap();
    let json = serde_json::to_string(&pt).unwrap();
    assert!(json.contains(r#""compact":false"#), "{json}");
    assert!(json.contains(r#""outline":true"#), "{json}");
    let back: PivotTableData = serde_json::from_str(&json).unwrap();
    let out = String::from_utf8(back.to_xml().unwrap()).unwrap();
    assert!(out.contains(r#"compact="0""#), "{out}");
    assert!(!out.contains("outline="), "{out}");

    let bare: PivotFieldData = serde_json::from_str("{}").unwrap();
    assert!(bare.compact && bare.outline);
    let pt = PivotTableData {
        pivot_fields: vec![bare],
        ..PivotTableData::parse(&pivot_wrap("")).unwrap()
    };
    let out = String::from_utf8(pt.to_xml().unwrap()).unwrap();
    assert!(
        !out.contains("compact=") && !out.contains("outline="),
        "{out}"
    );
}

#[test]
fn pivot_table_full_value_round_trip_with_page_fields() {
    let xml = pivot_wrap(
        r#"<pivotFields count="2"><pivotField axis="axisRow" name="A &amp; B"><items count="2"><item x="0"/><item t="default"/></items></pivotField><pivotField dataField="1"/></pivotFields>
<rowFields count="1"><field x="0"/></rowFields>
<colFields count="1"><field x="-2"/></colFields>
<pageFields count="1"><pageField fld="0" item="1" name="Pg"/></pageFields>
<dataFields count="1"><dataField name="Sum" fld="1" subtotal="sum" numFmtId="3"/></dataFields>"#,
    );
    let mut pt = PivotTableData::parse(&xml).unwrap();
    // Fill the defaults the writer adds, so the comparison is exact.
    pt.data_caption = Some("Values".into());
    pt.location.first_header_row = Some(1);
    pt.location.first_data_row = Some(2);
    pt.location.first_data_col = Some(1);
    assert_eq!(pt.col_fields[0].x, -2);
    assert_eq!(PivotTableData::parse(&pt.to_xml().unwrap()).unwrap(), pt);
}

#[test]
fn cache_definition_and_records_round_trip_full_value() {
    let def = PivotCacheDefinitionData {
        cache_id: None,
        source: CacheSource {
            ref_range: "A1:B3".into(),
            sheet: "Q&A".into(),
            ..Default::default()
        },
        fields: vec![CacheFieldData {
            name: "F".into(),
            num_fmt_id: Some(14),
            shared_items: vec![CacheValue::Number { v: 1.5 }, CacheValue::Missing],
            shared_items_attrs: Some(SharedItemsAttrs {
                contains_blank: Some(true),
                min_value: Some(1.5),
                ..Default::default()
            }),
        }],
        record_count: Some(2),
        relationships: Vec::new(),
    };
    assert_eq!(
        PivotCacheDefinitionData::parse(&def.to_xml().unwrap()).unwrap(),
        def
    );

    let recs = PivotCacheRecordsData {
        records: vec![vec![
            CacheValue::String { v: "x&y".into() },
            CacheValue::Number { v: -0.25 },
            CacheValue::Number { v: 3.0 },
            CacheValue::Boolean { v: true },
            CacheValue::Boolean { v: false },
            CacheValue::DateTime {
                v: "2024-01-01T00:00:00".into(),
            },
            CacheValue::Missing,
            CacheValue::Error { v: "#N/A".into() },
            CacheValue::Index { v: 7 },
        ]],
    };
    assert_eq!(
        PivotCacheRecordsData::parse(&recs.to_xml().unwrap()).unwrap(),
        recs
    );
}

#[test]
fn malformed_numbers_and_markup_are_errors() {
    assert!(
        PivotTableData::parse(&pivot_wrap(
            r#"<rowFields count="1"><field x="a"/></rowFields>"#
        ))
        .is_err()
    );
    assert!(PivotTableData::parse(b"<root><a></b></root>").is_err());
    assert!(PivotCacheDefinitionData::parse(b"<root><a></b></root>").is_err());
    assert!(PivotCacheRecordsData::parse(b"<root><a></b></root>").is_err());
    assert!(
        PivotCacheRecordsData::parse(
            br#"<pivotCacheRecords><r><n v="abc"/></r></pivotCacheRecords>"#
        )
        .is_err()
    );
    assert!(
        PivotCacheRecordsData::parse(
            br#"<pivotCacheRecords><r><s v="&bogus;"/></r></pivotCacheRecords>"#
        )
        .is_err()
    );
}

#[test]
fn empty_cache_fields_are_still_written() {
    // CT_PivotCacheDefinition requires <cacheFields> (minOccurs=1).
    let def = PivotCacheDefinitionData {
        cache_id: None,
        source: CacheSource {
            ref_range: "A1:A1".into(),
            sheet: "Data".into(),
            ..Default::default()
        },
        fields: Vec::new(),
        record_count: None,
        relationships: Vec::new(),
    };
    let out = String::from_utf8(def.to_xml().unwrap()).unwrap();
    assert!(
        out.contains(r#"</cacheSource><cacheFields count="0"/></pivotCacheDefinition>"#),
        "{out}"
    );
    let back = PivotCacheDefinitionData::parse(out.as_bytes()).unwrap();
    assert!(back.fields.is_empty());
    assert_eq!(back, def);
}

#[test]
fn empty_records_and_start_form_values_keep_row_alignment() {
    let recs = PivotCacheRecordsData::parse(
        br#"<pivotCacheRecords count="2"><r/><r><n v="1"/></r></pivotCacheRecords>"#,
    )
    .unwrap();
    assert_eq!(
        recs.records,
        vec![vec![], vec![CacheValue::Number { v: 1.0 }]]
    );

    // Start-form values are read, and their member-property children are skipped.
    let recs = PivotCacheRecordsData::parse(
        br#"<pivotCacheRecords count="1"><r><s v="a"></s><n v="5"><x v="0"/></n></r></pivotCacheRecords>"#,
    )
    .unwrap();
    assert_eq!(
        recs.records,
        vec![vec![
            CacheValue::String { v: "a".into() },
            CacheValue::Number { v: 5.0 },
        ]]
    );
}

// ---------------------------------------------------------------------------
// Excel's own pivot markup (Excel 16, saved through COM on 2026-11-12) and the
// parts the writer gives for it: completion adds nothing to a table Excel saved.
// ---------------------------------------------------------------------------

/// A pivot table with a row, a column and a page field and two data fields,
/// as Excel saved it.
pub(crate) const EXCEL_TWO_AXES_TABLE: &str = r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xr:uid="{983EEBE7-C6F8-4679-B272-3E5DA27C221D}" name="PT2" cacheId="9" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"><location ref="H5:N10" firstHeaderRow="1" firstDataRow="3" firstDataCol="1" rowPageCount="1" colPageCount="1"/><pivotFields count="5"><pivotField axis="axisRow" showAll="0"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField><pivotField axis="axisCol" showAll="0"><items count="3"><item x="1"/><item x="0"/><item t="default"/></items></pivotField><pivotField axis="axisPage" showAll="0"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField><pivotField dataField="1" showAll="0"/><pivotField dataField="1" showAll="0"/></pivotFields><rowFields count="1"><field x="0"/></rowFields><rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems><colFields count="2"><field x="1"/><field x="-2"/></colFields><colItems count="6"><i><x/><x/></i><i r="1" i="1"><x v="1"/></i><i><x v="1"/><x/></i><i r="1" i="1"><x v="1"/></i><i t="grand"><x/></i><i t="grand" i="1"><x/></i></colItems><pageFields count="1"><pageField fld="2" hier="-1"/></pageFields><dataFields count="2"><dataField name="Total Revenue" fld="3" baseField="0" baseItem="0"/><dataField name="Total Units" fld="4" baseField="0" baseItem="0"/></dataFields><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/><extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotTableDefinition hideValuesRow="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/></ext><ext uri="{747A6164-185A-40DC-8AA5-F01512510D54}" xmlns:xpdl="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout"><xpdl:pivotTableDefinition16/></ext></extLst></pivotTableDefinition>"#;

/// The cache definition of [`EXCEL_TWO_AXES_TABLE`], as Excel saved it
/// (`refreshedBy` anonymised).
pub(crate) const EXCEL_TWO_AXES_CACHE: &str = r#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" r:id="rId1" refreshedBy="Author" refreshedDate="46338.421466666667" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="4" xr:uid="{5F7E10E8-5AE1-43D4-B029-33AE3E6A6CC0}"><cacheSource type="worksheet"><worksheetSource ref="A1:E5" sheet="Sheet1"/></cacheSource><cacheFields count="5"><cacheField name="Region" numFmtId="0"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField><cacheField name="Product" numFmtId="0"><sharedItems count="2"><s v="Pen"/><s v="Ink"/></sharedItems></cacheField><cacheField name="Year" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="2025" maxValue="2026" count="2"><n v="2025"/><n v="2026"/></sharedItems></cacheField><cacheField name="Revenue" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="10" maxValue="40"/></cacheField><cacheField name="Units" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="4"/></cacheField></cacheFields><extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotCacheDefinition/></ext><ext uri="{F4336242-56EE-4DD3-A5F7-E900642BF6E9}" xmlns:xlpdsc="http://schemas.microsoft.com/office/spreadsheetml/2026/pivotDataSourceClean"><xlpdsc:dataSourceClean>1</xlpdsc:dataSourceClean></ext></extLst></pivotCacheDefinition>"#;

/// Two row fields, the outer with the Sum and Count subtotals and the inner
/// with none, as Excel saved them.
pub(crate) const EXCEL_SUBTOTALS_TABLE: &str = r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xr:uid="{CDD10DE9-1134-4BFC-B3B3-2E0A47EF825D}" name="PT3" cacheId="14" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"><location ref="E3:F14" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="3"><pivotField axis="axisRow" showAll="0" sumSubtotal="1" countASubtotal="1"><items count="4"><item x="0"/><item x="1"/><item t="sum"/><item t="countA"/></items></pivotField><pivotField axis="axisRow" showAll="0" defaultSubtotal="0"><items count="2"><item x="1"/><item x="0"/></items></pivotField><pivotField dataField="1" showAll="0"/></pivotFields><rowFields count="2"><field x="0"/><field x="1"/></rowFields><rowItems count="11"><i><x/></i><i r="1"><x/></i><i r="1"><x v="1"/></i><i t="sum"><x/></i><i t="countA"><x/></i><i><x v="1"/></i><i r="1"><x/></i><i r="1"><x v="1"/></i><i t="sum"><x v="1"/></i><i t="countA"><x v="1"/></i><i t="grand"><x/></i></rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField name="Total Revenue" fld="2" baseField="0" baseItem="0"/></dataFields><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/><extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotTableDefinition hideValuesRow="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/></ext><ext uri="{747A6164-185A-40DC-8AA5-F01512510D54}" xmlns:xpdl="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout"><xpdl:pivotTableDefinition16/></ext></extLst></pivotTableDefinition>"#;

/// [`EXCEL_TWO_AXES_TABLE`] read and written: the writer adds no items to it.
pub(crate) const EXCEL_TWO_AXES_TABLE_WRITTEN: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PT2" cacheId="9" dataCaption="Values"><location ref="H5:N10" firstHeaderRow="1" firstDataRow="3" firstDataCol="1"/><pivotFields count="5"><pivotField axis="axisRow"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField><pivotField axis="axisCol"><items count="3"><item x="1"/><item x="0"/><item t="default"/></items></pivotField><pivotField axis="axisPage"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField><pivotField dataField="1"/><pivotField dataField="1"/></pivotFields><rowFields count="1"><field x="0"/></rowFields><colFields count="2"><field x="1"/><field x="-2"/></colFields><pageFields count="1"><pageField fld="2"/></pageFields><dataFields count="2"><dataField name="Total Revenue" fld="3" subtotal="sum"/><dataField name="Total Units" fld="4" subtotal="sum"/></dataFields></pivotTableDefinition>"#;

/// [`EXCEL_SUBTOTALS_TABLE`] read and written: the writer adds no items to it.
pub(crate) const EXCEL_SUBTOTALS_TABLE_WRITTEN: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PT3" cacheId="14" dataCaption="Values"><location ref="E3:F14" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="3"><pivotField axis="axisRow" sumSubtotal="1" countASubtotal="1"><items count="4"><item x="0"/><item x="1"/><item t="sum"/><item t="countA"/></items></pivotField><pivotField axis="axisRow" defaultSubtotal="0"><items count="2"><item x="1"/><item x="0"/></items></pivotField><pivotField dataField="1"/></pivotFields><rowFields count="2"><field x="0"/><field x="1"/></rowFields><dataFields count="1"><dataField name="Total Revenue" fld="2" subtotal="sum"/></dataFields></pivotTableDefinition>"#;

/// [`EXCEL_TWO_AXES_CACHE`] read and written with records: no `refreshOnLoad`
/// is added, since its table needs no completion.
pub(crate) const EXCEL_TWO_AXES_CACHE_WRITTEN: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" recordCount="4"><cacheSource type="worksheet"><worksheetSource ref="A1:E5" sheet="Sheet1"/></cacheSource><cacheFields count="5"><cacheField name="Region" numFmtId="0"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField><cacheField name="Product" numFmtId="0"><sharedItems count="2"><s v="Pen"/><s v="Ink"/></sharedItems></cacheField><cacheField name="Year" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="2025" maxValue="2026" count="2"><n v="2025"/><n v="2026"/></sharedItems></cacheField><cacheField name="Revenue" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="10" maxValue="40"/></cacheField><cacheField name="Units" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="4"/></cacheField></cacheFields></pivotCacheDefinition>"#;

/// A table whose only pivot field is `field`, on the row axis when it has
/// one, with one data field.
fn one_field_table(field: PivotFieldData) -> PivotTableData {
    PivotTableData {
        row_fields: if field.axis.is_some() {
            vec![PivotFieldRef { x: 0 }]
        } else {
            Vec::new()
        },
        pivot_fields: vec![field],
        data_fields: vec![PivotDataFieldData {
            name: None,
            fld: 0,
            subtotal: SubtotalFunction::Sum,
            num_fmt_id: None,
        }],
        ..PivotTableData::parse(&pivot_wrap("")).unwrap()
    }
}

/// A field on `axis` with no items, the given subtotals and `defaultSubtotal`.
fn bare_field(
    axis: Option<PivotAxis>,
    subtotals: Vec<SubtotalFunction>,
    default_subtotal: Option<bool>,
) -> PivotFieldData {
    PivotFieldData {
        axis,
        name: None,
        items: Vec::new(),
        subtotals,
        compact: true,
        outline: true,
        data_field: axis.is_none(),
        default_subtotal,
    }
}

/// A cache whose first field has `shared` string items.
fn cache_with_shared(shared: usize) -> PivotCacheDefinitionData {
    let mut cache =
        PivotCacheDefinitionData::parse(&cache_wrap(r#"<cacheField name="F"/>"#)).unwrap();
    cache.fields[0].shared_items = (0..shared)
        .map(|i| CacheValue::String { v: format!("v{i}") })
        .collect();
    cache
}

/// The `<pivotFields>` element of a written table.
fn pivot_fields_xml(xml: &[u8]) -> String {
    let out = String::from_utf8(xml.to_vec()).unwrap();
    let start = out.find("<pivotFields").unwrap();
    let end = out.find("</pivotFields>").unwrap() + "</pivotFields>".len();
    out[start..end].to_owned()
}

#[test]
fn axis_fields_without_items_get_the_items_excel_needs() {
    use PivotAxis::{AxisCol, AxisPage, AxisRow, AxisValues};
    use SubtotalFunction::{Count, Sum};
    // Excel 16 crashes on Refresh when an axis field whose
    // default subtotal applies has no items, refuses `<items count="0"/>`,
    // and writes one `<item x>` per shared item and then one item per
    // subtotal (`default`, or `sum`, `countA`, ... for custom subtotals).
    let x3 = r#"<item x="0"/><item x="1"/><item x="2"/>"#;
    // (label, field, shared items of its cache field (None: no cache),
    // expected <pivotField>, still without items after a read)
    let cases: [(&str, PivotFieldData, Option<usize>, String, bool); 9] = [
        (
            "row field, 3 shared items",
            bare_field(Some(AxisRow), Vec::new(), None),
            Some(3),
            format!(
                r#"<pivotField axis="axisRow"><items count="4">{x3}<item t="default"/></items></pivotField>"#
            ),
            false,
        ),
        (
            "column field, custom subtotals in Excel's order",
            bare_field(Some(AxisCol), vec![Count, Sum, Count], None),
            Some(3),
            format!(
                r#"<pivotField axis="axisCol" countASubtotal="1" sumSubtotal="1"><items count="5">{x3}<item t="sum"/><item t="countA"/></items></pivotField>"#
            ),
            false,
        ),
        (
            "page field without a default subtotal",
            bare_field(Some(AxisPage), Vec::new(), Some(false)),
            Some(3),
            format!(
                r#"<pivotField axis="axisPage" defaultSubtotal="0"><items count="3">{x3}</items></pivotField>"#
            ),
            false,
        ),
        (
            "explicit default subtotal",
            bare_field(Some(AxisRow), Vec::new(), Some(true)),
            Some(1),
            r#"<pivotField axis="axisRow" defaultSubtotal="1"><items count="2"><item x="0"/><item t="default"/></items></pivotField>"#.into(),
            false,
        ),
        (
            "cache field without shared items",
            bare_field(Some(AxisRow), Vec::new(), None),
            Some(0),
            r#"<pivotField axis="axisRow"><items count="1"><item t="default"/></items></pivotField>"#
                .into(),
            false,
        ),
        (
            "no cache at all",
            bare_field(Some(AxisRow), Vec::new(), None),
            None,
            r#"<pivotField axis="axisRow"><items count="1"><item t="default"/></items></pivotField>"#
                .into(),
            false,
        ),
        (
            "nothing to write: no shared item, no subtotal",
            bare_field(Some(AxisRow), Vec::new(), Some(false)),
            Some(0),
            r#"<pivotField axis="axisRow" defaultSubtotal="0"/>"#.into(),
            true,
        ),
        (
            "data-only field",
            bare_field(None, Vec::new(), None),
            Some(3),
            r#"<pivotField dataField="1"/>"#.into(),
            false,
        ),
        (
            "values axis",
            bare_field(Some(AxisValues), Vec::new(), None),
            Some(3),
            r#"<pivotField axis="axisValues"/>"#.into(),
            false,
        ),
    ];
    for (label, field, shared, expected, still_bare) in cases {
        let table = one_field_table(field);
        let before = table.clone();
        let cache = shared.map(cache_with_shared);
        let xml = match &cache {
            Some(cache) => table.to_xml_with_cache(Some(cache), &mut 0).unwrap(),
            None => table.to_xml().unwrap(),
        };
        assert_eq!(
            pivot_fields_xml(&xml),
            format!(r#"<pivotFields count="1">{expected}</pivotFields>"#),
            "{label}"
        );
        assert_eq!(table, before, "{label}: the model is unchanged");
        let back = PivotTableData::parse(&xml).unwrap();
        assert_eq!(back.has_axis_field_without_items(), still_bare, "{label}");
    }
}

#[test]
fn only_an_axis_field_without_items_needs_its_layout_computed() {
    for axis in [PivotAxis::AxisRow, PivotAxis::AxisCol, PivotAxis::AxisPage] {
        let table = one_field_table(bare_field(Some(axis), Vec::new(), None));
        assert!(table.has_axis_field_without_items(), "{axis:?}");
    }
    for field in [
        bare_field(None, Vec::new(), None),
        bare_field(Some(PivotAxis::AxisValues), Vec::new(), None),
    ] {
        assert!(!one_field_table(field).has_axis_field_without_items());
    }
    for xml in [EXCEL_TWO_AXES_TABLE, EXCEL_SUBTOTALS_TABLE] {
        let table = PivotTableData::parse(xml.as_bytes()).unwrap();
        assert!(!table.has_axis_field_without_items());
    }
}

#[test]
fn excel_tables_are_written_as_before() {
    let cache = PivotCacheDefinitionData::parse(EXCEL_TWO_AXES_CACHE.as_bytes()).unwrap();
    for (xml, written) in [
        (EXCEL_TWO_AXES_TABLE, EXCEL_TWO_AXES_TABLE_WRITTEN),
        (EXCEL_SUBTOTALS_TABLE, EXCEL_SUBTOTALS_TABLE_WRITTEN),
    ] {
        let table = PivotTableData::parse(xml.as_bytes()).unwrap();
        // Excel saves a table with one pivot field per field of its cache.
        let mut own_cache = cache.clone();
        own_cache.fields.truncate(table.pivot_fields.len());
        let mut added = 0;
        for out in [
            table.to_xml().unwrap(),
            table
                .to_xml_with_cache(Some(&own_cache), &mut added)
                .unwrap(),
        ] {
            assert_eq!(String::from_utf8(out).unwrap(), written);
        }
        assert_eq!(added, 0, "nothing is added to a table Excel saved");
    }
    assert_eq!(
        String::from_utf8(cache.to_xml_with_records(Some("rId1")).unwrap()).unwrap(),
        EXCEL_TWO_AXES_CACHE_WRITTEN
    );
}

#[test]
fn two_data_fields_put_the_values_field_on_the_columns() {
    // Excel 16 refuses a table with two data fields and no
    // `<field x="-2"/>` (the values field) in its row or column fields.
    let data = |fld| PivotDataFieldData {
        name: None,
        fld,
        subtotal: SubtotalFunction::Sum,
        num_fmt_id: None,
    };
    let refs = |xs: &[i32]| xs.iter().map(|&x| PivotFieldRef { x }).collect::<Vec<_>>();
    let mut base = one_field_table(bare_field(Some(PivotAxis::AxisRow), Vec::new(), None));
    base.data_fields = vec![data(1), data(2)];
    /// (label, row fields, column fields, expected `<colFields>`)
    type Case<'a> = (&'a str, &'a [i32], &'a [i32], Option<&'a str>);
    let cases: [Case<'_>; 4] = [
        (
            "no column field",
            &[0],
            &[],
            Some(r#"<colFields count="1"><field x="-2"/></colFields>"#),
        ),
        (
            "a column field",
            &[],
            &[0],
            Some(r#"<colFields count="2"><field x="0"/><field x="-2"/></colFields>"#),
        ),
        ("values already on the rows", &[0, -2], &[], None),
        (
            "values already on the columns",
            &[],
            &[-2, 0],
            Some(r#"<colFields count="2"><field x="-2"/><field x="0"/></colFields>"#),
        ),
    ];
    for (label, rows, cols, expected) in cases {
        let table = PivotTableData {
            row_fields: refs(rows),
            col_fields: refs(cols),
            ..base.clone()
        };
        let before = table.clone();
        let out = String::from_utf8(table.to_xml().unwrap()).unwrap();
        match expected {
            Some(cols) => assert!(out.contains(cols), "{label}: {out}"),
            None => assert!(!out.contains("<colFields"), "{label}: {out}"),
        }
        assert_eq!(table, before, "{label}: the model is unchanged");
    }
    // One data field needs no values field.
    base.data_fields.truncate(1);
    let out = String::from_utf8(base.to_xml().unwrap()).unwrap();
    assert!(!out.contains("<colFields"), "{out}");
}

#[test]
fn refresh_on_load_is_written_only_when_asked() {
    let cache = cache_with_shared(2);
    let plain = cache.to_xml_with_records(Some("rId1")).unwrap();
    assert_eq!(cache.write_xml(Some("rId1"), false).unwrap(), plain);
    let refreshing = String::from_utf8(cache.write_xml(Some("rId1"), true).unwrap()).unwrap();
    assert_eq!(
        refreshing,
        String::from_utf8(plain).unwrap().replacen(
            r#" r:id="rId1""#,
            r#" r:id="rId1" refreshOnLoad="1""#,
            1
        )
    );
    // Not modelled: a read drops it, as it did before.
    assert_eq!(
        PivotCacheDefinitionData::parse(refreshing.as_bytes()).unwrap(),
        cache
    );
}

/// A cache with `names.len()` fields named `names`, none with shared items.
fn cache_with_fields(names: &[&str]) -> PivotCacheDefinitionData {
    let mut cache = cache_with_shared(0);
    cache.fields = names
        .iter()
        .map(|&name| CacheFieldData {
            name: name.into(),
            ..CacheFieldData::default()
        })
        .collect();
    cache
}

#[test]
fn pivot_fields_are_extended_to_the_cache_field_count() {
    // A 5-field cache with fields 0 (row) and 3 (data) used: Excel 16 refuses
    // `<pivotFields count="4">` and opens and refreshes the table with a
    // fifth `<pivotField/>`.
    let cache = cache_with_fields(&["Region", "Product", "Year", "Revenue", "Units"]);
    let unused = PivotFieldData {
        data_field: false,
        ..bare_field(None, Vec::new(), None)
    };
    let mut table = one_field_table(bare_field(Some(PivotAxis::AxisRow), Vec::new(), None));
    table
        .pivot_fields
        .extend([unused.clone(), unused, bare_field(None, Vec::new(), None)]);
    table.data_fields[0].fld = 3;
    let before = table.clone();
    let mut added = 0;
    let xml = table.to_xml_with_cache(Some(&cache), &mut added).unwrap();
    assert_eq!(
        pivot_fields_xml(&xml),
        r#"<pivotFields count="5"><pivotField axis="axisRow"><items count="1"><item t="default"/></items></pivotField><pivotField/><pivotField/><pivotField dataField="1"/><pivotField/></pivotFields>"#
    );
    assert_eq!(added, 1, "the appended field is counted");
    assert_eq!(table, before, "the model is unchanged");
    // Read back, the table has a field per cache field and gains none.
    let back = PivotTableData::parse(&xml).unwrap();
    assert_eq!(back.pivot_fields.len(), 5);
    assert_eq!(back.to_xml_with_cache(Some(&cache), &mut 0).unwrap(), xml);
    // Without a cache, or with no more cache fields than pivot fields, the
    // pivot fields are written as modelled.
    for cache in [None, Some(&cache_with_fields(&["A", "B"]))] {
        let xml = table.to_xml_with_cache(cache, &mut 0).unwrap();
        assert!(
            pivot_fields_xml(&xml).starts_with(r#"<pivotFields count="4">"#),
            "{cache:?}"
        );
    }
    // A table without pivot fields gets one per cache field.
    let empty = PivotTableData::parse(&pivot_wrap("")).unwrap();
    let xml = empty.to_xml_with_cache(Some(&cache), &mut 0).unwrap();
    assert_eq!(
        pivot_fields_xml(&xml),
        format!(
            r#"<pivotFields count="5">{}</pivotFields>"#,
            "<pivotField/>".repeat(5)
        )
    );
}

#[test]
fn completed_items_are_capped_per_field_and_per_workbook() {
    let table = one_field_table(bare_field(Some(PivotAxis::AxisRow), Vec::new(), None));

    // One field past Excel's limit of unique items per field.
    let over = cache_with_shared(MAX_FIELD_ITEMS + 1);
    let mut added = 0;
    let err = table
        .to_xml_with_cache(Some(&over), &mut added)
        .unwrap_err();
    assert_eq!(err.code(), "SIZE_LIMIT");
    let msg = err.to_string();
    assert!(
        msg.contains("1048576") && msg.contains("explicit items"),
        "{msg}"
    );
    assert_eq!(added, 0);

    // Items the model gives are written as given and not counted.
    let mut given = table.clone();
    given.pivot_fields[0].items = vec![PivotItem {
        t: None,
        x: Some(0),
    }];
    let mut added = 0;
    given.to_xml_with_cache(Some(&over), &mut added).unwrap();
    assert_eq!(added, 0);
    drop(over);

    // The running total is checked before the table is written.
    let cache = cache_with_shared(2);
    let mut added = MAX_ADDED_ITEMS - 2;
    table.to_xml_with_cache(Some(&cache), &mut added).unwrap();
    assert_eq!(added, MAX_ADDED_ITEMS);
    let err = table
        .to_xml_with_cache(Some(&cache), &mut added)
        .unwrap_err();
    assert_eq!(err.code(), "SIZE_LIMIT");
    let msg = err.to_string();
    assert!(
        msg.contains("4194304") && msg.contains("explicit items"),
        "{msg}"
    );
    assert!(added > MAX_ADDED_ITEMS);
}

/// The cache definition Excel 16 saves for a pivot table inserted from the
/// table `Sales` (`PivotCaches().Create(1, "Sales")` through COM), with
/// `refreshedBy` anonymised.
const EXCEL_TABLE_SOURCE_CACHE: &str = r#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" r:id="rId1" refreshedBy="Author" refreshedDate="46338.560222453707" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="4" xr:uid="{61517C26-2271-41C6-BDD1-C7A40A20D77D}"><cacheSource type="worksheet"><worksheetSource name="Sales"/></cacheSource><cacheFields count="2"><cacheField name="Region" numFmtId="0"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField><cacheField name="Revenue" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="10" maxValue="40"/></cacheField></cacheFields></pivotCacheDefinition>"#;

/// The `<cacheSource>` Excel 16 saves for a pivot table over multiple
/// consolidation ranges (`PivotTableWizard(3, ...)` through COM).
const EXCEL_CONSOLIDATION_SOURCE: &str = r#"<cacheSource type="consolidation"><consolidation><pages count="1"><page count="2"><pageItem name="Item1"/><pageItem name="Item2"/></page></pages><rangeSets count="2"><rangeSet i1="0" ref="A1:B5" sheet="Q1"/><rangeSet i1="1" ref="A1:B5" sheet="Q2"/></rangeSets></consolidation></cacheSource>"#;

/// The `<cacheSource>` element of a written cache definition.
fn written_cache_source(def: &PivotCacheDefinitionData) -> String {
    let out = String::from_utf8(def.to_xml().unwrap()).unwrap();
    let start = out.find("<cacheSource").expect("a cacheSource");
    let end = out.find("<cacheFields").expect("cacheFields after it");
    assert!(
        !out.contains(r#"="""#),
        "an attribute is written empty: {out}"
    );
    out[start..end].to_owned()
}

#[test]
fn every_worksheet_source_is_written_as_read() {
    let w = |inner: &str| {
        format!(
            r#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cacheSource type="worksheet">{inner}</cacheSource><cacheFields count="0"/></pivotCacheDefinition>"#
        )
    };
    // (input <worksheetSource>, ref, sheet, name, written <worksheetSource>)
    let cases: [(&str, &str, &str, Option<&str>, &str); 5] = [
        (
            r#"<worksheetSource name="Sales"/>"#,
            "",
            "",
            Some("Sales"),
            r#"<worksheetSource name="Sales"/>"#,
        ),
        (
            r#"<worksheetSource ref="A1:B5" sheet="Data Sheet"/>"#,
            "A1:B5",
            "Data Sheet",
            None,
            r#"<worksheetSource ref="A1:B5" sheet="Data Sheet"/>"#,
        ),
        (
            r#"<worksheetSource name="Local" sheet="Data"/>"#,
            "",
            "Data",
            Some("Local"),
            r#"<worksheetSource name="Local" sheet="Data"/>"#,
        ),
        // A source in another workbook, as Excel 16 saves it: its r:id
        // names an externalLinkPath relationship in the definition's .rels,
        // which this part alone does not carry, so writing it would dangle
        // (Excel refuses that file) and it is dropped (see
        // a_worksheet_source_r_id_is_written_only_with_its_relationship).
        (
            r#"<worksheetSource ref="A1:B5" sheet="Data" r:id="rId2"/>"#,
            "A1:B5",
            "Data",
            None,
            r#"<worksheetSource ref="A1:B5" sheet="Data"/>"#,
        ),
        (
            r#"<worksheetSource name="Q&amp;A"/>"#,
            "",
            "",
            Some("Q&A"),
            r#"<worksheetSource name="Q&amp;A"/>"#,
        ),
    ];
    for (inner, ref_range, sheet, name, written) in cases {
        let def = PivotCacheDefinitionData::parse(w(inner).as_bytes()).unwrap();
        let src = &def.source;
        assert_eq!(src.ref_range, ref_range, "{inner}");
        assert_eq!(src.sheet, sheet, "{inner}");
        assert_eq!(src.name.as_deref(), name, "{inner}");
        assert_eq!(src.raw_xml, None, "{inner}");
        assert_eq!(
            written_cache_source(&def),
            format!(r#"<cacheSource type="worksheet">{written}</cacheSource>"#),
        );
        assert_eq!(
            PivotCacheDefinitionData::parse(&def.to_xml().unwrap()).unwrap(),
            PivotCacheDefinitionData {
                source: CacheSource {
                    r_id: None,
                    ..def.source.clone()
                },
                ..def
            }
        );
    }
}

#[test]
fn excel_table_source_survives_a_round_trip() {
    let def = PivotCacheDefinitionData::parse(EXCEL_TABLE_SOURCE_CACHE.as_bytes()).unwrap();
    assert_eq!(def.source.name.as_deref(), Some("Sales"));
    assert_eq!(
        written_cache_source(&def),
        r#"<cacheSource type="worksheet"><worksheetSource name="Sales"/></cacheSource>"#
    );
    // The source keys the JSON bridge carries: no new key for a range source.
    let json = serde_json::to_string(&def.source).unwrap();
    assert_eq!(json, r#"{"ref":"","sheet":"","name":"Sales"}"#);
    let range = PivotCacheDefinitionData::parse(EXCEL_TWO_AXES_CACHE.as_bytes()).unwrap();
    assert_eq!(
        serde_json::to_string(&range.source).unwrap(),
        r#"{"ref":"A1:E5","sheet":"Sheet1"}"#
    );
}

#[test]
fn a_range_source_is_written_byte_for_byte_as_before() {
    let def = PivotCacheDefinitionData {
        cache_id: None,
        source: CacheSource {
            ref_range: "A1:B3".into(),
            sheet: "Q&A".into(),
            ..Default::default()
        },
        fields: Vec::new(),
        record_count: Some(2),
        relationships: Vec::new(),
    };
    assert_eq!(
        String::from_utf8(def.to_xml_with_records(Some("rId1")).unwrap()).unwrap(),
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" recordCount="2"><cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="Q&amp;A"/></cacheSource><cacheFields count="0"/></pivotCacheDefinition>"#
    );
    // A cacheSource without `type` is a worksheet source (the schema default).
    let bare = PivotCacheDefinitionData::parse(
        br#"<pivotCacheDefinition><cacheSource><worksheetSource ref="A1:B3" sheet="S"/></cacheSource></pivotCacheDefinition>"#,
    )
    .unwrap();
    assert_eq!(bare.source.raw_xml, None);
    assert_eq!(
        written_cache_source(&bare),
        r#"<cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="S"/></cacheSource>"#
    );
    // A source with no value at all writes no attribute, never `ref=""`.
    let empty = PivotCacheDefinitionData {
        source: CacheSource {
            name: Some(String::new()),
            ..Default::default()
        },
        ..bare
    };
    assert_eq!(
        written_cache_source(&empty),
        r#"<cacheSource type="worksheet"><worksheetSource/></cacheSource>"#
    );
}

#[test]
fn a_source_that_is_not_a_worksheet_is_kept_verbatim() {
    let wrap = |root_ns: &str, source: &str| {
        format!(
            r#"<x:pivotCacheDefinition {root_ns}>{source}<x:cacheFields count="0"/></x:pivotCacheDefinition>"#
        )
    };
    let main = r#"xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#;
    // Excel's consolidation source with every element prefixed `x:`.
    let prefixed = EXCEL_CONSOLIDATION_SOURCE
        .replace('<', "<x:")
        .replace("<x:/", "</x:");
    // (source as read, as written)
    let cases = [
        (
            prefixed.clone(),
            prefixed.replacen("<x:cacheSource", &format!("<x:cacheSource {main}"), 1),
        ),
        (
            r#"<x:cacheSource type="external" connectionId="3"/>"#.to_owned(),
            r#"<x:cacheSource xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" type="external" connectionId="3"/>"#.to_owned(),
        ),
    ];
    for (read, written) in cases {
        let def = PivotCacheDefinitionData::parse(wrap(main, &read).as_bytes()).unwrap();
        assert_eq!(def.source.raw_xml.as_deref(), Some(written.as_str()));
        let out = String::from_utf8(def.to_xml().unwrap()).unwrap();
        assert!(out.contains(&format!("{written}<cacheFields")), "{out}");
        assert!(!out.contains("worksheetSource"), "{out}");
        assert_eq!(
            PivotCacheDefinitionData::parse(out.as_bytes()).unwrap(),
            def
        );
    }
    // Excel's own unprefixed consolidation source is written unchanged.
    let def = PivotCacheDefinitionData::parse(
        format!(
            r#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{EXCEL_CONSOLIDATION_SOURCE}<cacheFields count="0"/></pivotCacheDefinition>"#
        )
        .as_bytes(),
    )
    .unwrap();
    assert_eq!(written_cache_source(&def), EXCEL_CONSOLIDATION_SOURCE);
    // No prefix shifts the kept span: none, one UTF-8 BOM or two (quick-xml
    // skips the first BOM without counting it; the second is text before
    // the declaration).
    for prefix in ["", "\u{FEFF}", "\u{FEFF}\u{FEFF}"] {
        let bom = PivotCacheDefinitionData::parse(
            format!(
                "{prefix}<?xml version=\"1.0\" encoding=\"UTF-8\"?><pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">{EXCEL_CONSOLIDATION_SOURCE}<cacheFields count=\"0\"/></pivotCacheDefinition>"
            )
            .as_bytes(),
        )
        .unwrap();
        assert_eq!(
            bom.source.raw_xml.as_deref(),
            Some(EXCEL_CONSOLIDATION_SOURCE),
            "{} BOMs",
            prefix.chars().count()
        );
        assert_eq!(
            PivotCacheDefinitionData::parse(&bom.to_xml().unwrap()).unwrap(),
            bom
        );
    }
}

#[test]
fn a_kept_source_span_must_be_one_element() {
    // The span the parser hands over is exact; one that is not a whole
    // element is an error, never a shifted copy.
    let data = b"<a/> x<cacheSource/>";
    for (start, end) in [(0, 4), (5, 20), (6, 19)] {
        let mut st = CacheDefParser::new();
        let span = &data[start..end];
        let res = st.keep_raw_source(data, start, end);
        let whole = span.trim_ascii_start().starts_with(b"<") && span.ends_with(b">");
        assert_eq!(res.is_ok(), whole, "{:?}", std::str::from_utf8(span));
        if let Err(e) = res {
            assert!(matches!(e, ModernXlsxError::XmlParse(_)), "{e:?}");
        }
    }
}

/// The cache definition Excel 16 saved for a pivot table over a
/// consolidation range in another workbook (`PivotCaches().Create(3,
/// "[xcons-source.xlsx]Data!R1C1:R3C3")` through COM), `refreshedBy`
/// anonymised and `extLst` left out. Its `.rels` holds `rId1` to the records
/// and `rId2`, the `externalLinkPath` the `<rangeSet r:id>` names.
const EXCEL_EXTERNAL_CONSOLIDATION_CACHE: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" r:id="rId1" refreshedBy="Author" refreshedDate="46338.648185416663" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="4" xr:uid="{8BD3491D-3E7C-4E19-ACCD-60BE3C78527F}"><cacheSource type="consolidation"><consolidation autoPage="0"><rangeSets count="1"><rangeSet ref="A1:C3" sheet="Data" r:id="rId2"/></rangeSets></consolidation></cacheSource><cacheFields count="1"><cacheField name="Row" numFmtId="0"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>"#;

/// The source of [`EXCEL_EXTERNAL_CONSOLIDATION_CACHE`] as the model keeps it.
const EXTERNAL_CONSOLIDATION_SOURCE: &str = r#"<cacheSource type="consolidation"><consolidation autoPage="0"><rangeSets count="1"><rangeSet ref="A1:C3" sheet="Data" r:id="rId2"/></rangeSets></consolidation></cacheSource>"#;

/// The `externalLinkPath` relationship Excel writes beside a source in
/// another workbook.
fn external_link(id: &str) -> CacheRelationship {
    CacheRelationship {
        id: id.to_owned(),
        rel_type:
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
                .to_owned(),
        target: "xcons-source.xlsx".to_owned(),
        target_mode: Some("External".to_owned()),
    }
}

#[test]
fn a_kept_source_relies_on_the_r_prefix_the_root_declares_once() {
    let def =
        PivotCacheDefinitionData::parse(EXCEL_EXTERNAL_CONSOLIDATION_CACHE.as_bytes()).unwrap();
    // The root binds `r` to the relationships namespace, as the writer's
    // root always does: the kept span does not declare it again.
    assert_eq!(
        def.source.raw_xml.as_deref(),
        Some(EXTERNAL_CONSOLIDATION_SOURCE)
    );
    // A root that does not declare `r` reads the same, so a second read of
    // the written part (whose root declares it) cannot add a declaration.
    let bare = EXCEL_EXTERNAL_CONSOLIDATION_CACHE.replacen(
        r#" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
        "",
        1,
    );
    assert_eq!(
        PivotCacheDefinitionData::parse(bare.as_bytes()).unwrap(),
        def
    );
    let def = PivotCacheDefinitionData {
        relationships: vec![external_link("rId2")],
        ..def
    };
    let once = def.to_xml_with_records(Some("rId1")).unwrap();
    let text = String::from_utf8(once.clone()).unwrap();
    assert_eq!(text.matches("xmlns:r=").count(), 1, "{text}");
    assert!(
        text.contains(&format!("{EXTERNAL_CONSOLIDATION_SOURCE}<cacheFields")),
        "{text}"
    );
    let again = PivotCacheDefinitionData {
        relationships: vec![external_link("rId2")],
        ..PivotCacheDefinitionData::parse(&once).unwrap()
    };
    assert_eq!(again, def);
    assert_eq!(again.to_xml_with_records(Some("rId1")).unwrap(), once);
    // A prefix bound to another URI inside the span stays declared there.
    let other = EXCEL_EXTERNAL_CONSOLIDATION_CACHE
        .replace("<cacheSource ", r#"<cacheSource xmlns:r="urn:other" "#);
    let kept = PivotCacheDefinitionData::parse(other.as_bytes()).unwrap();
    assert!(
        kept.source
            .raw_xml
            .as_deref()
            .unwrap()
            .starts_with(r#"<cacheSource xmlns:r="urn:other" "#)
    );
}

#[test]
fn a_relationship_id_in_a_kept_source_is_written_only_when_kept() {
    let read =
        PivotCacheDefinitionData::parse(EXCEL_EXTERNAL_CONSOLIDATION_CACHE.as_bytes()).unwrap();
    // With the relationship it names, the source is written as read.
    let kept = PivotCacheDefinitionData {
        relationships: vec![external_link("rId2")],
        ..read.clone()
    };
    assert_eq!(written_cache_source(&kept), EXTERNAL_CONSOLIDATION_SOURCE);
    // Without it (a model built by hand), the r:id would dangle and Excel
    // refuses such a file: the attribute is dropped, nothing else changes.
    for relationships in [Vec::new(), vec![external_link("rId3")]] {
        let dangling = PivotCacheDefinitionData {
            relationships,
            ..read.clone()
        };
        assert_eq!(
            written_cache_source(&dangling),
            EXTERNAL_CONSOLIDATION_SOURCE.replace(r#" r:id="rId2""#, "")
        );
    }
    // Any prefix bound to the relationships namespace inside the span counts;
    // an `id` in another namespace is not a relationship id.
    let raw = r#"<cacheSource type="external" xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x q:id="rId7" a="&quot;'"/><y xmlns:r="urn:other" r:id="rId7"/><z id="rId7"/></cacheSource>"#;
    let def = PivotCacheDefinitionData {
        source: CacheSource {
            raw_xml: Some(raw.to_owned()),
            ..Default::default()
        },
        ..read
    };
    assert_eq!(
        written_cache_source(&def),
        r#"<cacheSource type="external" xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x a="&quot;'"/><y xmlns:r="urn:other" r:id="rId7"/><z id="rId7"/></cacheSource>"#
    );
}

#[test]
fn the_records_relationship_takes_an_id_no_kept_relationship_uses() {
    let def = |ids: &[&str]| PivotCacheDefinitionData {
        relationships: ids.iter().map(|id| external_link(id)).collect(),
        ..PivotCacheDefinitionData::parse(EXCEL_TWO_AXES_CACHE.as_bytes()).unwrap()
    };
    // rId1 when it is free, so a file without other relationships is written
    // as before.
    assert_eq!(def(&[]).records_r_id(), "rId1");
    assert_eq!(def(&["rId2"]).records_r_id(), "rId1");
    assert_eq!(def(&["rId1", "rId2"]).records_r_id(), "rId3");
    assert_eq!(def(&["rId1", "link"]).records_r_id(), "rId2");
}

#[test]
fn a_worksheet_source_r_id_is_written_only_with_its_relationship() {
    let xml = EXCEL_TWO_AXES_CACHE.replace(
        r#"<worksheetSource ref="A1:E5" sheet="Sheet1"/>"#,
        r#"<worksheetSource ref="A1:E5" sheet="Sheet1" r:id="rId2"/>"#,
    );
    let read = PivotCacheDefinitionData::parse(xml.as_bytes()).unwrap();
    assert_eq!(read.source.r_id.as_deref(), Some("rId2"));
    let kept = PivotCacheDefinitionData {
        relationships: vec![external_link("rId2")],
        ..read.clone()
    };
    assert_eq!(
        written_cache_source(&kept),
        r#"<cacheSource type="worksheet"><worksheetSource ref="A1:E5" sheet="Sheet1" r:id="rId2"/></cacheSource>"#
    );
    assert_eq!(
        PivotCacheDefinitionData {
            relationships: vec![external_link("rId2")],
            ..PivotCacheDefinitionData::parse(&kept.to_xml().unwrap()).unwrap()
        },
        kept
    );
    assert_eq!(
        written_cache_source(&read),
        r#"<cacheSource type="worksheet"><worksheetSource ref="A1:E5" sheet="Sheet1"/></cacheSource>"#
    );
    // The JSON bridge gains no key for a pivot without them.
    let plain = PivotCacheDefinitionData::parse(EXCEL_TWO_AXES_CACHE.as_bytes()).unwrap();
    let json = serde_json::to_string(&plain).unwrap();
    assert!(
        !json.contains("relationships") && !json.contains("rId"),
        "{json}"
    );
    let json = serde_json::to_value(&kept).unwrap();
    assert_eq!(json["source"]["rId"], "rId2");
    assert_eq!(
        json["relationships"],
        serde_json::json!([{
            "id": "rId2",
            "type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
            "target": "xcons-source.xlsx",
            "targetMode": "External",
        }])
    );
}
