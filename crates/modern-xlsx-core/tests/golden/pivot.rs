//! Golden test for the pivot parts. It serialises only the subtree under test
//! (see `crate::common::golden_subtree`).

use modern_xlsx_core::ooxml::pivot_table::{
    CacheFieldData, CacheSource, CacheValue, PivotAxis, PivotCacheDefinitionData,
    PivotCacheRecordsData, PivotDataFieldData, PivotFieldData, PivotFieldRef, PivotItem,
    PivotLocation, PivotPageFieldData, PivotTableData, SharedItemsAttrs, SubtotalFunction,
};
use modern_xlsx_core::ooxml::worksheet::WorksheetXml;

use crate::common::{assert_same, golden_subtree, minimal_workbook, roundtrip, text_row};

/// A pivot table with a page field, a data field and subtotal flags, over a
/// cache whose names need escaping, written then read back.
#[test]
fn golden_pivot_roundtrip() {
    let cache = PivotCacheDefinitionData {
        cache_id: Some(3),
        source: CacheSource {
            ref_range: "A1:B3".into(),
            sheet: "Q&A".into(),
            ..Default::default()
        },
        fields: vec![
            CacheFieldData {
                name: "Region & Area".into(),
                shared_items: vec![
                    CacheValue::String { v: "East".into() },
                    CacheValue::String { v: "West".into() },
                ],
                ..Default::default()
            },
            CacheFieldData {
                name: "Amt".into(),
                num_fmt_id: Some(0),
                shared_items_attrs: Some(SharedItemsAttrs {
                    contains_semi_mixed_types: Some(false),
                    contains_string: Some(false),
                    contains_number: Some(true),
                    min_value: Some(1.0),
                    max_value: Some(9.0),
                    ..Default::default()
                }),
                ..Default::default()
            },
        ],
        record_count: Some(2),
        relationships: Vec::new(),
    };
    let records = PivotCacheRecordsData {
        records: vec![
            vec![CacheValue::Index { v: 0 }, CacheValue::Number { v: 1.0 }],
            vec![CacheValue::Index { v: 1 }, CacheValue::Number { v: 9.0 }],
        ],
    };
    let table = PivotTableData {
        name: "P & L".into(),
        data_caption: None,
        location: PivotLocation {
            ref_range: "D1:E4".into(),
            first_header_row: None,
            first_data_row: None,
            first_data_col: None,
        },
        pivot_fields: vec![
            PivotFieldData {
                axis: Some(PivotAxis::AxisRow),
                name: None,
                items: vec![
                    PivotItem {
                        t: None,
                        x: Some(0),
                    },
                    PivotItem {
                        t: None,
                        x: Some(1),
                    },
                    PivotItem {
                        t: Some("default".into()),
                        x: None,
                    },
                ],
                subtotals: vec![SubtotalFunction::Sum, SubtotalFunction::Count],
                compact: false,
                outline: true,
                data_field: false,
                default_subtotal: Some(false),
            },
            PivotFieldData {
                axis: None,
                name: None,
                items: Vec::new(),
                subtotals: Vec::new(),
                compact: true,
                outline: true,
                data_field: true,
                default_subtotal: None,
            },
        ],
        row_fields: vec![PivotFieldRef { x: 0 }],
        col_fields: Vec::new(),
        data_fields: vec![PivotDataFieldData {
            name: Some("Sum of Amt".into()),
            fld: 1,
            subtotal: SubtotalFunction::Sum,
            num_fmt_id: None,
        }],
        page_fields: vec![PivotPageFieldData {
            fld: 0,
            item: None,
            name: None,
        }],
        cache_id: 3,
    };

    let mut ws = WorksheetXml {
        rows: vec![text_row(1, "A1", "x")],
        ..Default::default()
    };
    ws.pivot_tables = vec![table.clone()];
    let mut wb = minimal_workbook("Sheet1", ws);
    wb.pivot_caches = vec![cache.clone()];
    wb.pivot_cache_records = vec![records.clone()];

    let wb2 = roundtrip(&wb);
    assert_same(
        &wb2.pivot_cache_records,
        &wb.pivot_cache_records,
        "pivot cache records",
    );
    golden_subtree(
        &(
            &wb2.sheets[0].worksheet.pivot_tables,
            &wb2.pivot_caches,
            &wb2.pivot_cache_records,
        ),
        "pivot_roundtrip",
    );
}
