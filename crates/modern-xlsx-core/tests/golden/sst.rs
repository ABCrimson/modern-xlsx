//! Golden test for the shared string table: rich-text runs with the full run
//! properties survive write -> read.

use modern_xlsx_core::ooxml::color::Color;
use modern_xlsx_core::ooxml::shared_strings::RichTextRun;
use modern_xlsx_core::ooxml::worksheet::{Cell, CellType, Row};

use crate::common::{golden_subtree, minimal_workbook, minimal_worksheet, roundtrip, text_cell};

#[test]
fn golden_sst_rich() {
    let rich = Cell {
        reference: "A1".into(),
        cell_type: CellType::SharedString,
        value: Some("x2 and more".into()),
        rich_text: Some(vec![
            RichTextRun {
                text: "x".into(),
                bold: Some(false),
                font_name: Some("Aptos".into()),
                font_size: Some(11.0),
                family: Some(2),
                charset: Some(204),
                scheme: Some("minor".into()),
                color_spec: Some(Color {
                    theme: Some(1),
                    ..Color::default()
                }),
                ..Default::default()
            },
            RichTextRun {
                text: "2".into(),
                vert_align: Some("superscript".into()),
                underline: Some(true),
                underline_style: Some("double".into()),
                color: Some("FFFF0000".into()),
                ..Default::default()
            },
            RichTextRun {
                text: " and more".into(),
                ..Default::default()
            },
        ]),
        ..Default::default()
    };
    let ws = minimal_worksheet(vec![Row {
        index: 1,
        cells: vec![rich, text_cell("B1", "plain")],
        ..Default::default()
    }]);
    let wb2 = roundtrip(&minimal_workbook("Sheet1", ws));
    golden_subtree(&wb2.shared_strings, "sst_rich");
}
