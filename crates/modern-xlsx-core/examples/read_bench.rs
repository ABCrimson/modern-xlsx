//! Quick benchmark: generate a large XLSX, then read it back natively.
use std::time::Instant;

use modern_xlsx_core::ooxml::worksheet::{Cell, CellType, Row, WorksheetXml};
use modern_xlsx_core::{SheetData, WorkbookData};

fn main() {
    // Build a 100K-row workbook with 10 columns (mixed types).
    let mut rows = Vec::with_capacity(100_001);
    // Header row
    let header_cells: Vec<Cell> = (0..10)
        .map(|c| Cell {
            reference: format!("{}{}", (b'A' + c) as char, 1),
            cell_type: CellType::SharedString,
            style_index: None,
            value: Some(format!("Col{c}")),
            formula: None,
            formula_type: None,
            formula_ref: None,
            shared_index: None,
            inline_string: None,
            dynamic_array: None,
        })
        .collect();
    rows.push(Row {
        index: 1,
        cells: header_cells,
        height: None,
        hidden: false,
    });

    for r in 0..100_000u32 {
        let mut cells = Vec::with_capacity(10);
        for c in 0..10u8 {
            let col_letter = (b'A' + c) as char;
            let cell_ref = format!("{}{}", col_letter, r + 2);
            let (ct, val) = if c % 3 == 0 {
                (CellType::SharedString, format!("String value {r}-{c}"))
            } else if c % 3 == 1 {
                (CellType::Number, format!("{}", r * 1000 + c as u32))
            } else {
                (
                    CellType::Boolean,
                    if r % 2 == 0 {
                        "1".to_string()
                    } else {
                        "0".to_string()
                    },
                )
            };
            cells.push(Cell {
                reference: cell_ref,
                cell_type: ct,
                style_index: None,
                value: Some(val),
                formula: None,
                formula_type: None,
                formula_ref: None,
                shared_index: None,
                inline_string: None,
                dynamic_array: None,
            });
        }
        rows.push(Row {
            index: r + 2,
            cells,
            height: None,
            hidden: false,
        });
    }

    let ws = WorksheetXml {
        dimension: Some("A1:J100001".to_string()),
        rows,
        merge_cells: Vec::new(),
        auto_filter: None,
        frozen_pane: None,
        columns: Vec::new(),
        data_validations: Vec::new(),
        conditional_formatting: Vec::new(),
        hyperlinks: Vec::new(),
        page_setup: None,
        sheet_protection: None,
        comments: Vec::new(),
        tab_color: None,
    };
    let wb = WorkbookData {
        sheets: vec![SheetData {
            name: "Data".to_string(),
            worksheet: ws,
        }],
        date_system: modern_xlsx_core::dates::DateSystem::Date1900,
        styles: modern_xlsx_core::ooxml::styles::Styles::default_styles(),
        defined_names: Vec::new(),
        shared_strings: None,
        doc_properties: None,
        theme_colors: None,
        calc_chain: Vec::new(),
        workbook_views: Vec::new(),
        preserved_entries: Default::default(),
    };

    // Write to XLSX
    eprintln!("Writing 100K-row workbook...");
    let t0 = Instant::now();
    let xlsx_bytes = modern_xlsx_core::writer::write_xlsx(&wb).expect("write");
    let write_ms = t0.elapsed().as_millis();
    eprintln!("  Write: {write_ms}ms ({} bytes)", xlsx_bytes.len());

    // Read it back
    eprintln!("Reading back...");
    let t1 = Instant::now();
    let wb2 = modern_xlsx_core::reader::read_xlsx(&xlsx_bytes).expect("read");
    let read_ms = t1.elapsed().as_millis();
    eprintln!(
        "  Read: {read_ms}ms ({} rows)",
        wb2.sheets[0].worksheet.rows.len()
    );

    // JSON serialize (simulates the WASM path)
    eprintln!("JSON serialize...");
    let t2 = Instant::now();
    let json = serde_json::to_string(&wb2).expect("json");
    let json_ms = t2.elapsed().as_millis();
    eprintln!("  JSON serialize: {json_ms}ms ({} bytes)", json.len());

    // JSON parse back
    eprintln!("JSON parse...");
    let t3 = Instant::now();
    let _wb3: WorkbookData = serde_json::from_str(&json).expect("parse");
    let parse_ms = t3.elapsed().as_millis();
    eprintln!("  JSON parse: {parse_ms}ms");

    eprintln!("\nTotal (read + json_ser): {}ms", read_ms + json_ms);

    // Benchmark the streaming JSON path (read_xlsx_json).
    eprintln!("\n--- Streaming JSON path (read_xlsx_json) ---");
    let t4 = Instant::now();
    let json2 = modern_xlsx_core::reader::read_xlsx_json(&xlsx_bytes).expect("read_xlsx_json");
    let stream_ms = t4.elapsed().as_millis();
    eprintln!("  read_xlsx_json: {stream_ms}ms ({} bytes)", json2.len());

    // Verify the streaming JSON is valid by parsing it.
    let t5 = Instant::now();
    let _wb4: WorkbookData = serde_json::from_str(&json2).expect("parse streaming json");
    let parse2_ms = t5.elapsed().as_millis();
    eprintln!("  JSON parse (streaming): {parse2_ms}ms");
    eprintln!(
        "\nSpeedup: {:.1}x ({}ms vs {}ms)",
        (read_ms + json_ms) as f64 / stream_ms as f64,
        stream_ms,
        read_ms + json_ms
    );
}
