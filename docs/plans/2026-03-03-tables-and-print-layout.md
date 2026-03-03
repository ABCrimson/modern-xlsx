# Tables & Print Layout Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Implement Excel Tables (ListObjects) with full CRUD and roundtrip, plus complete Print & Page Layout features (headers/footers, print titles, print areas, row/column grouping).

**Architecture:** New Rust `ooxml/tables.rs` module for table XML parsing/writing, extensions to `worksheet.rs` for headers/footers and outline levels, extensions to `workbook.rs` for print titles/areas. TypeScript types and API methods mirror the Rust structs across the WASM JSON boundary.

**Tech Stack:** Rust 1.94 (quick-xml SAX, serde), TypeScript 6.0 (ESM), Vitest 4.1, wasm-pack 0.14

---

## Phase 1: Tables & Structured References (0.2.x)

### Task 1: Table Rust Types

**Files:**
- Create: `crates/modern-xlsx-core/src/ooxml/tables.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/mod.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

**Step 1: Create `tables.rs` with struct definitions**

```rust
// crates/modern-xlsx-core/src/ooxml/tables.rs
use serde::{Deserialize, Serialize};

/// An Excel Table (ListObject) definition from `xl/tables/table{n}.xml`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableDefinition {
    /// Unique table ID across the workbook.
    pub id: u32,
    /// Internal name (used by VBA).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Display name (must be unique, used in structured references).
    pub display_name: String,
    /// Cell range including header row, e.g. "A1:D10".
    #[serde(rename = "ref")]
    pub ref_range: String,
    /// Number of header rows (default 1, set 0 for no header).
    #[serde(default = "default_header_row_count")]
    pub header_row_count: u32,
    /// Number of totals rows (default 0).
    #[serde(default)]
    pub totals_row_count: u32,
    /// Whether the totals row is shown (default true).
    #[serde(default = "default_true")]
    pub totals_row_shown: bool,
    /// Table columns.
    pub columns: Vec<TableColumn>,
    /// Style information.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_info: Option<TableStyleInfo>,
    /// Table-scoped auto-filter range.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_filter_ref: Option<String>,
}

fn default_header_row_count() -> u32 { 1 }
fn default_true() -> bool { true }

/// A column within a table definition.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableColumn {
    /// Column ID (unique within the table).
    pub id: u32,
    /// Column name (matches header cell text).
    pub name: String,
    /// Totals row aggregate function.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_function: Option<String>,
    /// Totals row label text (when function is "none").
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_label: Option<String>,
    /// Calculated column formula (structured references).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calculated_column_formula: Option<String>,
    /// DXF ID for header row styling.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_row_dxf_id: Option<u32>,
    /// DXF ID for data area styling.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_dxf_id: Option<u32>,
    /// DXF ID for totals row styling.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_dxf_id: Option<u32>,
}

/// Table style info from `<tableStyleInfo>`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableStyleInfo {
    /// Built-in style name (e.g. "TableStyleMedium2").
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(default)]
    pub show_first_column: bool,
    #[serde(default)]
    pub show_last_column: bool,
    #[serde(default = "default_true")]
    pub show_row_stripes: bool,
    #[serde(default)]
    pub show_column_stripes: bool,
}
```

**Step 2: Register module in `ooxml/mod.rs`**

Add `pub mod tables;` to the module list in `crates/modern-xlsx-core/src/ooxml/mod.rs`.

**Step 3: Add `tables` field to `WorksheetXml`**

In `crates/modern-xlsx-core/src/ooxml/worksheet.rs`, add to the `WorksheetXml` struct:

```rust
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tables: Vec<super::tables::TableDefinition>,
```

**Step 4: Add TypeScript types**

In `packages/modern-xlsx/src/types.ts`, add:

```typescript
// ---------------------------------------------------------------------------
// Table definitions (Excel ListObjects)
// ---------------------------------------------------------------------------

export interface TableColumnData {
  id: number;
  name: string;
  totalsRowFunction?: string | null;
  totalsRowLabel?: string | null;
  calculatedColumnFormula?: string | null;
  headerRowDxfId?: number | null;
  dataDxfId?: number | null;
  totalsRowDxfId?: number | null;
}

export interface TableStyleInfoData {
  name?: string | null;
  showFirstColumn: boolean;
  showLastColumn: boolean;
  showRowStripes: boolean;
  showColumnStripes: boolean;
}

export interface TableDefinitionData {
  id: number;
  name?: string | null;
  displayName: string;
  ref: string;
  headerRowCount: number;
  totalsRowCount: number;
  totalsRowShown: boolean;
  columns: TableColumnData[];
  styleInfo?: TableStyleInfoData | null;
  autoFilterRef?: string | null;
}
```

Add to `WorksheetData`:
```typescript
  tables?: TableDefinitionData[];
```

**Step 5: Write Rust unit tests**

In `tables.rs`, add `#[cfg(test)]` module:

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_definition_serde_roundtrip() {
        let table = TableDefinition {
            id: 1,
            name: Some("Table1".into()),
            display_name: "Table1".into(),
            ref_range: "A1:C4".into(),
            header_row_count: 1,
            totals_row_count: 0,
            totals_row_shown: false,
            columns: vec![
                TableColumn { id: 1, name: "Name".into(), ..Default::default() },
                TableColumn { id: 2, name: "Age".into(), ..Default::default() },
                TableColumn { id: 3, name: "City".into(), ..Default::default() },
            ],
            style_info: Some(TableStyleInfo {
                name: Some("TableStyleMedium2".into()),
                show_first_column: false,
                show_last_column: false,
                show_row_stripes: true,
                show_column_stripes: false,
            }),
            auto_filter_ref: Some("A1:C4".into()),
        };
        let json = serde_json::to_string(&table).unwrap();
        let parsed: TableDefinition = serde_json::from_str(&json).unwrap();
        assert_eq!(parsed.id, 1);
        assert_eq!(parsed.display_name, "Table1");
        assert_eq!(parsed.columns.len(), 3);
    }
}
```

**Step 6: Implement `Default` for `TableColumn`**

```rust
impl Default for TableColumn {
    fn default() -> Self {
        Self {
            id: 0,
            name: String::new(),
            totals_row_function: None,
            totals_row_label: None,
            calculated_column_formula: None,
            header_row_dxf_id: None,
            data_dxf_id: None,
            totals_row_dxf_id: None,
        }
    }
}
```

**Step 7: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All existing tests pass + new `test_table_definition_serde_roundtrip` passes.

**Step 8: Commit**

```bash
git add crates/modern-xlsx-core/src/ooxml/tables.rs crates/modern-xlsx-core/src/ooxml/mod.rs crates/modern-xlsx-core/src/ooxml/worksheet.rs packages/modern-xlsx/src/types.ts
git commit -m "feat(tables): add Rust TableDefinition types and TypeScript TableDefinitionData"
```

---

### Task 2: Table XML Parser (Reader)

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/tables.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/content_types.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/relationships.rs`
- Modify: `crates/modern-xlsx-core/src/reader.rs`

**Step 1: Add content type and relationship constants**

In `content_types.rs`, add:
```rust
pub const CT_TABLE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
```

In `relationships.rs`, add:
```rust
pub const REL_TABLE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
```

**Step 2: Implement SAX parser in `tables.rs`**

```rust
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use super::push_entity;
use crate::{ModernXlsxError, Result};
use super::SPREADSHEET_NS;

impl TableDefinition {
    /// Parse a table definition from `xl/tables/table{n}.xml`.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::with_capacity(512);

        let mut id: u32 = 0;
        let mut name: Option<String> = None;
        let mut display_name = String::new();
        let mut ref_range = String::new();
        let mut header_row_count: u32 = 1;
        let mut totals_row_count: u32 = 0;
        let mut totals_row_shown: bool = true;
        let mut columns: Vec<TableColumn> = Vec::new();
        let mut style_info: Option<TableStyleInfo> = None;
        let mut auto_filter_ref: Option<String> = None;

        // Current tableColumn being built (for child elements).
        let mut current_col: Option<TableColumn> = None;
        let mut in_calculated_formula = false;
        let mut formula_buf = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let is_empty = matches!(reader.read_event_into(&mut []), _);
                    // Re-check: we handle Start and Empty together.
                    match e.local_name().as_ref() {
                        b"table" => {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"id" => id = std::str::from_utf8(&attr.value).unwrap_or("0").parse().unwrap_or(0),
                                    b"name" => name = Some(std::str::from_utf8(&attr.value).unwrap_or_default().to_owned()),
                                    b"displayName" => display_name = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    b"ref" => ref_range = std::str::from_utf8(&attr.value).unwrap_or_default().to_owned(),
                                    b"headerRowCount" => header_row_count = std::str::from_utf8(&attr.value).unwrap_or("1").parse().unwrap_or(1),
                                    b"totalsRowCount" => totals_row_count = std::str::from_utf8(&attr.value).unwrap_or("0").parse().unwrap_or(0),
                                    b"totalsRowShown" => totals_row_shown = std::str::from_utf8(&attr.value).unwrap_or("1") != "0",
                                    _ => {}
                                }
                            }
                        }
                        b"autoFilter" => { /* parse auto_filter_ref */ }
                        b"tableColumn" => { /* parse column attrs */ }
                        b"tableStyleInfo" => { /* parse style info */ }
                        b"calculatedColumnFormula" => { /* start formula capture */ }
                        _ => {}
                    }
                }
                Ok(Event::End(ref e)) => { /* close elements, push columns */ }
                Ok(Event::Text(ref e)) => { /* capture formula text */ }
                Ok(Event::Eof) => break,
                _ => {}
            }
            buf.clear();
        }
        // ... build and return TableDefinition
    }
}
```

The actual parser should follow the same patterns as `WorksheetXml::parse()` — SAX-style with `push_entity()` for text content.

**Step 3: Wire table reading into `reader.rs`**

In `reader.rs`, after parsing each worksheet's `.rels` file (where comments are resolved), also resolve table relationships:

```rust
// In parse_common() or the sheet-loading loop:
for rel in ws_rels.find_by_type(REL_TABLE) {
    let table_path = resolve_relative_path(&ws_dir, &rel.target);
    if let Some(table_data) = entries.get(&table_path) {
        let table = TableDefinition::parse(table_data)?;
        worksheet.tables.push(table);
    }
}
```

Add table paths to the known/excluded paths so they aren't preserved as opaque blobs.

**Step 4: Write Rust parser tests**

```rust
#[test]
fn test_parse_table_xml() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="1" name="Table1" displayName="Table1"
       ref="A1:C4" totalsRowShown="0">
  <autoFilter ref="A1:C4"/>
  <tableColumns count="3">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Age"/>
    <tableColumn id="3" name="City"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2"
    showFirstColumn="0" showLastColumn="0"
    showRowStripes="1" showColumnStripes="0"/>
</table>"#;
    let table = TableDefinition::parse(xml).unwrap();
    assert_eq!(table.id, 1);
    assert_eq!(table.display_name, "Table1");
    assert_eq!(table.ref_range, "A1:C4");
    assert_eq!(table.columns.len(), 3);
    assert_eq!(table.columns[0].name, "Name");
    assert!(table.style_info.is_some());
    let si = table.style_info.unwrap();
    assert_eq!(si.name.as_deref(), Some("TableStyleMedium2"));
    assert!(si.show_row_stripes);
    assert!(!si.show_column_stripes);
}

#[test]
fn test_parse_table_with_totals_and_formulas() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="2" displayName="Sales" ref="A1:D6" totalsRowCount="1">
  <autoFilter ref="A1:D5"/>
  <tableColumns count="4">
    <tableColumn id="1" name="Product"/>
    <tableColumn id="2" name="Qty" totalsRowFunction="sum"/>
    <tableColumn id="3" name="Price" totalsRowFunction="average"/>
    <tableColumn id="4" name="Total" totalsRowFunction="sum">
      <calculatedColumnFormula>Sales[@Qty]*Sales[@Price]</calculatedColumnFormula>
    </tableColumn>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showRowStripes="1"/>
</table>"#;
    let table = TableDefinition::parse(xml).unwrap();
    assert_eq!(table.totals_row_count, 1);
    assert_eq!(table.columns[1].totals_row_function.as_deref(), Some("sum"));
    assert_eq!(table.columns[3].calculated_column_formula.as_deref(), Some("Sales[@Qty]*Sales[@Price]"));
}
```

**Step 5: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All pass.

**Step 6: Commit**

```bash
git add -A
git commit -m "feat(tables): SAX parser for table XML, reader wiring, relationship constants"
```

---

### Task 3: Table XML Writer

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/tables.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`
- Modify: `crates/modern-xlsx-core/src/writer.rs`

**Step 1: Implement `to_xml()` on `TableDefinition`**

```rust
impl TableDefinition {
    /// Serialize to XML bytes for `xl/tables/table{n}.xml`.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf = Vec::with_capacity(1024);
        let mut writer = Writer::new_with_indent(&mut buf, b' ', 2);

        writer.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

        let mut table_elem = BytesStart::new("table");
        table_elem.push_attribute(("xmlns", SPREADSHEET_NS));
        table_elem.push_attribute(("id", itoa::Buffer::new().format(self.id)));
        if let Some(ref n) = self.name {
            table_elem.push_attribute(("name", n.as_str()));
        }
        table_elem.push_attribute(("displayName", self.display_name.as_str()));
        table_elem.push_attribute(("ref", self.ref_range.as_str()));
        if self.header_row_count != 1 {
            table_elem.push_attribute(("headerRowCount", itoa::Buffer::new().format(self.header_row_count)));
        }
        if self.totals_row_count > 0 {
            table_elem.push_attribute(("totalsRowCount", itoa::Buffer::new().format(self.totals_row_count)));
        }
        if !self.totals_row_shown {
            table_elem.push_attribute(("totalsRowShown", "0"));
        }
        writer.write_event(Event::Start(table_elem))?;

        // <autoFilter>
        if let Some(ref af_ref) = self.auto_filter_ref {
            let mut af = BytesStart::new("autoFilter");
            af.push_attribute(("ref", af_ref.as_str()));
            writer.write_event(Event::Empty(af))?;
        }

        // <tableColumns>
        let mut tc_elem = BytesStart::new("tableColumns");
        tc_elem.push_attribute(("count", itoa::Buffer::new().format(self.columns.len())));
        writer.write_event(Event::Start(tc_elem))?;

        for col in &self.columns {
            if col.calculated_column_formula.is_some() {
                let mut c = BytesStart::new("tableColumn");
                c.push_attribute(("id", itoa::Buffer::new().format(col.id)));
                c.push_attribute(("name", col.name.as_str()));
                // ... add totalsRowFunction, DXF IDs if present
                writer.write_event(Event::Start(c))?;
                if let Some(ref formula) = col.calculated_column_formula {
                    writer.write_event(Event::Start(BytesStart::new("calculatedColumnFormula")))?;
                    writer.write_event(Event::Text(BytesText::new(formula)))?;
                    writer.write_event(Event::End(BytesEnd::new("calculatedColumnFormula")))?;
                }
                writer.write_event(Event::End(BytesEnd::new("tableColumn")))?;
            } else {
                let mut c = BytesStart::new("tableColumn");
                c.push_attribute(("id", itoa::Buffer::new().format(col.id)));
                c.push_attribute(("name", col.name.as_str()));
                // ... add optional attrs
                writer.write_event(Event::Empty(c))?;
            }
        }

        writer.write_event(Event::End(BytesEnd::new("tableColumns")))?;

        // <tableStyleInfo>
        if let Some(ref si) = self.style_info {
            let mut s = BytesStart::new("tableStyleInfo");
            if let Some(ref sn) = si.name {
                s.push_attribute(("name", sn.as_str()));
            }
            s.push_attribute(("showFirstColumn", if si.show_first_column { "1" } else { "0" }));
            s.push_attribute(("showLastColumn", if si.show_last_column { "1" } else { "0" }));
            s.push_attribute(("showRowStripes", if si.show_row_stripes { "1" } else { "0" }));
            s.push_attribute(("showColumnStripes", if si.show_column_stripes { "1" } else { "0" }));
            writer.write_event(Event::Empty(s))?;
        }

        writer.write_event(Event::End(BytesEnd::new("table")))?;
        Ok(buf)
    }
}
```

**Step 2: Add `<tableParts>` to worksheet writer**

In `worksheet.rs`, at the end of `to_xml_with_sst()`, before closing `</worksheet>`:

```rust
// <tableParts> — only if tables are present.
if !self.tables.is_empty() {
    let mut tp = BytesStart::new("tableParts");
    tp.push_attribute(("count", itoa::Buffer::new().format(self.tables.len())));
    writer.write_event(Event::Start(tp))?;
    for (i, _) in self.tables.iter().enumerate() {
        let mut part = BytesStart::new("tablePart");
        // rId will be set by the writer based on worksheet rels.
        // For now, use a placeholder that the writer replaces.
        part.push_attribute(("r:id", format!("rIdTable{}", i + 1).as_str()));
        writer.write_event(Event::Empty(part))?;
    }
    writer.write_event(Event::End(BytesEnd::new("tableParts")))?;
}
```

**Step 3: Wire table writing into `writer.rs`**

In `writer.rs`, within the sheet iteration loop (after comments handling):

```rust
// Write tables if the worksheet has any.
if !sheet.worksheet.tables.is_empty() {
    for (t_idx, table) in sheet.worksheet.tables.iter().enumerate() {
        let table_num = /* global table counter */;
        let table_path = format!("xl/tables/table{table_num}.xml");
        let table_xml = table.to_xml()?;

        content_types.add_override(format!("/{table_path}"), CT_TABLE);

        entries.push(ZipEntry {
            name: table_path,
            data: table_xml,
        });

        // Add table relationship to worksheet .rels
        ws_rels.add(
            format!("rIdTable{}", t_idx + 1),
            REL_TABLE,
            format!("../tables/table{table_num}.xml"),
        );
    }
}
```

**Step 4: Write roundtrip test**

```rust
#[test]
fn test_table_xml_roundtrip() {
    let table = TableDefinition { /* ... */ };
    let xml = table.to_xml().unwrap();
    let parsed = TableDefinition::parse(&xml).unwrap();
    assert_eq!(parsed.id, table.id);
    assert_eq!(parsed.display_name, table.display_name);
    assert_eq!(parsed.columns.len(), table.columns.len());
}
```

**Step 5: Run tests**

Run: `cargo test -p modern-xlsx-core`

**Step 6: Commit**

```bash
git add -A
git commit -m "feat(tables): table XML writer, tableParts in worksheet, writer pipeline integration"
```

---

### Task 4: TypeScript Table Read/Write API

**Files:**
- Modify: `packages/modern-xlsx/src/workbook.ts`
- Modify: `packages/modern-xlsx/src/index.ts`
- Create: `packages/modern-xlsx/__tests__/tables.test.ts`

**Step 1: Add `tables` getter to Worksheet class**

```typescript
get tables(): readonly TableDefinitionData[] {
  return this.sheetData.worksheet.tables ?? [];
}

getTable(displayName: string): TableDefinitionData | undefined {
  return this.tables.find((t) => t.displayName === displayName);
}
```

**Step 2: Add `addTable()` method to Worksheet class**

```typescript
addTable(
  ref: string,
  columns: { name: string; totalsRowFunction?: string; calculatedColumnFormula?: string }[],
  options?: {
    displayName?: string;
    styleName?: string;
    showHeaderRow?: boolean;
    showTotalsRow?: boolean;
    showRowStripes?: boolean;
    showColumnStripes?: boolean;
  },
): TableDefinitionData {
  const tables = (this.sheetData.worksheet.tables ??= []);

  // Auto-generate unique table ID (max existing + 1, across workbook via parent ref).
  const existingIds = tables.map((t) => t.id);
  const nextId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 1;

  const displayName = options?.displayName ?? `Table${nextId}`;
  const headerRowCount = options?.showHeaderRow === false ? 0 : 1;
  const totalsRowCount = options?.showTotalsRow ? 1 : 0;

  const table: TableDefinitionData = {
    id: nextId,
    displayName,
    ref: ref,
    headerRowCount,
    totalsRowCount,
    totalsRowShown: totalsRowCount > 0,
    columns: columns.map((c, i) => ({
      id: i + 1,
      name: c.name,
      totalsRowFunction: c.totalsRowFunction ?? null,
      calculatedColumnFormula: c.calculatedColumnFormula ?? null,
    })),
    styleInfo: {
      name: options?.styleName ?? 'TableStyleMedium2',
      showFirstColumn: false,
      showLastColumn: false,
      showRowStripes: options?.showRowStripes ?? true,
      showColumnStripes: options?.showColumnStripes ?? false,
    },
    autoFilterRef: ref,
  };

  tables.push(table);
  return table;
}

removeTable(displayName: string): boolean {
  const tables = this.sheetData.worksheet.tables;
  if (!tables) return false;
  const idx = tables.findIndex((t) => t.displayName === displayName);
  if (idx === -1) return false;
  tables.splice(idx, 1);
  return true;
}
```

**Step 3: Export new types from `index.ts`**

Add to `index.ts` exports:
```typescript
export type { TableDefinitionData, TableColumnData, TableStyleInfoData } from './types.js';
```

**Step 4: Write tests**

```typescript
// packages/modern-xlsx/__tests__/tables.test.ts
import { describe, it, expect, beforeAll } from 'vitest';
import { initWasm, Workbook, readBuffer } from '../src/index.js';

beforeAll(async () => { await initWasm(); });

describe('Excel Tables', () => {
  it('should add and roundtrip a basic table', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    // Populate header + data cells
    ws.cell('A1').value = 'Name';
    ws.cell('B1').value = 'Age';
    ws.cell('A2').value = 'Alice';
    ws.cell('B2').value = 30;

    ws.addTable('A1:B2', [{ name: 'Name' }, { name: 'Age' }], {
      displayName: 'People',
      styleName: 'TableStyleMedium2',
    });

    expect(ws.tables.length).toBe(1);
    expect(ws.tables[0].displayName).toBe('People');

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Data')!;
    expect(ws2.tables.length).toBe(1);
    expect(ws2.tables[0].displayName).toBe('People');
    expect(ws2.tables[0].columns.length).toBe(2);
  });

  it('should roundtrip table with totals and formulas', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sales');
    ws.addTable('A1:C5', [
      { name: 'Product' },
      { name: 'Qty', totalsRowFunction: 'sum' },
      { name: 'Total', totalsRowFunction: 'sum', calculatedColumnFormula: 'Sales[@Qty]*10' },
    ], { displayName: 'Sales', showTotalsRow: true });

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const table = wb2.getSheet('Sales')!.getTable('Sales')!;
    expect(table.totalsRowCount).toBe(1);
    expect(table.columns[1].totalsRowFunction).toBe('sum');
    expect(table.columns[2].calculatedColumnFormula).toBe('Sales[@Qty]*10');
  });

  it('should remove a table', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addTable('A1:B2', [{ name: 'A' }, { name: 'B' }]);
    expect(ws.tables.length).toBe(1);
    ws.removeTable('Table1');
    expect(ws.tables.length).toBe(0);
  });

  it('should support multiple tables per sheet', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Multi');
    ws.addTable('A1:B3', [{ name: 'X' }, { name: 'Y' }], { displayName: 'T1' });
    ws.addTable('D1:E3', [{ name: 'P' }, { name: 'Q' }], { displayName: 'T2' });
    expect(ws.tables.length).toBe(2);

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.getSheet('Multi')!.tables.length).toBe(2);
  });
});
```

**Step 5: Build WASM and run all tests**

Run:
```bash
cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm --no-opt
pnpm -C packages/modern-xlsx test
```

**Step 6: Commit**

```bash
git add -A
git commit -m "feat(tables): TypeScript Table API — addTable, removeTable, getTable, roundtrip tests"
```

---

### Task 5: Built-in Table Style Constants

**Files:**
- Create: `packages/modern-xlsx/src/table-styles.ts`
- Modify: `packages/modern-xlsx/src/index.ts`

**Step 1: Create style constants**

```typescript
// packages/modern-xlsx/src/table-styles.ts

/** All 60 built-in Excel table style names. */
export const TABLE_STYLES = {
  light: Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`) as string[],
  medium: Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`) as string[],
  dark: Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`) as string[],
} as const;

/** Set of all valid built-in table style names. */
export const VALID_TABLE_STYLES = new Set([
  ...TABLE_STYLES.light,
  ...TABLE_STYLES.medium,
  ...TABLE_STYLES.dark,
]);

/** Totals row aggregate functions. */
export type TotalsRowFunction =
  | 'none' | 'sum' | 'min' | 'max' | 'average'
  | 'count' | 'countNums' | 'stdDev' | 'var' | 'custom';
```

**Step 2: Export from index.ts**

```typescript
export { TABLE_STYLES, VALID_TABLE_STYLES } from './table-styles.js';
export type { TotalsRowFunction } from './table-styles.js';
```

**Step 3: Add validation in `addTable()`**

In `workbook.ts`, inside `addTable()`:
```typescript
if (options?.styleName && !VALID_TABLE_STYLES.has(options.styleName)) {
  console.warn(`Unknown table style: ${options.styleName}`);
}
```

**Step 4: Write test**

```typescript
it('should accept all built-in style names', () => {
  expect(VALID_TABLE_STYLES.size).toBe(60);
  expect(VALID_TABLE_STYLES.has('TableStyleMedium2')).toBe(true);
  expect(VALID_TABLE_STYLES.has('TableStyleLight21')).toBe(true);
  expect(VALID_TABLE_STYLES.has('TableStyleDark11')).toBe(true);
  expect(VALID_TABLE_STYLES.has('FakeStyle')).toBe(false);
});
```

**Step 5: Run tests and commit**

```bash
pnpm -C packages/modern-xlsx test
git add -A
git commit -m "feat(tables): built-in table style constants and validation"
```

---

### Task 6: Table Utilities (tableToJson, jsonToTable)

**Files:**
- Modify: `packages/modern-xlsx/src/utils.ts`
- Modify: `packages/modern-xlsx/src/index.ts`

**Step 1: Implement `tableToJson()`**

```typescript
export function tableToJson(
  ws: { rows: readonly RowData[]; tables?: readonly TableDefinitionData[] },
  tableName: string,
): Record<string, unknown>[] {
  const table = ws.tables?.find((t) => t.displayName === tableName);
  if (!table) throw new Error(`Table "${tableName}" not found`);
  // Decode table ref, extract header names and data rows from ws.rows
  // Return array of objects keyed by column names
}
```

**Step 2: Write tests for table utilities**

**Step 3: Run tests and commit**

```bash
git add -A
git commit -m "feat(tables): tableToJson and jsonToTable utility functions"
```

---

## Phase 2: Print & Page Layout (0.3.x)

### Task 7: Headers & Footers — Rust Types + Parser

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`
- Modify: `packages/modern-xlsx/src/types.ts`

**Step 1: Add Rust struct**

In `worksheet.rs`, add:

```rust
/// Header/footer configuration from `<headerFooter>`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooter {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub odd_header: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub odd_footer: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_header: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_footer: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_header: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_footer: Option<String>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub different_odd_even: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub different_first: bool,
    #[serde(default = "default_true", skip_serializing_if = "is_true")]
    pub scale_with_doc: bool,
    #[serde(default = "default_true", skip_serializing_if = "is_true")]
    pub align_with_margins: bool,
}

fn default_true() -> bool { true }
fn is_true(v: &bool) -> bool { *v }
```

Add to `WorksheetXml`:
```rust
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer: Option<HeaderFooter>,
```

**Step 2: Add parser in `WorksheetXml::parse()`**

Add a new `ParseState::HeaderFooter` variant. When encountering `<headerFooter>`, parse its attributes (`differentOddEven`, `differentFirst`, `scaleWithDoc`, `alignWithMargins`). For child elements (`oddHeader`, `oddFooter`, etc.), capture text content via `push_entity()` into the appropriate field.

```rust
// In the ParseState enum:
HeaderFooter,
HeaderFooterChild, // tracks which child (oddHeader, oddFooter, etc.)

// In the match block:
(ParseState::Root, b"headerFooter") => {
    let mut hf = HeaderFooter {
        odd_header: None, odd_footer: None,
        even_header: None, even_footer: None,
        first_header: None, first_footer: None,
        different_odd_even: false, different_first: false,
        scale_with_doc: true, align_with_margins: true,
    };
    for attr in e.attributes().flatten() {
        match attr.key.as_ref() {
            b"differentOddEven" => hf.different_odd_even = std::str::from_utf8(&attr.value).unwrap_or("0") == "1",
            b"differentFirst" => hf.different_first = std::str::from_utf8(&attr.value).unwrap_or("0") == "1",
            b"scaleWithDoc" => hf.scale_with_doc = std::str::from_utf8(&attr.value).unwrap_or("1") != "0",
            b"alignWithMargins" => hf.align_with_margins = std::str::from_utf8(&attr.value).unwrap_or("1") != "0",
            _ => {}
        }
    }
    header_footer = Some(hf);
    state = ParseState::HeaderFooter;
}
```

**Step 3: Add TypeScript type**

In `types.ts`:
```typescript
export interface HeaderFooterData {
  oddHeader?: string | null;
  oddFooter?: string | null;
  evenHeader?: string | null;
  evenFooter?: string | null;
  firstHeader?: string | null;
  firstFooter?: string | null;
  differentOddEven?: boolean;
  differentFirst?: boolean;
  scaleWithDoc?: boolean;
  alignWithMargins?: boolean;
}
```

Add to `WorksheetData`:
```typescript
  headerFooter?: HeaderFooterData | null;
```

**Step 4: Write Rust tests**

```rust
#[test]
fn test_parse_header_footer() {
    let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <headerFooter>
    <oddHeader>&amp;CConfidential</oddHeader>
    <oddFooter>&amp;LPage &amp;P&amp;R&amp;D</oddFooter>
  </headerFooter>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    let hf = ws.header_footer.as_ref().unwrap();
    assert_eq!(hf.odd_header.as_deref(), Some("&CConfidential"));
    assert_eq!(hf.odd_footer.as_deref(), Some("&LPage &P&R&D"));
}
```

**Step 5: Run tests and commit**

```bash
cargo test -p modern-xlsx-core
git add -A
git commit -m "feat(print): HeaderFooter Rust types and SAX parser"
```

---

### Task 8: Headers & Footers — Writer + TypeScript API

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`
- Modify: `packages/modern-xlsx/src/workbook.ts`

**Step 1: Add writer in `to_xml_with_sst()`**

After `<pageSetup>` and before closing `</worksheet>`:

```rust
// <headerFooter> — only if present.
if let Some(ref hf) = self.header_footer {
    let mut elem = BytesStart::new("headerFooter");
    if hf.different_odd_even {
        elem.push_attribute(("differentOddEven", "1"));
    }
    if hf.different_first {
        elem.push_attribute(("differentFirst", "1"));
    }
    if !hf.scale_with_doc {
        elem.push_attribute(("scaleWithDoc", "0"));
    }
    if !hf.align_with_margins {
        elem.push_attribute(("alignWithMargins", "0"));
    }
    writer.write_event(Event::Start(elem))?;

    fn write_hf_child(writer: &mut Writer<&mut Vec<u8>>, tag: &str, val: &Option<String>) -> Result<()> {
        if let Some(ref text) = val {
            writer.write_event(Event::Start(BytesStart::new(tag)))?;
            writer.write_event(Event::Text(BytesText::new(text)))?;
            writer.write_event(Event::End(BytesEnd::new(tag)))?;
        }
        Ok(())
    }
    write_hf_child(&mut writer, "oddHeader", &hf.odd_header)?;
    write_hf_child(&mut writer, "oddFooter", &hf.odd_footer)?;
    write_hf_child(&mut writer, "evenHeader", &hf.even_header)?;
    write_hf_child(&mut writer, "evenFooter", &hf.even_footer)?;
    write_hf_child(&mut writer, "firstHeader", &hf.first_header)?;
    write_hf_child(&mut writer, "firstFooter", &hf.first_footer)?;

    writer.write_event(Event::End(BytesEnd::new("headerFooter")))?;
}
```

**Step 2: Add TypeScript getter/setter on Worksheet**

```typescript
get headerFooter(): HeaderFooterData | null {
  return this.sheetData.worksheet.headerFooter ?? null;
}

set headerFooter(hf: HeaderFooterData | null) {
  this.sheetData.worksheet.headerFooter = hf;
}
```

**Step 3: Write roundtrip test**

```rust
#[test]
fn test_header_footer_roundtrip() {
    let mut ws = WorksheetXml { /* minimal */ };
    ws.header_footer = Some(HeaderFooter {
        odd_header: Some("&CPage &P of &N".into()),
        odd_footer: Some("&L&D&R&F".into()),
        ..Default::default()
    });
    let xml = ws.to_xml_with_sst(None).unwrap();
    let ws2 = WorksheetXml::parse(&xml).unwrap();
    assert_eq!(ws2.header_footer.as_ref().unwrap().odd_header.as_deref(), Some("&CPage &P of &N"));
}
```

**Step 4: Run tests and commit**

```bash
cargo test -p modern-xlsx-core
git add -A
git commit -m "feat(print): HeaderFooter writer and TypeScript API"
```

---

### Task 9: HeaderFooterBuilder (TypeScript Helper)

**Files:**
- Create: `packages/modern-xlsx/src/header-footer.ts`
- Modify: `packages/modern-xlsx/src/index.ts`

**Step 1: Create builder**

```typescript
// packages/modern-xlsx/src/header-footer.ts

/** Fluent builder for Excel header/footer format strings. */
export class HeaderFooterBuilder {
  private parts: string[] = [];

  left(text: string): this { this.parts.push(`&L${text}`); return this; }
  center(text: string): this { this.parts.push(`&C${text}`); return this; }
  right(text: string): this { this.parts.push(`&R${text}`); return this; }

  pageNumber(): string { return '&P'; }
  totalPages(): string { return '&N'; }
  date(): string { return '&D'; }
  time(): string { return '&T'; }
  fileName(): string { return '&F'; }
  sheetName(): string { return '&A'; }
  filePath(): string { return '&Z'; }

  bold(text: string): string { return `&B${text}&B`; }
  italic(text: string): string { return `&I${text}&I`; }
  underline(text: string): string { return `&U${text}&U`; }
  fontSize(size: number): string { return `&${size}`; }
  fontName(name: string): string { return `&"${name}"`; }

  build(): string { return this.parts.join(''); }
}
```

**Step 2: Export and test**

**Step 3: Commit**

```bash
git add -A
git commit -m "feat(print): HeaderFooterBuilder fluent helper"
```

---

### Task 10: Print Titles & Print Areas

**Files:**
- Modify: `packages/modern-xlsx/src/workbook.ts`
- Modify: `packages/modern-xlsx/__tests__/print.test.ts`

Print titles and print areas are stored as special `definedName` entries in the workbook (`_xlnm.Print_Titles`, `_xlnm.Print_Area`). The Rust side already roundtrips definedNames, so we only need TypeScript convenience methods.

**Step 1: Add `printTitles` getter/setter to Worksheet**

```typescript
get printTitles(): { rows?: [number, number]; columns?: [number, number] } | null {
  const sheetIndex = this.workbookRef.data.sheets.indexOf(this.sheetData);
  const dn = this.workbookRef.data.definedNames?.find(
    (d) => d.name === '_xlnm.Print_Titles' && d.sheetId === sheetIndex,
  );
  if (!dn) return null;
  // Parse value like "Sheet1!$1:$3,Sheet1!$A:$B"
  // Return { rows: [0, 2], columns: [0, 1] } (0-based)
}

set printTitles(titles: { rows?: [number, number]; columns?: [number, number] } | null) {
  const sheetIndex = this.workbookRef.data.sheets.indexOf(this.sheetData);
  const names = (this.workbookRef.data.definedNames ??= []);
  const idx = names.findIndex((d) => d.name === '_xlnm.Print_Titles' && d.sheetId === sheetIndex);

  if (!titles) {
    if (idx !== -1) names.splice(idx, 1);
    return;
  }

  const sheetName = this.name.includes(' ') ? `'${this.name}'` : this.name;
  const parts: string[] = [];
  if (titles.rows) {
    parts.push(`${sheetName}!$${titles.rows[0] + 1}:$${titles.rows[1] + 1}`);
  }
  if (titles.columns) {
    parts.push(`${sheetName}!$${columnToLetter(titles.columns[0])}:$${columnToLetter(titles.columns[1])}`);
  }
  const value = parts.join(',');

  if (idx !== -1) {
    names[idx].value = value;
  } else {
    names.push({ name: '_xlnm.Print_Titles', value, sheetId: sheetIndex });
  }
}
```

**Step 2: Add `printArea` getter/setter to Worksheet**

```typescript
get printArea(): string | null {
  const sheetIndex = this.workbookRef.data.sheets.indexOf(this.sheetData);
  const dn = this.workbookRef.data.definedNames?.find(
    (d) => d.name === '_xlnm.Print_Area' && d.sheetId === sheetIndex,
  );
  return dn?.value ?? null;
}

set printArea(area: string | null) {
  const sheetIndex = this.workbookRef.data.sheets.indexOf(this.sheetData);
  const names = (this.workbookRef.data.definedNames ??= []);
  const idx = names.findIndex((d) => d.name === '_xlnm.Print_Area' && d.sheetId === sheetIndex);

  if (!area) {
    if (idx !== -1) names.splice(idx, 1);
    return;
  }

  // Auto-prefix with sheet name if not already present
  const sheetName = this.name.includes(' ') ? `'${this.name}'` : this.name;
  const value = area.includes('!') ? area : `${sheetName}!${area}`;

  if (idx !== -1) {
    names[idx].value = value;
  } else {
    names.push({ name: '_xlnm.Print_Area', value, sheetId: sheetIndex });
  }
}
```

**Step 3: Write tests**

```typescript
describe('Print Titles & Areas', () => {
  it('should roundtrip print titles (rows)', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Report');
    ws.printTitles = { rows: [0, 1] }; // Repeat rows 1-2

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Report')!;
    expect(ws2.printTitles).toEqual({ rows: [0, 1] });
  });

  it('should roundtrip print area', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.printArea = '$A$1:$H$50';

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.getSheet('Data')!.printArea).toContain('$A$1:$H$50');
  });
});
```

**Step 4: Run tests and commit**

```bash
pnpm -C packages/modern-xlsx test
git add -A
git commit -m "feat(print): printTitles and printArea getters/setters via definedNames"
```

---

### Task 11: Row Grouping (Outline Levels)

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`
- Modify: `packages/modern-xlsx/src/workbook.ts`
- Modify: `packages/modern-xlsx/src/types.ts`

**Step 1: Add `outline_level` and `collapsed` to Rust `Row` struct**

```rust
pub struct Row {
    pub index: u32,
    pub cells: Vec<Cell>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub hidden: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub collapsed: bool,
}
```

**Step 2: Parse `outlineLevel` and `collapsed` in row reader**

In the `<row>` parsing section of `WorksheetXml::parse()`:
```rust
b"outlineLevel" => row.outline_level = std::str::from_utf8(&attr.value).ok().and_then(|v| v.parse().ok()),
b"collapsed" => row.collapsed = std::str::from_utf8(&attr.value).unwrap_or("0") == "1",
```

**Step 3: Write `outlineLevel` and `collapsed` in row writer**

In `to_xml_with_sst()`, when writing `<row>` start element:
```rust
if let Some(level) = row.outline_level {
    if level > 0 {
        row_elem.push_attribute(("outlineLevel", itoa::Buffer::new().format(level)));
    }
}
if row.collapsed {
    row_elem.push_attribute(("collapsed", "1"));
}
```

**Step 4: Add `outline_level` and `collapsed` to Rust `ColumnInfo`**

```rust
pub struct ColumnInfo {
    pub min: u32,
    pub max: u32,
    pub width: f64,
    pub hidden: bool,
    pub custom_width: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub collapsed: bool,
}
```

Parse and write `outlineLevel`/`collapsed` on `<col>` elements similarly.

**Step 5: Add `OutlineProperties` struct and parse `<sheetPr><outlinePr>`**

```rust
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OutlineProperties {
    #[serde(default = "default_true")]
    pub summary_below: bool,
    #[serde(default = "default_true")]
    pub summary_right: bool,
}
```

Add to `WorksheetXml`:
```rust
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_properties: Option<OutlineProperties>,
```

Parse from `<sheetPr><outlinePr>` in the reader, write in the writer.

**Step 6: Add TypeScript `groupRows()` and `groupColumns()`**

```typescript
groupRows(startRow: number, endRow: number, options?: { level?: number; collapsed?: boolean }): void {
  const level = options?.level ?? 1;
  const collapsed = options?.collapsed ?? false;
  for (let r = startRow; r <= endRow; r++) {
    const row = this.ensureRow(r);
    row.outlineLevel = level;
    if (collapsed) row.hidden = true;
  }
  if (collapsed) {
    // Set collapsed flag on the summary row (below the group)
    const summaryRow = this.ensureRow(endRow + 1);
    summaryRow.collapsed = true;
  }
}

groupColumns(startCol: number, endCol: number, options?: { level?: number; collapsed?: boolean }): void {
  // Similar — set outlineLevel on ColumnInfo entries
}
```

**Step 7: Update TypeScript types**

In `types.ts`, `RowData` already has `outlineLevel`. Add `collapsed`:
```typescript
export interface RowData {
  index: number;
  cells: CellData[];
  height: number | null;
  hidden: boolean;
  outlineLevel?: number | null;
  collapsed?: boolean;
}
```

`ColumnInfo` already has `outlineLevel`. Add `collapsed`:
```typescript
export interface ColumnInfo {
  min: number;
  max: number;
  width: number;
  hidden: boolean;
  customWidth: boolean;
  outlineLevel?: number | null;
  collapsed?: boolean;
}
```

Add to `WorksheetData`:
```typescript
  outlineProperties?: { summaryBelow?: boolean; summaryRight?: boolean } | null;
```

**Step 8: Write tests**

```rust
#[test]
fn test_row_outline_roundtrip() {
    let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>Header</v></c></row>
    <row r="2" outlineLevel="1"><c r="A2"><v>Detail 1</v></c></row>
    <row r="3" outlineLevel="1"><c r="A3"><v>Detail 2</v></c></row>
    <row r="4"><c r="A4"><v>Summary</v></c></row>
  </sheetData>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert_eq!(ws.rows[1].outline_level, Some(1));
    assert_eq!(ws.rows[2].outline_level, Some(1));

    // Roundtrip
    let xml2 = ws.to_xml_with_sst(None).unwrap();
    let ws2 = WorksheetXml::parse(&xml2).unwrap();
    assert_eq!(ws2.rows[1].outline_level, Some(1));
}

#[test]
fn test_collapsed_row_group() {
    let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2" outlineLevel="1" hidden="1"><c r="A2"><v>Hidden</v></c></row>
    <row r="3" outlineLevel="1" hidden="1"><c r="A3"><v>Hidden</v></c></row>
    <row r="4" collapsed="1"><c r="A4"><v>Summary</v></c></row>
  </sheetData>
</worksheet>"#;
    let ws = WorksheetXml::parse(xml.as_bytes()).unwrap();
    assert!(ws.rows[0].hidden);
    assert_eq!(ws.rows[0].outline_level, Some(1));
    assert!(ws.rows[2].collapsed);
}
```

TypeScript tests:
```typescript
it('should roundtrip row grouping', async () => {
  const wb = new Workbook();
  const ws = wb.addSheet('Groups');
  ws.cell('A1').value = 'Header';
  ws.cell('A2').value = 'Detail 1';
  ws.cell('A3').value = 'Detail 2';
  ws.cell('A4').value = 'Summary';
  ws.groupRows(1, 2, { level: 1 });

  const buf = await wb.toBuffer();
  const wb2 = await readBuffer(buf);
  const rows = wb2.getSheet('Groups')!.rows;
  const row2 = rows.find((r) => r.index === 2)!;
  expect(row2.outlineLevel).toBe(1);
});
```

**Step 9: Build WASM and run all tests**

```bash
cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm --no-opt
cargo test -p modern-xlsx-core
pnpm -C packages/modern-xlsx test
```

**Step 10: Commit**

```bash
git add -A
git commit -m "feat(print): row/column grouping with outlineLevel, collapsed, outlineProperties"
```

---

## Phase 3: Build, Audit & Ship

### Task 12: WASM Build + Full Test Suite

**Step 1: Build WASM**

```bash
cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm --no-opt
```

**Step 2: Run full Rust test suite**

```bash
cargo test -p modern-xlsx-core
```
Expected: All pass (including new table, header/footer, outline tests).

**Step 3: Build TypeScript**

```bash
pnpm -C packages/modern-xlsx build
```

**Step 4: Run full TypeScript test suite**

```bash
pnpm -C packages/modern-xlsx test
```
Expected: All pass.

**Step 5: Lint and typecheck**

```bash
pnpm -C packages/modern-xlsx lint
pnpm -C packages/modern-xlsx typecheck
cargo clippy -p modern-xlsx-core -- -D warnings
```

**Step 6: Commit any fixes**

---

### Task 13: Full Codebase Audit

Apply the audit principles to ALL new code:

1. **Modernization**: Every line uses newest Rust 1.94 / TS 6.0 idioms
2. **Performance**: `Vec::with_capacity()` on all parse buffers, `itoa::Buffer` for int formatting, `push_entity()` for XML text, no unnecessary allocations
3. **Correctness**: All error paths handled, no `unwrap()` on fallible operations in production code, all serde attributes correct

Run:
```bash
cargo clippy -p modern-xlsx-core -- -D warnings
pnpm -C packages/modern-xlsx lint
```

---

### Task 14: Update Wiki & Documentation

**Files:**
- Update: Wiki API-Reference.md
- Update: Wiki Changelog.md
- Update: Wiki Feature-Comparison.md
- Update: Wiki Home.md
- Update: Wiki Examples.md

Add documentation for:
- Table API (addTable, removeTable, getTable, tableToJson)
- HeaderFooterBuilder
- printTitles, printArea
- groupRows, groupColumns
- Built-in table style constants

---

### Task 15: Version Bump & Publish

**Step 1: Bump version to 0.6.0**

Update `packages/modern-xlsx/package.json` and `src/index.ts` VERSION constant.

**Step 2: Build and publish**

```bash
pnpm -C packages/modern-xlsx build
cd packages/modern-xlsx && npm publish --access public
```

**Step 3: Push to GitHub**

```bash
git push origin master
```

**Step 4: Create GitHub release**

```bash
gh release create v0.6.0 --title "v0.6.0 — Tables & Print Layout" --notes "..."
```
