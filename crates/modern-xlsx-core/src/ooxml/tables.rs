//! Excel Table (ListObject) definitions — `xl/tables/table{n}.xml`.

use serde::{Deserialize, Serialize};

fn default_header_row_count() -> u32 {
    1
}

fn default_true() -> bool {
    true
}

fn is_false(v: &bool) -> bool {
    !v
}

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
    /// Cell range including header row, e.g. `"A1:D10"`.
    #[serde(rename = "ref")]
    pub ref_range: String,
    /// Number of header rows (default 1, set 0 for no header).
    #[serde(default = "default_header_row_count")]
    pub header_row_count: u32,
    /// Number of totals rows (default 0).
    #[serde(default)]
    pub totals_row_count: u32,
    /// Whether the totals row is shown.
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

/// A column within a table definition.
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableColumn {
    /// Column ID (unique within the table).
    pub id: u32,
    /// Column name (matches header cell text).
    pub name: String,
    /// Totals row aggregate function (sum, min, max, average, count, countNums, stdDev, var, custom, none).
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
    /// Built-in style name (e.g. `"TableStyleMedium2"`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub show_first_column: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub show_last_column: bool,
    #[serde(default = "default_true")]
    pub show_row_stripes: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub show_column_stripes: bool,
}

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
                TableColumn {
                    id: 1,
                    name: "Name".into(),
                    ..Default::default()
                },
                TableColumn {
                    id: 2,
                    name: "Age".into(),
                    ..Default::default()
                },
                TableColumn {
                    id: 3,
                    name: "City".into(),
                    ..Default::default()
                },
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
        assert_eq!(parsed.columns[0].name, "Name");
        assert!(parsed.style_info.is_some());
    }

    #[test]
    fn test_table_column_with_formula() {
        let col = TableColumn {
            id: 1,
            name: "Total".into(),
            totals_row_function: Some("sum".into()),
            calculated_column_formula: Some("Sales[@Qty]*Sales[@Price]".into()),
            ..Default::default()
        };
        let json = serde_json::to_string(&col).unwrap();
        let parsed: TableColumn = serde_json::from_str(&json).unwrap();
        assert_eq!(parsed.totals_row_function.as_deref(), Some("sum"));
        assert_eq!(
            parsed.calculated_column_formula.as_deref(),
            Some("Sales[@Qty]*Sales[@Price]")
        );
    }

    #[test]
    fn test_table_style_info_defaults() {
        let json = r#"{"showRowStripes":true}"#;
        let si: TableStyleInfo = serde_json::from_str(json).unwrap();
        assert!(si.show_row_stripes);
        assert!(!si.show_first_column);
        assert!(!si.show_last_column);
        assert!(!si.show_column_stripes);
        assert!(si.name.is_none());
    }
}
