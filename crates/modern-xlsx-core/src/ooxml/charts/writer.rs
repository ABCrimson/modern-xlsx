//! Chart XML writer — `ChartData::to_xml` and drawing XML generation.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;

use super::types::*;
use super::WorksheetChart;
use crate::{ModernXlsxError, Result};

// ---------------------------------------------------------------------------
// Drawing XML generation
// ---------------------------------------------------------------------------

impl ChartAnchor {
    /// Generate the complete `xl/drawings/drawing{n}.xml` for a worksheet's charts.
    ///
    /// Each chart gets a `<xdr:twoCellAnchor>` referencing its chart via the
    /// corresponding relationship ID from `chart_r_ids`.
    pub fn generate_drawing_xml(
        charts: &[WorksheetChart],
        chart_r_ids: &[String],
    ) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(512 + charts.len() * 512);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();

        let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

        // XML declaration.
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(map_err)?;

        // <xdr:wsDr>
        let mut root = BytesStart::new("xdr:wsDr");
        root.push_attribute(("xmlns:xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"));
        root.push_attribute(("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"));
        root.push_attribute(("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
        writer.write_event(Event::Start(root)).map_err(map_err)?;

        for (i, wsc) in charts.iter().enumerate() {
            let anchor = &wsc.anchor;
            let r_id = &chart_r_ids[i];
            let cnv_id = i as u32 + 2; // id starts at 2

            let is_one_cell = anchor.ext_cx.is_some() && anchor.ext_cy.is_some();
            let anchor_tag = if is_one_cell {
                "xdr:oneCellAnchor"
            } else {
                "xdr:twoCellAnchor"
            };

            // <xdr:twoCellAnchor> or <xdr:oneCellAnchor>
            writer
                .write_event(Event::Start(BytesStart::new(anchor_tag)))
                .map_err(map_err)?;

            // <xdr:from>
            writer
                .write_event(Event::Start(BytesStart::new("xdr:from")))
                .map_err(map_err)?;
            Self::write_cell_pos(&mut writer, anchor.from_col, anchor.from_col_off, anchor.from_row, anchor.from_row_off, &mut ibuf)?;
            writer
                .write_event(Event::End(BytesEnd::new("xdr:from")))
                .map_err(map_err)?;

            if is_one_cell {
                // <xdr:ext cx="..." cy="..."/>
                let mut ext_elem = BytesStart::new("xdr:ext");
                ext_elem.push_attribute(("cx", ibuf.format(anchor.ext_cx.unwrap_or(0))));
                ext_elem.push_attribute(("cy", ibuf.format(anchor.ext_cy.unwrap_or(0))));
                writer.write_event(Event::Empty(ext_elem)).map_err(map_err)?;
            } else {
                // <xdr:to>
                writer
                    .write_event(Event::Start(BytesStart::new("xdr:to")))
                    .map_err(map_err)?;
                Self::write_cell_pos(&mut writer, anchor.to_col, anchor.to_col_off, anchor.to_row, anchor.to_row_off, &mut ibuf)?;
                writer
                    .write_event(Event::End(BytesEnd::new("xdr:to")))
                    .map_err(map_err)?;
            }

            // <xdr:graphicFrame macro="">
            let mut gf = BytesStart::new("xdr:graphicFrame");
            gf.push_attribute(("macro", ""));
            writer.write_event(Event::Start(gf)).map_err(map_err)?;

            // <xdr:nvGraphicFramePr>
            writer
                .write_event(Event::Start(BytesStart::new("xdr:nvGraphicFramePr")))
                .map_err(map_err)?;
            let mut cnv_pr = BytesStart::new("xdr:cNvPr");
            cnv_pr.push_attribute(("id", ibuf.format(cnv_id)));
            let chart_name = format!("Chart {}", i + 1);
            cnv_pr.push_attribute(("name", chart_name.as_str()));
            writer.write_event(Event::Empty(cnv_pr)).map_err(map_err)?;
            writer
                .write_event(Event::Empty(BytesStart::new("xdr:cNvGraphicFramePr")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("xdr:nvGraphicFramePr")))
                .map_err(map_err)?;

            // <xdr:xfrm>
            writer
                .write_event(Event::Start(BytesStart::new("xdr:xfrm")))
                .map_err(map_err)?;
            let mut off = BytesStart::new("a:off");
            off.push_attribute(("x", "0"));
            off.push_attribute(("y", "0"));
            writer.write_event(Event::Empty(off)).map_err(map_err)?;
            let mut ext = BytesStart::new("a:ext");
            ext.push_attribute(("cx", "0"));
            ext.push_attribute(("cy", "0"));
            writer.write_event(Event::Empty(ext)).map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("xdr:xfrm")))
                .map_err(map_err)?;

            // <a:graphic>
            writer
                .write_event(Event::Start(BytesStart::new("a:graphic")))
                .map_err(map_err)?;
            let mut gd = BytesStart::new("a:graphicData");
            gd.push_attribute(("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart"));
            writer.write_event(Event::Start(gd)).map_err(map_err)?;
            let mut chart_ref = BytesStart::new("c:chart");
            chart_ref.push_attribute(("xmlns:c", "http://schemas.openxmlformats.org/drawingml/2006/chart"));
            chart_ref.push_attribute(("r:id", r_id.as_str()));
            writer.write_event(Event::Empty(chart_ref)).map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:graphicData")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:graphic")))
                .map_err(map_err)?;

            // </xdr:graphicFrame>
            writer
                .write_event(Event::End(BytesEnd::new("xdr:graphicFrame")))
                .map_err(map_err)?;

            // <xdr:clientData/>
            writer
                .write_event(Event::Empty(BytesStart::new("xdr:clientData")))
                .map_err(map_err)?;

            // </xdr:twoCellAnchor> or </xdr:oneCellAnchor>
            writer
                .write_event(Event::End(BytesEnd::new(anchor_tag)))
                .map_err(map_err)?;
        }

        // </xdr:wsDr>
        writer
            .write_event(Event::End(BytesEnd::new("xdr:wsDr")))
            .map_err(map_err)?;

        Ok(buf)
    }

    /// Write the `<xdr:col>`, `<xdr:colOff>`, `<xdr:row>`, `<xdr:rowOff>` children.
    fn write_cell_pos(
        writer: &mut Writer<&mut Vec<u8>>,
        col: u32,
        col_off: u64,
        row: u32,
        row_off: u64,
        ibuf: &mut itoa::Buffer,
    ) -> Result<()> {
        let map_err = |e: std::io::Error| ModernXlsxError::XmlWrite(e.to_string());

        writer
            .write_event(Event::Start(BytesStart::new("xdr:col")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(ibuf.format(col))))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("xdr:col")))
            .map_err(map_err)?;

        writer
            .write_event(Event::Start(BytesStart::new("xdr:colOff")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(ibuf.format(col_off))))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("xdr:colOff")))
            .map_err(map_err)?;

        writer
            .write_event(Event::Start(BytesStart::new("xdr:row")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(ibuf.format(row))))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("xdr:row")))
            .map_err(map_err)?;

        writer
            .write_event(Event::Start(BytesStart::new("xdr:rowOff")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(ibuf.format(row_off))))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("xdr:rowOff")))
            .map_err(map_err)?;

        Ok(())
    }
}

// ---------------------------------------------------------------------------
// XML Writer helpers
// ---------------------------------------------------------------------------

/// Map `std::io::Error` to `ModernXlsxError::XmlWrite`.
#[inline]
fn map_err(e: std::io::Error) -> ModernXlsxError {
    ModernXlsxError::XmlWrite(e.to_string())
}

impl ChartData {
    /// Serialize this chart to valid `xl/charts/chart{n}.xml` bytes.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut buf: Vec<u8> = Vec::with_capacity(1024);
        let mut writer = Writer::new(&mut buf);

        // XML declaration
        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(map_err)?;

        // <c:chartSpace>
        let mut cs = BytesStart::new("c:chartSpace");
        cs.push_attribute((
            "xmlns:c",
            "http://schemas.openxmlformats.org/drawingml/2006/chart",
        ));
        cs.push_attribute((
            "xmlns:a",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
        ));
        cs.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        writer.write_event(Event::Start(cs)).map_err(map_err)?;

        // <c:style> — ECMA-376 requires <c:style> before <c:chart>
        if let Some(style_id) = self.style_id {
            let mut ibuf = itoa::Buffer::new();
            let mut style = BytesStart::new("c:style");
            style.push_attribute(("val", ibuf.format(style_id)));
            writer.write_event(Event::Empty(style)).map_err(map_err)?;
        }

        // <c:chart>
        writer
            .write_event(Event::Start(BytesStart::new("c:chart")))
            .map_err(map_err)?;

        // <c:view3D>
        if let Some(ref v) = self.view_3d {
            Self::write_view_3d(&mut writer, v)?;
        }

        // <c:title> (chart-level)
        if let Some(ref title) = self.title {
            Self::write_title(&mut writer, title)?;
        }

        // <c:plotArea>
        writer
            .write_event(Event::Start(BytesStart::new("c:plotArea")))
            .map_err(map_err)?;

        // <c:layout> or <c:layout/>
        self.write_layout(&mut writer)?;

        // Chart-type-specific element
        self.write_chart_type_element(&mut writer)?;

        // Secondary chart type element (combo charts).
        if let Some(ref secondary) = self.secondary_chart {
            secondary.write_chart_type_element(&mut writer)?;
        }

        // Axes
        if let Some(ref axis) = self.cat_axis {
            Self::write_axis(&mut writer, "c:catAx", axis)?;
        }
        if let Some(ref axis) = self.val_axis {
            Self::write_axis(&mut writer, "c:valAx", axis)?;
        }

        // Secondary value axis (combo charts).
        if let Some(ref axis) = self.secondary_val_axis {
            Self::write_axis(&mut writer, "c:valAx", axis)?;
        }

        // <c:dTable>
        if self.show_data_table {
            writer
                .write_event(Event::Start(BytesStart::new("c:dTable")))
                .map_err(map_err)?;
            let mut keys = BytesStart::new("c:showKeys");
            keys.push_attribute(("val", "1"));
            writer.write_event(Event::Empty(keys)).map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:dTable")))
                .map_err(map_err)?;
        }

        // </c:plotArea>
        writer
            .write_event(Event::End(BytesEnd::new("c:plotArea")))
            .map_err(map_err)?;

        // <c:legend>
        if let Some(ref legend) = self.legend {
            Self::write_legend(&mut writer, legend)?;
        }

        // </c:chart>
        writer
            .write_event(Event::End(BytesEnd::new("c:chart")))
            .map_err(map_err)?;

        // <c:printSettings>
        Self::write_print_settings(&mut writer)?;

        // </c:chartSpace>
        writer
            .write_event(Event::End(BytesEnd::new("c:chartSpace")))
            .map_err(map_err)?;

        Ok(buf)
    }

    // -----------------------------------------------------------------------
    // Layout
    // -----------------------------------------------------------------------

    fn write_layout(&self, writer: &mut Writer<&mut Vec<u8>>) -> Result<()> {
        if let Some(ref layout) = self.plot_area_layout {
            writer
                .write_event(Event::Start(BytesStart::new("c:layout")))
                .map_err(map_err)?;

            writer
                .write_event(Event::Start(BytesStart::new("c:manualLayout")))
                .map_err(map_err)?;

            Self::write_f64_element(writer, "c:x", layout.x)?;
            Self::write_f64_element(writer, "c:y", layout.y)?;
            Self::write_f64_element(writer, "c:w", layout.w)?;
            Self::write_f64_element(writer, "c:h", layout.h)?;

            writer
                .write_event(Event::End(BytesEnd::new("c:manualLayout")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:layout")))
                .map_err(map_err)?;
        } else {
            writer
                .write_event(Event::Empty(BytesStart::new("c:layout")))
                .map_err(map_err)?;
        }
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Chart-type-specific element
    // -----------------------------------------------------------------------

    fn write_chart_type_element(&self, writer: &mut Writer<&mut Vec<u8>>) -> Result<()> {
        let tag = self.chart_type_xml_tag();

        writer
            .write_event(Event::Start(BytesStart::new(tag)))
            .map_err(map_err)?;

        let mut ibuf = itoa::Buffer::new();

        // Bar direction (bar/column)
        if matches!(self.chart_type, ChartType::Bar | ChartType::Column) {
            let dir = if self.chart_type == ChartType::Bar
                || self.bar_dir_horizontal == Some(true)
            {
                "bar"
            } else {
                "col"
            };
            let mut bd = BytesStart::new("c:barDir");
            bd.push_attribute(("val", dir));
            writer.write_event(Event::Empty(bd)).map_err(map_err)?;
        }

        // Grouping
        if let Some(grouping) = self.grouping
            && matches!(
                self.chart_type,
                ChartType::Bar
                    | ChartType::Column
                    | ChartType::Line
                    | ChartType::Area
            )
        {
            let mut g = BytesStart::new("c:grouping");
            g.push_attribute(("val", grouping.xml_val()));
            writer.write_event(Event::Empty(g)).map_err(map_err)?;
        }

        // Scatter style
        if let Some(style) = self.scatter_style
            && self.chart_type == ChartType::Scatter
        {
            let mut ss = BytesStart::new("c:scatterStyle");
            ss.push_attribute(("val", style.xml_val()));
            writer.write_event(Event::Empty(ss)).map_err(map_err)?;
        }

        // Radar style
        if let Some(style) = self.radar_style
            && self.chart_type == ChartType::Radar
        {
            let mut rs = BytesStart::new("c:radarStyle");
            rs.push_attribute(("val", style.xml_val()));
            writer.write_event(Event::Empty(rs)).map_err(map_err)?;
        }

        // Series
        let uses_xy =
            matches!(self.chart_type, ChartType::Scatter | ChartType::Bubble);

        for ser in &self.series {
            self.write_series(writer, ser, uses_xy, &mut ibuf)?;
        }

        // Chart-level data labels
        if let Some(ref dl) = self.data_labels {
            Self::write_data_labels(writer, dl)?;
        }

        // Hole size for doughnut
        if let Some(hole_size) = self.hole_size
            && self.chart_type == ChartType::Doughnut
        {
            let mut hs = BytesStart::new("c:holeSize");
            hs.push_attribute(("val", ibuf.format(hole_size)));
            writer.write_event(Event::Empty(hs)).map_err(map_err)?;
        }

        // Axis IDs for chart types that have axes
        if self.has_axes() {
            if let Some(ref cat_ax) = self.cat_axis {
                let mut id = BytesStart::new("c:axId");
                id.push_attribute(("val", ibuf.format(cat_ax.id)));
                writer.write_event(Event::Empty(id)).map_err(map_err)?;
            }
            if let Some(ref val_ax) = self.val_axis {
                let mut id = BytesStart::new("c:axId");
                id.push_attribute(("val", ibuf.format(val_ax.id)));
                writer.write_event(Event::Empty(id)).map_err(map_err)?;
            }
        }

        writer
            .write_event(Event::End(BytesEnd::new(tag)))
            .map_err(map_err)?;

        Ok(())
    }

    fn chart_type_xml_tag(&self) -> &'static str {
        match self.chart_type {
            ChartType::Bar | ChartType::Column => "c:barChart",
            ChartType::Line => "c:lineChart",
            ChartType::Pie => "c:pieChart",
            ChartType::Doughnut => "c:doughnutChart",
            ChartType::Scatter => "c:scatterChart",
            ChartType::Area => "c:areaChart",
            ChartType::Radar => "c:radarChart",
            ChartType::Bubble => "c:bubbleChart",
            ChartType::Stock => "c:stockChart",
        }
    }

    fn has_axes(&self) -> bool {
        !matches!(self.chart_type, ChartType::Pie | ChartType::Doughnut)
    }

    // -----------------------------------------------------------------------
    // Series
    // -----------------------------------------------------------------------

    fn write_series(
        &self,
        writer: &mut Writer<&mut Vec<u8>>,
        ser: &ChartSeries,
        uses_xy: bool,
        ibuf: &mut itoa::Buffer,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:ser")))
            .map_err(map_err)?;

        // <c:idx val="0"/>
        let mut idx = BytesStart::new("c:idx");
        idx.push_attribute(("val", ibuf.format(ser.idx)));
        writer.write_event(Event::Empty(idx)).map_err(map_err)?;

        // <c:order val="0"/>
        let mut ord = BytesStart::new("c:order");
        ord.push_attribute(("val", ibuf.format(ser.order)));
        writer.write_event(Event::Empty(ord)).map_err(map_err)?;

        // <c:tx>
        if let Some(ref name) = ser.name {
            writer
                .write_event(Event::Start(BytesStart::new("c:tx")))
                .map_err(map_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("c:strRef")))
                .map_err(map_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("c:f")))
                .map_err(map_err)?;
            writer
                .write_event(Event::Text(BytesText::new(name)))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:f")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:strRef")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:tx")))
                .map_err(map_err)?;
        }

        // <c:spPr>
        Self::write_sp_pr(writer, ser)?;

        // <c:marker>
        if let Some(marker) = ser.marker {
            Self::write_marker(writer, marker)?;
        }

        // Explosion (pie/doughnut)
        if let Some(explosion) = ser.explosion {
            let mut exp = BytesStart::new("c:explosion");
            exp.push_attribute(("val", ibuf.format(explosion)));
            writer.write_event(Event::Empty(exp)).map_err(map_err)?;
        }

        if uses_xy {
            // Scatter/Bubble: xVal + yVal
            if let Some(ref x_ref) = ser.x_val_ref {
                Self::write_ref_element(writer, "c:xVal", "c:numRef", x_ref)?;
            }
            Self::write_ref_element(writer, "c:yVal", "c:numRef", &ser.val_ref)?;

            // Bubble size
            if let Some(ref bub_ref) = ser.bubble_size_ref {
                Self::write_ref_element(writer, "c:bubbleSize", "c:numRef", bub_ref)?;
            }
        } else {
            // Standard: cat + val
            if let Some(ref cat_ref) = ser.cat_ref {
                Self::write_ref_element(writer, "c:cat", "c:strRef", cat_ref)?;
            }
            Self::write_ref_element(writer, "c:val", "c:numRef", &ser.val_ref)?;
        }

        // <c:smooth>
        if let Some(smooth) = ser.smooth {
            let mut sm = BytesStart::new("c:smooth");
            sm.push_attribute(("val", if smooth { "1" } else { "0" }));
            writer.write_event(Event::Empty(sm)).map_err(map_err)?;
        }

        // Series-level data labels
        if let Some(ref dl) = ser.data_labels {
            Self::write_data_labels(writer, dl)?;
        }

        // Trendline
        if let Some(ref tl) = ser.trendline {
            Self::write_trendline(writer, tl, ibuf)?;
        }

        // Error bars
        if let Some(ref eb) = ser.error_bars {
            Self::write_error_bars(writer, eb)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:ser")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Shape properties (fill + line)
    // -----------------------------------------------------------------------

    fn write_sp_pr(
        writer: &mut Writer<&mut Vec<u8>>,
        ser: &ChartSeries,
    ) -> Result<()> {
        if ser.fill_color.is_none() && ser.line_color.is_none() {
            return Ok(());
        }

        writer
            .write_event(Event::Start(BytesStart::new("c:spPr")))
            .map_err(map_err)?;

        // Fill
        if let Some(ref fill) = ser.fill_color {
            writer
                .write_event(Event::Start(BytesStart::new("a:solidFill")))
                .map_err(map_err)?;
            let mut clr = BytesStart::new("a:srgbClr");
            clr.push_attribute(("val", fill.as_str()));
            writer.write_event(Event::Empty(clr)).map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:solidFill")))
                .map_err(map_err)?;
        }

        // Line
        if let Some(ref line_color) = ser.line_color {
            let mut ln = BytesStart::new("a:ln");
            if let Some(w) = ser.line_width {
                let mut ibuf = itoa::Buffer::new();
                ln.push_attribute(("w", ibuf.format(w)));
            }
            writer.write_event(Event::Start(ln)).map_err(map_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("a:solidFill")))
                .map_err(map_err)?;
            let mut clr = BytesStart::new("a:srgbClr");
            clr.push_attribute(("val", line_color.as_str()));
            writer.write_event(Event::Empty(clr)).map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:solidFill")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:ln")))
                .map_err(map_err)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:spPr")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Marker
    // -----------------------------------------------------------------------

    fn write_marker(
        writer: &mut Writer<&mut Vec<u8>>,
        marker: MarkerStyle,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:marker")))
            .map_err(map_err)?;

        let mut sym = BytesStart::new("c:symbol");
        sym.push_attribute(("val", marker.xml_val()));
        writer.write_event(Event::Empty(sym)).map_err(map_err)?;

        let mut sz = BytesStart::new("c:size");
        sz.push_attribute(("val", "5"));
        writer.write_event(Event::Empty(sz)).map_err(map_err)?;

        writer
            .write_event(Event::End(BytesEnd::new("c:marker")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Reference elements (cat/val/xVal/yVal/bubbleSize)
    // -----------------------------------------------------------------------

    fn write_ref_element(
        writer: &mut Writer<&mut Vec<u8>>,
        outer_tag: &str,
        ref_tag: &str,
        formula: &str,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new(outer_tag)))
            .map_err(map_err)?;
        writer
            .write_event(Event::Start(BytesStart::new(ref_tag)))
            .map_err(map_err)?;
        writer
            .write_event(Event::Start(BytesStart::new("c:f")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(formula)))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("c:f")))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new(ref_tag)))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new(outer_tag)))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Title
    // -----------------------------------------------------------------------

    fn write_title(
        writer: &mut Writer<&mut Vec<u8>>,
        title: &ChartTitle,
    ) -> Result<()> {
        let mut ibuf = itoa::Buffer::new();

        writer
            .write_event(Event::Start(BytesStart::new("c:title")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Start(BytesStart::new("c:tx")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Start(BytesStart::new("c:rich")))
            .map_err(map_err)?;

        // <a:bodyPr/>
        writer
            .write_event(Event::Empty(BytesStart::new("a:bodyPr")))
            .map_err(map_err)?;
        // <a:lstStyle/>
        writer
            .write_event(Event::Empty(BytesStart::new("a:lstStyle")))
            .map_err(map_err)?;

        // <a:p>
        writer
            .write_event(Event::Start(BytesStart::new("a:p")))
            .map_err(map_err)?;

        let font_size = title.font_size.unwrap_or(1400);
        let bold = title.bold.unwrap_or(true);

        // <a:pPr>
        writer
            .write_event(Event::Start(BytesStart::new("a:pPr")))
            .map_err(map_err)?;

        // <a:defRPr>
        let mut def_rpr = BytesStart::new("a:defRPr");
        def_rpr.push_attribute(("sz", ibuf.format(font_size)));
        if bold {
            def_rpr.push_attribute(("b", "1"));
        }
        if let Some(ref color) = title.color {
            writer
                .write_event(Event::Start(def_rpr))
                .map_err(map_err)?;
            Self::write_solid_fill(writer, color)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:defRPr")))
                .map_err(map_err)?;
        } else {
            writer
                .write_event(Event::Empty(def_rpr))
                .map_err(map_err)?;
        }

        // </a:pPr>
        writer
            .write_event(Event::End(BytesEnd::new("a:pPr")))
            .map_err(map_err)?;

        // <a:r>
        writer
            .write_event(Event::Start(BytesStart::new("a:r")))
            .map_err(map_err)?;

        // <a:rPr>
        let mut rpr = BytesStart::new("a:rPr");
        rpr.push_attribute(("lang", "en-US"));
        rpr.push_attribute(("sz", ibuf.format(font_size)));
        if bold {
            rpr.push_attribute(("b", "1"));
        }
        if let Some(ref color) = title.color {
            writer.write_event(Event::Start(rpr)).map_err(map_err)?;
            Self::write_solid_fill(writer, color)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:rPr")))
                .map_err(map_err)?;
        } else {
            writer.write_event(Event::Empty(rpr)).map_err(map_err)?;
        }

        // <a:t>text</a:t>
        writer
            .write_event(Event::Start(BytesStart::new("a:t")))
            .map_err(map_err)?;
        writer
            .write_event(Event::Text(BytesText::new(&title.text)))
            .map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("a:t")))
            .map_err(map_err)?;

        // </a:r>
        writer
            .write_event(Event::End(BytesEnd::new("a:r")))
            .map_err(map_err)?;

        // </a:p>
        writer
            .write_event(Event::End(BytesEnd::new("a:p")))
            .map_err(map_err)?;

        // </c:rich>
        writer
            .write_event(Event::End(BytesEnd::new("c:rich")))
            .map_err(map_err)?;
        // </c:tx>
        writer
            .write_event(Event::End(BytesEnd::new("c:tx")))
            .map_err(map_err)?;

        // <c:overlay val="0"/>
        let mut overlay = BytesStart::new("c:overlay");
        overlay.push_attribute(("val", if title.overlay { "1" } else { "0" }));
        writer
            .write_event(Event::Empty(overlay))
            .map_err(map_err)?;

        // </c:title>
        writer
            .write_event(Event::End(BytesEnd::new("c:title")))
            .map_err(map_err)?;

        Ok(())
    }

    fn write_solid_fill(
        writer: &mut Writer<&mut Vec<u8>>,
        color: &str,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("a:solidFill")))
            .map_err(map_err)?;
        let mut clr = BytesStart::new("a:srgbClr");
        clr.push_attribute(("val", color));
        writer.write_event(Event::Empty(clr)).map_err(map_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("a:solidFill")))
            .map_err(map_err)?;
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Data labels
    // -----------------------------------------------------------------------

    fn write_data_labels(
        writer: &mut Writer<&mut Vec<u8>>,
        dl: &DataLabels,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:dLbls")))
            .map_err(map_err)?;

        Self::write_bool_element(writer, "c:showVal", dl.show_val)?;
        Self::write_bool_element(writer, "c:showCatName", dl.show_cat_name)?;
        Self::write_bool_element(writer, "c:showSerName", dl.show_ser_name)?;
        Self::write_bool_element(writer, "c:showPercent", dl.show_percent)?;
        Self::write_bool_element(writer, "c:showLeaderLines", dl.show_leader_lines)?;

        if let Some(ref fmt) = dl.num_fmt {
            let mut nf = BytesStart::new("c:numFmt");
            nf.push_attribute(("formatCode", fmt.as_str()));
            nf.push_attribute(("sourceLinked", "0"));
            writer.write_event(Event::Empty(nf)).map_err(map_err)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:dLbls")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Axis
    // -----------------------------------------------------------------------

    fn write_axis(
        writer: &mut Writer<&mut Vec<u8>>,
        tag: &str,
        axis: &ChartAxis,
    ) -> Result<()> {
        let mut ibuf = itoa::Buffer::new();

        writer
            .write_event(Event::Start(BytesStart::new(tag)))
            .map_err(map_err)?;

        // <c:axId>
        let mut ax_id = BytesStart::new("c:axId");
        ax_id.push_attribute(("val", ibuf.format(axis.id)));
        writer
            .write_event(Event::Empty(ax_id))
            .map_err(map_err)?;

        // <c:scaling>
        writer
            .write_event(Event::Start(BytesStart::new("c:scaling")))
            .map_err(map_err)?;

        let mut orient = BytesStart::new("c:orientation");
        orient.push_attribute(("val", if axis.reversed { "maxMin" } else { "minMax" }));
        writer
            .write_event(Event::Empty(orient))
            .map_err(map_err)?;

        if let Some(min) = axis.min {
            Self::write_f64_element(writer, "c:min", min)?;
        }
        if let Some(max) = axis.max {
            Self::write_f64_element(writer, "c:max", max)?;
        }
        if let Some(log_base) = axis.log_base {
            Self::write_f64_element(writer, "c:logBase", log_base)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:scaling")))
            .map_err(map_err)?;

        // <c:delete>
        let mut del = BytesStart::new("c:delete");
        del.push_attribute(("val", if axis.delete { "1" } else { "0" }));
        writer.write_event(Event::Empty(del)).map_err(map_err)?;

        // <c:axPos>
        let default_pos = if tag == "c:catAx" { "b" } else { "l" };
        let pos = axis
            .position
            .map(|p| p.xml_val())
            .unwrap_or(default_pos);
        let mut ax_pos = BytesStart::new("c:axPos");
        ax_pos.push_attribute(("val", pos));
        writer
            .write_event(Event::Empty(ax_pos))
            .map_err(map_err)?;

        // Axis title
        if let Some(ref title) = axis.title {
            Self::write_title(writer, title)?;
        }

        // Gridlines
        if axis.major_gridlines {
            writer
                .write_event(Event::Empty(BytesStart::new("c:majorGridlines")))
                .map_err(map_err)?;
        }
        if axis.minor_gridlines {
            writer
                .write_event(Event::Empty(BytesStart::new("c:minorGridlines")))
                .map_err(map_err)?;
        }

        // Number format
        {
            let fmt_code = axis.num_fmt.as_deref().unwrap_or("General");
            let src_linked = if axis.source_linked { "1" } else { "0" };
            let mut nf = BytesStart::new("c:numFmt");
            nf.push_attribute(("formatCode", fmt_code));
            nf.push_attribute(("sourceLinked", src_linked));
            writer.write_event(Event::Empty(nf)).map_err(map_err)?;
        }

        // Tick marks
        if let Some(tm) = axis.major_tick_mark {
            let mut el = BytesStart::new("c:majorTickMark");
            el.push_attribute(("val", tm.xml_val()));
            writer.write_event(Event::Empty(el)).map_err(map_err)?;
        }
        if let Some(tm) = axis.minor_tick_mark {
            let mut el = BytesStart::new("c:minorTickMark");
            el.push_attribute(("val", tm.xml_val()));
            writer.write_event(Event::Empty(el)).map_err(map_err)?;
        }

        // Tick label position
        if let Some(tlp) = axis.tick_lbl_pos {
            let mut el = BytesStart::new("c:tickLblPos");
            el.push_attribute(("val", tlp.xml_val()));
            writer.write_event(Event::Empty(el)).map_err(map_err)?;
        }

        // <c:crossAx>
        let mut cross = BytesStart::new("c:crossAx");
        cross.push_attribute(("val", ibuf.format(axis.cross_ax)));
        writer
            .write_event(Event::Empty(cross))
            .map_err(map_err)?;

        // Crosses at
        if let Some(crosses_at) = axis.crosses_at {
            Self::write_f64_element(writer, "c:crossesAt", crosses_at)?;
        }

        // Major/minor unit
        if let Some(major) = axis.major_unit {
            Self::write_f64_element(writer, "c:majorUnit", major)?;
        }
        if let Some(minor) = axis.minor_unit {
            Self::write_f64_element(writer, "c:minorUnit", minor)?;
        }

        // <c:txPr> — axis tick label font size
        if let Some(sz) = axis.font_size {
            writer
                .write_event(Event::Start(BytesStart::new("c:txPr")))
                .map_err(map_err)?;
            writer
                .write_event(Event::Empty(BytesStart::new("a:bodyPr")))
                .map_err(map_err)?;
            writer
                .write_event(Event::Empty(BytesStart::new("a:lstStyle")))
                .map_err(map_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("a:p")))
                .map_err(map_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("a:pPr")))
                .map_err(map_err)?;
            let mut def_rpr = BytesStart::new("a:defRPr");
            def_rpr.push_attribute(("sz", ibuf.format(sz)));
            writer
                .write_event(Event::Empty(def_rpr))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:pPr")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:p")))
                .map_err(map_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:txPr")))
                .map_err(map_err)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new(tag)))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Legend
    // -----------------------------------------------------------------------

    fn write_legend(
        writer: &mut Writer<&mut Vec<u8>>,
        legend: &ChartLegend,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:legend")))
            .map_err(map_err)?;

        let mut pos = BytesStart::new("c:legendPos");
        pos.push_attribute(("val", legend.position.xml_val()));
        writer.write_event(Event::Empty(pos)).map_err(map_err)?;

        let mut overlay = BytesStart::new("c:overlay");
        overlay.push_attribute(("val", if legend.overlay { "1" } else { "0" }));
        writer
            .write_event(Event::Empty(overlay))
            .map_err(map_err)?;

        writer
            .write_event(Event::End(BytesEnd::new("c:legend")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Trendline
    // -----------------------------------------------------------------------

    fn write_trendline(
        writer: &mut Writer<&mut Vec<u8>>,
        tl: &Trendline,
        ibuf: &mut itoa::Buffer,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:trendline")))
            .map_err(map_err)?;

        let mut tt = BytesStart::new("c:trendlineType");
        tt.push_attribute(("val", tl.trend_type.xml_val()));
        writer.write_event(Event::Empty(tt)).map_err(map_err)?;

        if let Some(order) = tl.order {
            let mut el = BytesStart::new("c:order");
            el.push_attribute(("val", ibuf.format(order)));
            writer.write_event(Event::Empty(el)).map_err(map_err)?;
        }
        if let Some(period) = tl.period {
            let mut el = BytesStart::new("c:period");
            el.push_attribute(("val", ibuf.format(period)));
            writer.write_event(Event::Empty(el)).map_err(map_err)?;
        }
        if let Some(fwd) = tl.forward {
            Self::write_f64_element(writer, "c:forward", fwd)?;
        }
        if let Some(bwd) = tl.backward {
            Self::write_f64_element(writer, "c:backward", bwd)?;
        }
        if tl.display_eq {
            Self::write_bool_element(writer, "c:dispEq", true)?;
        }
        if tl.display_r_sqr {
            Self::write_bool_element(writer, "c:dispRSqr", true)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:trendline")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Error bars
    // -----------------------------------------------------------------------

    fn write_error_bars(
        writer: &mut Writer<&mut Vec<u8>>,
        eb: &ErrorBars,
    ) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:errBars")))
            .map_err(map_err)?;

        let mut dir = BytesStart::new("c:errBarType");
        dir.push_attribute(("val", eb.direction.xml_val()));
        writer.write_event(Event::Empty(dir)).map_err(map_err)?;

        let mut vt = BytesStart::new("c:errValType");
        vt.push_attribute(("val", eb.err_type.xml_val()));
        writer.write_event(Event::Empty(vt)).map_err(map_err)?;

        if let Some(val) = eb.value {
            Self::write_f64_element(writer, "c:val", val)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:errBars")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // View 3D
    // -----------------------------------------------------------------------

    fn write_view_3d(
        writer: &mut Writer<&mut Vec<u8>>,
        v: &View3D,
    ) -> Result<()> {
        let mut ibuf = itoa::Buffer::new();

        writer
            .write_event(Event::Start(BytesStart::new("c:view3D")))
            .map_err(map_err)?;

        if let Some(rx) = v.rot_x {
            let mut el = BytesStart::new("c:rotX");
            el.push_attribute(("val", ibuf.format(rx)));
            writer.write_event(Event::Empty(el)).map_err(map_err)?;
        }
        if let Some(ry) = v.rot_y {
            let mut el = BytesStart::new("c:rotY");
            el.push_attribute(("val", ibuf.format(ry)));
            writer.write_event(Event::Empty(el)).map_err(map_err)?;
        }
        if let Some(p) = v.perspective {
            let mut el = BytesStart::new("c:perspective");
            el.push_attribute(("val", ibuf.format(p)));
            writer.write_event(Event::Empty(el)).map_err(map_err)?;
        }
        if let Some(ra) = v.r_ang_ax {
            let mut el = BytesStart::new("c:rAngAx");
            el.push_attribute(("val", if ra { "1" } else { "0" }));
            writer.write_event(Event::Empty(el)).map_err(map_err)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:view3D")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Print settings (always present)
    // -----------------------------------------------------------------------

    fn write_print_settings(writer: &mut Writer<&mut Vec<u8>>) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:printSettings")))
            .map_err(map_err)?;

        // <c:headerFooter/>
        writer
            .write_event(Event::Empty(BytesStart::new("c:headerFooter")))
            .map_err(map_err)?;

        // <c:pageMargins>
        let mut pm = BytesStart::new("c:pageMargins");
        pm.push_attribute(("b", "0.75"));
        pm.push_attribute(("l", "0.7"));
        pm.push_attribute(("r", "0.7"));
        pm.push_attribute(("t", "0.75"));
        pm.push_attribute(("header", "0.3"));
        pm.push_attribute(("footer", "0.3"));
        writer.write_event(Event::Empty(pm)).map_err(map_err)?;

        // <c:pageSetup/>
        writer
            .write_event(Event::Empty(BytesStart::new("c:pageSetup")))
            .map_err(map_err)?;

        writer
            .write_event(Event::End(BytesEnd::new("c:printSettings")))
            .map_err(map_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Utility helpers
    // -----------------------------------------------------------------------

    fn write_bool_element(
        writer: &mut Writer<&mut Vec<u8>>,
        tag: &str,
        value: bool,
    ) -> Result<()> {
        let mut el = BytesStart::new(tag);
        el.push_attribute(("val", if value { "1" } else { "0" }));
        writer.write_event(Event::Empty(el)).map_err(map_err)?;
        Ok(())
    }

    fn write_f64_element(
        writer: &mut Writer<&mut Vec<u8>>,
        tag: &str,
        value: f64,
    ) -> Result<()> {
        if !value.is_finite() {
            return Ok(()); // Skip NaN/Infinity — not valid xsd:double in OOXML
        }
        let mut el = BytesStart::new(tag);
        let formatted = value.to_string();
        el.push_attribute(("val", formatted.as_str()));
        writer.write_event(Event::Empty(el)).map_err(map_err)?;
        Ok(())
    }
}
