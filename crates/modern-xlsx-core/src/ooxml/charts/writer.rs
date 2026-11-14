//! Chart XML writer — [`ChartData::to_xml`] for `xl/charts/chart{n}.xml` and
//! [`ChartAnchor::generate_drawing_xml`] for `xl/drawings/drawing{n}.xml`.

use log::warn;
use quick_xml::Writer;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

use super::WorksheetChart;
use super::types::{
    ChartAnchor, ChartAxis, ChartData, ChartGrouping, ChartLegend, ChartSeries, ChartTitle,
    ChartType, DataLabels, ErrorBars, MarkerStyle, RadarStyle, ScatterStyle, Trendline, View3D,
};
use crate::ooxml::xml::{
    DML_CHART_NS, DML_MAIN_NS, REL_NS, XDR_NS, XsdBuf, fmt_xsd_double, xml_attr,
};
use crate::ooxml::{write_xml_decl, xml_write_err};
use crate::{ModernXlsxError, Result};

// ---------------------------------------------------------------------------
// Drawing XML generation
// ---------------------------------------------------------------------------

impl ChartAnchor {
    /// Generate the complete `xl/drawings/drawing{n}.xml` for a worksheet's charts.
    ///
    /// Each chart gets an anchor referencing its chart via the corresponding
    /// relationship ID from `chart_r_ids`: an `<xdr:absoluteAnchor>` when `pos_x`,
    /// `pos_y`, `ext_cx` and `ext_cy` are all set, else an `<xdr:oneCellAnchor>` when
    /// `ext_cx` and `ext_cy` are set, else an `<xdr:twoCellAnchor>`.
    ///
    /// # Errors
    ///
    /// `XmlWrite` when `chart_r_ids` and `charts` differ in length, or when XML
    /// writing fails.
    pub fn generate_drawing_xml(
        charts: &[WorksheetChart],
        chart_r_ids: &[String],
    ) -> Result<Vec<u8>> {
        if charts.len() != chart_r_ids.len() {
            return Err(ModernXlsxError::XmlWrite(format!(
                "drawing: {} chart(s) but {} relationship id(s)",
                charts.len(),
                chart_r_ids.len()
            )));
        }
        let mut buf: Vec<u8> = Vec::with_capacity(512 + charts.len() * 512);
        let mut writer = Writer::new(&mut buf);
        let mut ibuf = itoa::Buffer::new();

        write_xml_decl(&mut writer)?;

        // <xdr:wsDr>
        let mut root = BytesStart::new("xdr:wsDr");
        root.push_attribute(("xmlns:xdr", XDR_NS));
        root.push_attribute(("xmlns:a", DML_MAIN_NS));
        root.push_attribute(("xmlns:r", REL_NS));
        writer
            .write_event(Event::Start(root))
            .map_err(xml_write_err)?;

        for (i, wsc) in charts.iter().enumerate() {
            let anchor = &wsc.anchor;
            let r_id = &chart_r_ids[i]; // lengths checked above
            let cnv_id = i as u32 + 2; // id starts at 2

            let has_ext = anchor.ext_cx.is_some() && anchor.ext_cy.is_some();
            let abs_pos = anchor.pos_x.zip(anchor.pos_y).filter(|_| has_ext);
            let is_one_cell = has_ext && abs_pos.is_none();
            let anchor_tag = if abs_pos.is_some() {
                "xdr:absoluteAnchor"
            } else if is_one_cell {
                "xdr:oneCellAnchor"
            } else {
                "xdr:twoCellAnchor"
            };

            // <xdr:twoCellAnchor>, <xdr:oneCellAnchor> or <xdr:absoluteAnchor>
            writer
                .write_event(Event::Start(BytesStart::new(anchor_tag)))
                .map_err(xml_write_err)?;

            if let Some((x, y)) = abs_pos {
                // <xdr:pos x="..." y="..."/>
                let mut pos = BytesStart::new("xdr:pos");
                pos.push_attribute(("x", ibuf.format(x)));
                pos.push_attribute(("y", ibuf.format(y)));
                writer
                    .write_event(Event::Empty(pos))
                    .map_err(xml_write_err)?;
            } else {
                // <xdr:from>
                writer
                    .write_event(Event::Start(BytesStart::new("xdr:from")))
                    .map_err(xml_write_err)?;
                Self::write_cell_pos(
                    &mut writer,
                    anchor.from_col,
                    anchor.from_col_off,
                    anchor.from_row,
                    anchor.from_row_off,
                    &mut ibuf,
                )?;
                writer
                    .write_event(Event::End(BytesEnd::new("xdr:from")))
                    .map_err(xml_write_err)?;
            }

            if has_ext {
                // <xdr:ext cx="..." cy="..."/>
                let mut ext_elem = BytesStart::new("xdr:ext");
                ext_elem.push_attribute(("cx", ibuf.format(anchor.ext_cx.unwrap_or(0))));
                ext_elem.push_attribute(("cy", ibuf.format(anchor.ext_cy.unwrap_or(0))));
                writer
                    .write_event(Event::Empty(ext_elem))
                    .map_err(xml_write_err)?;
            } else {
                // <xdr:to>
                writer
                    .write_event(Event::Start(BytesStart::new("xdr:to")))
                    .map_err(xml_write_err)?;
                Self::write_cell_pos(
                    &mut writer,
                    anchor.to_col,
                    anchor.to_col_off,
                    anchor.to_row,
                    anchor.to_row_off,
                    &mut ibuf,
                )?;
                writer
                    .write_event(Event::End(BytesEnd::new("xdr:to")))
                    .map_err(xml_write_err)?;
            }

            // <xdr:graphicFrame macro="">
            let mut gf = BytesStart::new("xdr:graphicFrame");
            gf.push_attribute(("macro", ""));
            writer
                .write_event(Event::Start(gf))
                .map_err(xml_write_err)?;

            // <xdr:nvGraphicFramePr>
            writer
                .write_event(Event::Start(BytesStart::new("xdr:nvGraphicFramePr")))
                .map_err(xml_write_err)?;
            let mut cnv_pr = BytesStart::new("xdr:cNvPr");
            cnv_pr.push_attribute(("id", ibuf.format(cnv_id)));
            let chart_name = format!("Chart {}", i + 1);
            cnv_pr.push_attribute(("name", chart_name.as_str()));
            writer
                .write_event(Event::Empty(cnv_pr))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::Empty(BytesStart::new("xdr:cNvGraphicFramePr")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("xdr:nvGraphicFramePr")))
                .map_err(xml_write_err)?;

            // <xdr:xfrm>
            writer
                .write_event(Event::Start(BytesStart::new("xdr:xfrm")))
                .map_err(xml_write_err)?;
            let mut off = BytesStart::new("a:off");
            off.push_attribute(("x", "0"));
            off.push_attribute(("y", "0"));
            writer
                .write_event(Event::Empty(off))
                .map_err(xml_write_err)?;
            let mut ext = BytesStart::new("a:ext");
            ext.push_attribute(("cx", "0"));
            ext.push_attribute(("cy", "0"));
            writer
                .write_event(Event::Empty(ext))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("xdr:xfrm")))
                .map_err(xml_write_err)?;

            // <a:graphic>
            writer
                .write_event(Event::Start(BytesStart::new("a:graphic")))
                .map_err(xml_write_err)?;
            let mut gd = BytesStart::new("a:graphicData");
            gd.push_attribute(("uri", DML_CHART_NS));
            writer
                .write_event(Event::Start(gd))
                .map_err(xml_write_err)?;
            let mut chart_ref = BytesStart::new("c:chart");
            chart_ref.push_attribute(("xmlns:c", DML_CHART_NS));
            chart_ref.push_attribute(("r:id", r_id.as_str()));
            writer
                .write_event(Event::Empty(chart_ref))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:graphicData")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:graphic")))
                .map_err(xml_write_err)?;

            // </xdr:graphicFrame>
            writer
                .write_event(Event::End(BytesEnd::new("xdr:graphicFrame")))
                .map_err(xml_write_err)?;

            // <xdr:clientData/>
            writer
                .write_event(Event::Empty(BytesStart::new("xdr:clientData")))
                .map_err(xml_write_err)?;

            // </xdr:twoCellAnchor>, </xdr:oneCellAnchor> or </xdr:absoluteAnchor>
            writer
                .write_event(Event::End(BytesEnd::new(anchor_tag)))
                .map_err(xml_write_err)?;
        }

        // </xdr:wsDr>
        writer
            .write_event(Event::End(BytesEnd::new("xdr:wsDr")))
            .map_err(xml_write_err)?;

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
        writer
            .write_event(Event::Start(BytesStart::new("xdr:col")))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::Text(BytesText::new(ibuf.format(col))))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("xdr:col")))
            .map_err(xml_write_err)?;

        writer
            .write_event(Event::Start(BytesStart::new("xdr:colOff")))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::Text(BytesText::new(ibuf.format(col_off))))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("xdr:colOff")))
            .map_err(xml_write_err)?;

        writer
            .write_event(Event::Start(BytesStart::new("xdr:row")))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::Text(BytesText::new(ibuf.format(row))))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("xdr:row")))
            .map_err(xml_write_err)?;

        writer
            .write_event(Event::Start(BytesStart::new("xdr:rowOff")))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::Text(BytesText::new(ibuf.format(row_off))))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("xdr:rowOff")))
            .map_err(xml_write_err)?;

        Ok(())
    }
}

// ---------------------------------------------------------------------------
// XML Writer helpers
// ---------------------------------------------------------------------------

impl ChartData {
    /// Serialize this chart to valid `xl/charts/chart{n}.xml` bytes, or return
    /// [`ChartData::raw_xml`] unchanged when it is set.
    ///
    /// Every chart type except pie and doughnut is drawn on two axes, which the
    /// schema requires its chart-type element to name. An axis slot the model leaves
    /// unset is written with a default axis (an unused id, crossing the other axis
    /// of the pair), so a chart with no axes at all gets the TypeScript
    /// `ChartBuilder`'s pair: ids 0 and 1, at the bottom and on the left. The
    /// secondary chart of a combo takes `secondary_cat_axis` and
    /// `secondary_val_axis`, else the axis the primary chart is drawn on in that
    /// slot. A chart that sets the axes it is drawn on gets no default axis.
    ///
    /// # Errors
    ///
    /// `XmlWrite` when the secondary chart has a secondary chart of its own, or
    /// when XML writing fails.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        if let Some(ref raw) = self.raw_xml {
            return Ok(raw.as_bytes().to_vec());
        }
        if self
            .secondary_chart
            .as_ref()
            .is_some_and(|s| s.secondary_chart.is_some())
        {
            return Err(ModernXlsxError::XmlWrite(
                "nested secondary charts are not supported; use at most one secondaryChart".into(),
            ));
        }

        let mut buf: Vec<u8> = Vec::with_capacity(1024);
        let mut writer = Writer::new(&mut buf);

        write_xml_decl(&mut writer)?;

        // <c:chartSpace>
        let mut cs = BytesStart::new("c:chartSpace");
        cs.push_attribute(("xmlns:c", DML_CHART_NS));
        cs.push_attribute(("xmlns:a", DML_MAIN_NS));
        cs.push_attribute(("xmlns:r", REL_NS));
        writer
            .write_event(Event::Start(cs))
            .map_err(xml_write_err)?;

        // <c:style>: in CT_ChartSpace it precedes <c:chart>
        if let Some(style_id) = self.style_id {
            let mut ibuf = itoa::Buffer::new();
            let mut style = BytesStart::new("c:style");
            style.push_attribute(("val", ibuf.format(style_id)));
            writer
                .write_event(Event::Empty(style))
                .map_err(xml_write_err)?;
        }

        // <c:chart>
        writer
            .write_event(Event::Start(BytesStart::new("c:chart")))
            .map_err(xml_write_err)?;

        // CT_Chart order: title, autoTitleDeleted, view3D, plotArea, legend.
        // <c:title> (chart-level)
        if let Some(ref title) = self.title {
            Self::write_title(&mut writer, title)?;
        }

        // <c:autoTitleDeleted>: without it Excel adds an automatic title to a
        // single-series chart that has none.
        Self::write_bool_element(
            &mut writer,
            "c:autoTitleDeleted",
            self.auto_title_deleted.unwrap_or(self.title.is_none()),
        )?;

        // <c:view3D>
        if let Some(ref v) = self.view_3d {
            Self::write_view_3d(&mut writer, v)?;
        }

        // <c:plotArea>
        writer
            .write_event(Event::Start(BytesStart::new("c:plotArea")))
            .map_err(xml_write_err)?;

        // <c:layout> or <c:layout/>
        self.write_layout(&mut writer)?;

        // Chart-type-specific element, then the secondary one of a combo chart; each
        // names the two axes it is drawn on.
        let xy = self.uses_xy();
        let primary_axes = if xy && self.cat_axis.is_none() {
            // Scatter/bubble read from a file: X is val_axis, Y is secondary_val_axis.
            [self.val_axis.as_ref(), self.secondary_val_axis.as_ref()]
        } else {
            [self.cat_axis.as_ref(), self.val_axis.as_ref()]
        };
        // A chart type with axes needs two c:axId, each naming an axis of the plot
        // area: an unset slot gets a default axis, written with the others below.
        // The pair to complete is the primary one; under a primary chart without
        // axes (pie, doughnut) it is the pair an axis-bearing secondary chart is
        // drawn on. Both elements take their ids from this one decision.
        let secondary_with_axes = self.secondary_chart.as_deref().filter(|s| s.has_axes());
        let drawn_pair = if self.has_axes() {
            Some((primary_axes, xy))
        } else {
            secondary_with_axes.map(|secondary| {
                let pair = [
                    self.secondary_cat_axis.as_ref().or(primary_axes[0]),
                    self.secondary_val_axis.as_ref().or(primary_axes[1]),
                ];
                (pair, secondary.uses_xy())
            })
        };
        let ([synth_first, synth_second], synth_xy) = drawn_pair
            .map_or(([None, None], false), |(pair, pair_xy)| {
                (self.synthesised_axes(pair), pair_xy)
            });
        let primary_axes = [
            primary_axes[0].or(synth_first.as_ref()),
            primary_axes[1].or(synth_second.as_ref()),
        ];
        self.write_chart_type_element(&mut writer, primary_axes)?;
        if let Some(ref secondary) = self.secondary_chart {
            // A secondary slot left unset shares the primary chart's axis in it.
            let secondary_axes = [
                self.secondary_cat_axis.as_ref().or(primary_axes[0]),
                self.secondary_val_axis.as_ref().or(primary_axes[1]),
            ];
            secondary.write_chart_type_element(&mut writer, secondary_axes)?;
        }

        // Axes. Scatter and bubble charts have value axes only, so a set cat_axis is
        // their X axis, written as c:valAx at the bottom.
        if let Some(ref axis) = self.cat_axis {
            let tag = if xy { "c:valAx" } else { "c:catAx" };
            Self::write_axis(&mut writer, tag, axis, "b")?;
        }
        if let Some(ref axis) = self.val_axis {
            let default_pos = if xy && self.cat_axis.is_none() {
                "b"
            } else {
                "l"
            };
            Self::write_axis(&mut writer, "c:valAx", axis, default_pos)?;
        }

        // Secondary axes (combo charts; the Y axis of a scatter or bubble chart).
        if let Some(ref axis) = self.secondary_val_axis {
            Self::write_axis(&mut writer, "c:valAx", axis, "l")?;
        }
        if let Some(ref axis) = self.secondary_cat_axis {
            Self::write_axis(&mut writer, "c:catAx", axis, "b")?;
        }

        // Default axes standing in for unset slots of the pair the charts are drawn on.
        if let Some(ref axis) = synth_first {
            let tag = if synth_xy { "c:valAx" } else { "c:catAx" };
            Self::write_axis(&mut writer, tag, axis, "b")?;
        }
        if let Some(ref axis) = synth_second {
            Self::write_axis(&mut writer, "c:valAx", axis, "l")?;
        }

        // <c:dTable>
        if self.show_data_table {
            writer
                .write_event(Event::Start(BytesStart::new("c:dTable")))
                .map_err(xml_write_err)?;
            let mut keys = BytesStart::new("c:showKeys");
            keys.push_attribute(("val", "1"));
            writer
                .write_event(Event::Empty(keys))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:dTable")))
                .map_err(xml_write_err)?;
        }

        // </c:plotArea>
        writer
            .write_event(Event::End(BytesEnd::new("c:plotArea")))
            .map_err(xml_write_err)?;

        // <c:legend>
        if let Some(ref legend) = self.legend {
            Self::write_legend(&mut writer, legend)?;
        }

        // </c:chart>
        writer
            .write_event(Event::End(BytesEnd::new("c:chart")))
            .map_err(xml_write_err)?;

        // <c:printSettings>
        Self::write_print_settings(&mut writer)?;

        // </c:chartSpace>
        writer
            .write_event(Event::End(BytesEnd::new("c:chartSpace")))
            .map_err(xml_write_err)?;

        Ok(buf)
    }

    // -----------------------------------------------------------------------
    // Layout
    // -----------------------------------------------------------------------

    fn write_layout(&self, writer: &mut Writer<&mut Vec<u8>>) -> Result<()> {
        if let Some(ref layout) = self.plot_area_layout {
            writer
                .write_event(Event::Start(BytesStart::new("c:layout")))
                .map_err(xml_write_err)?;

            writer
                .write_event(Event::Start(BytesStart::new("c:manualLayout")))
                .map_err(xml_write_err)?;

            Self::write_f64_element(writer, "c:x", layout.x)?;
            Self::write_f64_element(writer, "c:y", layout.y)?;
            Self::write_f64_element(writer, "c:w", layout.w)?;
            Self::write_f64_element(writer, "c:h", layout.h)?;

            writer
                .write_event(Event::End(BytesEnd::new("c:manualLayout")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:layout")))
                .map_err(xml_write_err)?;
        } else {
            writer
                .write_event(Event::Empty(BytesStart::new("c:layout")))
                .map_err(xml_write_err)?;
        }
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Chart-type-specific element
    // -----------------------------------------------------------------------

    /// Write this chart's chart-type element, ending in a `c:axId` for each set
    /// axis of `axes` when the chart type has axes.
    fn write_chart_type_element(
        &self,
        writer: &mut Writer<&mut Vec<u8>>,
        axes: [Option<&ChartAxis>; 2],
    ) -> Result<()> {
        let tag = self.chart_type_xml_tag();

        writer
            .write_event(Event::Start(BytesStart::new(tag)))
            .map_err(xml_write_err)?;

        let mut ibuf = itoa::Buffer::new();

        // Bar direction (bar/column)
        if matches!(self.chart_type, ChartType::Bar | ChartType::Column) {
            let dir = if self.chart_type == ChartType::Bar || self.bar_dir_horizontal == Some(true)
            {
                "bar"
            } else {
                "col"
            };
            let mut bd = BytesStart::new("c:barDir");
            bd.push_attribute(("val", dir));
            writer
                .write_event(Event::Empty(bd))
                .map_err(xml_write_err)?;
        }

        // Grouping
        if let Some(grouping) = self.grouping
            && matches!(
                self.chart_type,
                ChartType::Bar | ChartType::Column | ChartType::Line | ChartType::Area
            )
        {
            let mut g = BytesStart::new("c:grouping");
            g.push_attribute(("val", grouping.xml_val()));
            writer.write_event(Event::Empty(g)).map_err(xml_write_err)?;
        }

        // CT_ScatterChart requires c:scatterStyle and CT_RadarChart requires
        // c:radarStyle (minOccurs 1); an unset style gets Excel's default.
        if self.chart_type == ChartType::Scatter {
            let style = self.scatter_style.unwrap_or(ScatterStyle::LineMarker);
            let mut ss = BytesStart::new("c:scatterStyle");
            ss.push_attribute(("val", style.xml_val()));
            writer
                .write_event(Event::Empty(ss))
                .map_err(xml_write_err)?;
        }

        if self.chart_type == ChartType::Radar {
            let style = self.radar_style.unwrap_or(RadarStyle::Marker);
            let mut rs = BytesStart::new("c:radarStyle");
            rs.push_attribute(("val", style.xml_val()));
            writer
                .write_event(Event::Empty(rs))
                .map_err(xml_write_err)?;
        }

        // Series
        let uses_xy = self.uses_xy();

        for ser in &self.series {
            self.write_series(writer, ser, uses_xy, &mut ibuf)?;
        }

        // Chart-level data labels
        if let Some(ref dl) = self.data_labels {
            Self::write_data_labels(writer, dl)?;
        }

        // Stacked bars must overlap fully, or Excel draws them side by side.
        if matches!(self.chart_type, ChartType::Bar | ChartType::Column)
            && matches!(
                self.grouping,
                Some(ChartGrouping::Stacked | ChartGrouping::PercentStacked)
            )
        {
            let mut ov = BytesStart::new("c:overlap");
            ov.push_attribute(("val", "100"));
            writer
                .write_event(Event::Empty(ov))
                .map_err(xml_write_err)?;
        }

        // Hole size for doughnut
        if let Some(hole_size) = self.hole_size
            && self.chart_type == ChartType::Doughnut
        {
            let mut hs = BytesStart::new("c:holeSize");
            hs.push_attribute(("val", ibuf.format(hole_size)));
            writer
                .write_event(Event::Empty(hs))
                .map_err(xml_write_err)?;
        }

        // Axis IDs for chart types that have axes
        if self.has_axes() {
            for axis in axes.into_iter().flatten() {
                let mut id = BytesStart::new("c:axId");
                id.push_attribute(("val", ibuf.format(axis.id)));
                writer
                    .write_event(Event::Empty(id))
                    .map_err(xml_write_err)?;
            }
        }

        writer
            .write_event(Event::End(BytesEnd::new(tag)))
            .map_err(xml_write_err)?;

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

    /// A default axis for each unset slot of `axes`, the axis pair the chart-type
    /// elements are drawn on: its id is the smallest one no axis of this chart (or an
    /// earlier synthesised axis) uses, and it crosses the axis in the other slot. With
    /// no axis set that is the TypeScript `ChartBuilder`'s pair, ids 0 and 1.
    fn synthesised_axes(&self, axes: [Option<&ChartAxis>; 2]) -> [Option<ChartAxis>; 2] {
        let mut used: Vec<u32> = [
            &self.cat_axis,
            &self.val_axis,
            &self.secondary_cat_axis,
            &self.secondary_val_axis,
        ]
        .into_iter()
        .flatten()
        .map(|axis| axis.id)
        .collect();
        let mut unused_id = || {
            // Ends within used.len() + 1 steps.
            let mut id = 0;
            while used.contains(&id) {
                id += 1;
            }
            used.push(id);
            id
        };
        // Slot order: the first slot's id is chosen before the second's.
        let first = axes[0].map_or_else(&mut unused_id, |axis| axis.id);
        let second = axes[1].map_or_else(&mut unused_id, |axis| axis.id);
        let synthesise = |set: Option<&ChartAxis>, id, cross_ax| {
            set.is_none().then(|| ChartAxis {
                id,
                cross_ax,
                ..Default::default()
            })
        };
        [
            synthesise(axes[0], first, second),
            synthesise(axes[1], second, first),
        ]
    }

    fn has_axes(&self) -> bool {
        !matches!(self.chart_type, ChartType::Pie | ChartType::Doughnut)
    }

    /// Scatter and bubble series plot xVal/yVal on two value axes.
    fn uses_xy(&self) -> bool {
        matches!(self.chart_type, ChartType::Scatter | ChartType::Bubble)
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
            .map_err(xml_write_err)?;

        // <c:idx val="0"/>
        let mut idx = BytesStart::new("c:idx");
        idx.push_attribute(("val", ibuf.format(ser.idx)));
        writer
            .write_event(Event::Empty(idx))
            .map_err(xml_write_err)?;

        // <c:order val="0"/>
        let mut ord = BytesStart::new("c:order");
        ord.push_attribute(("val", ibuf.format(ser.order)));
        writer
            .write_event(Event::Empty(ord))
            .map_err(xml_write_err)?;

        // <c:tx> (CT_SerTx): the name formula as c:strRef/c:f, else the literal
        // name as c:v.
        if let Some(ref name) = ser.name {
            writer
                .write_event(Event::Start(BytesStart::new("c:tx")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("c:strRef")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("c:f")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::Text(BytesText::new(
                    crate::ooxml::sanitize_xml_text(name).as_ref(),
                )))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:f")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:strRef")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:tx")))
                .map_err(xml_write_err)?;
        } else if let Some(ref literal) = ser.name_literal {
            writer
                .write_event(Event::Start(BytesStart::new("c:tx")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("c:v")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::Text(BytesText::new(
                    crate::ooxml::sanitize_xml_text(literal).as_ref(),
                )))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:v")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:tx")))
                .map_err(xml_write_err)?;
        }

        // Children follow the CT_*Ser sequences of dml-chart.xsd: spPr, marker or
        // explosion, dLbls, trendline, errBars, cat/val (xVal/yVal/bubbleSize), smooth.
        // User-set fields the schema does not allow for this chart type are dropped.
        let kind = self.chart_type;
        let has_marker = matches!(
            kind,
            ChartType::Line | ChartType::Stock | ChartType::Scatter | ChartType::Radar
        );
        let has_smooth = matches!(
            kind,
            ChartType::Line | ChartType::Stock | ChartType::Scatter
        );
        let has_explosion = matches!(kind, ChartType::Pie | ChartType::Doughnut);
        let has_fit = !matches!(
            kind,
            ChartType::Pie | ChartType::Doughnut | ChartType::Radar
        );
        let drop_field = |set: bool, field: &str| {
            if set {
                warn!(
                    "chart series {}: `{field}` is not valid for a {kind:?} chart and was not written",
                    ser.idx
                );
            }
        };

        // <c:spPr>
        Self::write_sp_pr(writer, ser)?;

        // <c:marker>
        if has_marker {
            if let Some(marker) = ser.marker {
                Self::write_marker(writer, marker)?;
            }
        } else {
            drop_field(ser.marker.is_some(), "marker");
        }

        // Explosion (pie/doughnut)
        if has_explosion {
            if let Some(explosion) = ser.explosion {
                let mut exp = BytesStart::new("c:explosion");
                exp.push_attribute(("val", ibuf.format(explosion)));
                writer
                    .write_event(Event::Empty(exp))
                    .map_err(xml_write_err)?;
            }
        } else {
            drop_field(ser.explosion.is_some(), "explosion");
        }

        // Series-level data labels
        if let Some(ref dl) = ser.data_labels {
            Self::write_data_labels(writer, dl)?;
        }

        if has_fit {
            // Trendline
            if let Some(ref tl) = ser.trendline {
                Self::write_trendline(writer, tl, ibuf)?;
            }

            // Error bars
            if let Some(ref eb) = ser.error_bars {
                Self::write_error_bars(writer, eb)?;
            }
        } else {
            drop_field(ser.trendline.is_some(), "trendline");
            drop_field(ser.error_bars.is_some(), "errorBars");
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
        if has_smooth {
            if let Some(smooth) = ser.smooth {
                Self::write_bool_element(writer, "c:smooth", smooth)?;
            }
        } else {
            drop_field(ser.smooth.is_some(), "smooth");
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:ser")))
            .map_err(xml_write_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Shape properties (fill + line)
    // -----------------------------------------------------------------------

    fn write_sp_pr(writer: &mut Writer<&mut Vec<u8>>, ser: &ChartSeries) -> Result<()> {
        if ser.fill_color.is_none() && ser.line_color.is_none() {
            return Ok(());
        }

        writer
            .write_event(Event::Start(BytesStart::new("c:spPr")))
            .map_err(xml_write_err)?;

        // Fill
        if let Some(ref fill) = ser.fill_color {
            writer
                .write_event(Event::Start(BytesStart::new("a:solidFill")))
                .map_err(xml_write_err)?;
            let mut clr = BytesStart::new("a:srgbClr");
            clr.push_attribute(xml_attr("val", fill));
            writer
                .write_event(Event::Empty(clr))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:solidFill")))
                .map_err(xml_write_err)?;
        }

        // Line
        if let Some(ref line_color) = ser.line_color {
            let mut ln = BytesStart::new("a:ln");
            if let Some(w) = ser.line_width {
                let mut ibuf = itoa::Buffer::new();
                ln.push_attribute(("w", ibuf.format(w)));
            }
            writer
                .write_event(Event::Start(ln))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("a:solidFill")))
                .map_err(xml_write_err)?;
            let mut clr = BytesStart::new("a:srgbClr");
            clr.push_attribute(xml_attr("val", line_color));
            writer
                .write_event(Event::Empty(clr))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:solidFill")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:ln")))
                .map_err(xml_write_err)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:spPr")))
            .map_err(xml_write_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Marker
    // -----------------------------------------------------------------------

    fn write_marker(writer: &mut Writer<&mut Vec<u8>>, marker: MarkerStyle) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:marker")))
            .map_err(xml_write_err)?;

        let mut sym = BytesStart::new("c:symbol");
        sym.push_attribute(("val", marker.xml_val()));
        writer
            .write_event(Event::Empty(sym))
            .map_err(xml_write_err)?;

        let mut sz = BytesStart::new("c:size");
        sz.push_attribute(("val", "5"));
        writer
            .write_event(Event::Empty(sz))
            .map_err(xml_write_err)?;

        writer
            .write_event(Event::End(BytesEnd::new("c:marker")))
            .map_err(xml_write_err)?;

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
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::Start(BytesStart::new(ref_tag)))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::Start(BytesStart::new("c:f")))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::Text(BytesText::new(
                crate::ooxml::sanitize_xml_text(formula).as_ref(),
            )))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("c:f")))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::End(BytesEnd::new(ref_tag)))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::End(BytesEnd::new(outer_tag)))
            .map_err(xml_write_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Title
    // -----------------------------------------------------------------------

    fn write_title(writer: &mut Writer<&mut Vec<u8>>, title: &ChartTitle) -> Result<()> {
        let mut ibuf = itoa::Buffer::new();

        writer
            .write_event(Event::Start(BytesStart::new("c:title")))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::Start(BytesStart::new("c:tx")))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::Start(BytesStart::new("c:rich")))
            .map_err(xml_write_err)?;

        // <a:bodyPr/>
        writer
            .write_event(Event::Empty(BytesStart::new("a:bodyPr")))
            .map_err(xml_write_err)?;
        // <a:lstStyle/>
        writer
            .write_event(Event::Empty(BytesStart::new("a:lstStyle")))
            .map_err(xml_write_err)?;

        // <a:p>
        writer
            .write_event(Event::Start(BytesStart::new("a:p")))
            .map_err(xml_write_err)?;

        // <a:pPr>
        writer
            .write_event(Event::Start(BytesStart::new("a:pPr")))
            .map_err(xml_write_err)?;

        // <a:defRPr>
        let mut def_rpr = BytesStart::new("a:defRPr");
        if let Some(font_size) = title.font_size {
            def_rpr.push_attribute(("sz", ibuf.format(font_size)));
        }
        if let Some(bold) = title.bold {
            def_rpr.push_attribute(("b", if bold { "1" } else { "0" }));
        }
        if let Some(ref color) = title.color {
            writer
                .write_event(Event::Start(def_rpr))
                .map_err(xml_write_err)?;
            Self::write_solid_fill(writer, color)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:defRPr")))
                .map_err(xml_write_err)?;
        } else {
            writer
                .write_event(Event::Empty(def_rpr))
                .map_err(xml_write_err)?;
        }

        // </a:pPr>
        writer
            .write_event(Event::End(BytesEnd::new("a:pPr")))
            .map_err(xml_write_err)?;

        // <a:r>
        writer
            .write_event(Event::Start(BytesStart::new("a:r")))
            .map_err(xml_write_err)?;

        // <a:rPr>
        let mut rpr = BytesStart::new("a:rPr");
        rpr.push_attribute(("lang", "en-US"));
        if let Some(font_size) = title.font_size {
            rpr.push_attribute(("sz", ibuf.format(font_size)));
        }
        if let Some(bold) = title.bold {
            rpr.push_attribute(("b", if bold { "1" } else { "0" }));
        }
        if let Some(ref color) = title.color {
            writer
                .write_event(Event::Start(rpr))
                .map_err(xml_write_err)?;
            Self::write_solid_fill(writer, color)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:rPr")))
                .map_err(xml_write_err)?;
        } else {
            writer
                .write_event(Event::Empty(rpr))
                .map_err(xml_write_err)?;
        }

        // <a:t>text</a:t>
        writer
            .write_event(Event::Start(BytesStart::new("a:t")))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::Text(BytesText::new(
                crate::ooxml::sanitize_xml_text(&title.text).as_ref(),
            )))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("a:t")))
            .map_err(xml_write_err)?;

        // </a:r>
        writer
            .write_event(Event::End(BytesEnd::new("a:r")))
            .map_err(xml_write_err)?;

        // </a:p>
        writer
            .write_event(Event::End(BytesEnd::new("a:p")))
            .map_err(xml_write_err)?;

        // </c:rich>
        writer
            .write_event(Event::End(BytesEnd::new("c:rich")))
            .map_err(xml_write_err)?;
        // </c:tx>
        writer
            .write_event(Event::End(BytesEnd::new("c:tx")))
            .map_err(xml_write_err)?;

        // <c:overlay val="0"/>
        let mut overlay = BytesStart::new("c:overlay");
        overlay.push_attribute(("val", if title.overlay { "1" } else { "0" }));
        writer
            .write_event(Event::Empty(overlay))
            .map_err(xml_write_err)?;

        // </c:title>
        writer
            .write_event(Event::End(BytesEnd::new("c:title")))
            .map_err(xml_write_err)?;

        Ok(())
    }

    fn write_solid_fill(writer: &mut Writer<&mut Vec<u8>>, color: &str) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("a:solidFill")))
            .map_err(xml_write_err)?;
        let mut clr = BytesStart::new("a:srgbClr");
        clr.push_attribute(xml_attr("val", color));
        writer
            .write_event(Event::Empty(clr))
            .map_err(xml_write_err)?;
        writer
            .write_event(Event::End(BytesEnd::new("a:solidFill")))
            .map_err(xml_write_err)?;
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Data labels
    // -----------------------------------------------------------------------

    fn write_data_labels(writer: &mut Writer<&mut Vec<u8>>, dl: &DataLabels) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:dLbls")))
            .map_err(xml_write_err)?;

        // Group_DLbls order: numFmt, showLegendKey, showVal, showCatName, showSerName,
        // showPercent, showBubbleSize, then showLeaderLines.
        if let Some(ref fmt) = dl.num_fmt {
            let mut nf = BytesStart::new("c:numFmt");
            nf.push_attribute(xml_attr("formatCode", fmt));
            nf.push_attribute(("sourceLinked", "0"));
            writer
                .write_event(Event::Empty(nf))
                .map_err(xml_write_err)?;
        }

        if let Some(v) = dl.show_legend_key {
            Self::write_bool_element(writer, "c:showLegendKey", v)?;
        }
        Self::write_bool_element(writer, "c:showVal", dl.show_val)?;
        Self::write_bool_element(writer, "c:showCatName", dl.show_cat_name)?;
        Self::write_bool_element(writer, "c:showSerName", dl.show_ser_name)?;
        Self::write_bool_element(writer, "c:showPercent", dl.show_percent)?;
        if let Some(v) = dl.show_bubble_size {
            Self::write_bool_element(writer, "c:showBubbleSize", v)?;
        }
        Self::write_bool_element(writer, "c:showLeaderLines", dl.show_leader_lines)?;

        writer
            .write_event(Event::End(BytesEnd::new("c:dLbls")))
            .map_err(xml_write_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Axis
    // -----------------------------------------------------------------------

    /// Write `axis` as a `tag` element (`c:catAx` or `c:valAx`); `default_pos` is
    /// the `c:axPos` written when `axis.position` is unset.
    fn write_axis(
        writer: &mut Writer<&mut Vec<u8>>,
        tag: &str,
        axis: &ChartAxis,
        default_pos: &'static str,
    ) -> Result<()> {
        let mut ibuf = itoa::Buffer::new();

        writer
            .write_event(Event::Start(BytesStart::new(tag)))
            .map_err(xml_write_err)?;

        // <c:axId>
        let mut ax_id = BytesStart::new("c:axId");
        ax_id.push_attribute(("val", ibuf.format(axis.id)));
        writer
            .write_event(Event::Empty(ax_id))
            .map_err(xml_write_err)?;

        // <c:scaling>: logBase, orientation, max, min.
        writer
            .write_event(Event::Start(BytesStart::new("c:scaling")))
            .map_err(xml_write_err)?;

        if let Some(log_base) = axis.log_base {
            Self::write_f64_element(writer, "c:logBase", log_base)?;
        }

        let mut orient = BytesStart::new("c:orientation");
        orient.push_attribute(("val", if axis.reversed { "maxMin" } else { "minMax" }));
        writer
            .write_event(Event::Empty(orient))
            .map_err(xml_write_err)?;

        if let Some(max) = axis.max {
            Self::write_f64_element(writer, "c:max", max)?;
        }
        if let Some(min) = axis.min {
            Self::write_f64_element(writer, "c:min", min)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:scaling")))
            .map_err(xml_write_err)?;

        // <c:delete>
        Self::write_bool_element(writer, "c:delete", axis.delete)?;

        // <c:axPos>
        let pos = axis
            .position
            .map_or(default_pos, super::types::AxisPosition::xml_val);
        let mut ax_pos = BytesStart::new("c:axPos");
        ax_pos.push_attribute(("val", pos));
        writer
            .write_event(Event::Empty(ax_pos))
            .map_err(xml_write_err)?;

        // Gridlines, then the axis title (EG_AxShared order).
        if axis.major_gridlines {
            writer
                .write_event(Event::Empty(BytesStart::new("c:majorGridlines")))
                .map_err(xml_write_err)?;
        }
        if axis.minor_gridlines {
            writer
                .write_event(Event::Empty(BytesStart::new("c:minorGridlines")))
                .map_err(xml_write_err)?;
        }
        if let Some(ref title) = axis.title {
            Self::write_title(writer, title)?;
        }

        // Number format
        {
            let fmt_code = axis.num_fmt.as_deref().unwrap_or("General");
            let src_linked = if axis.source_linked { "1" } else { "0" };
            let mut nf = BytesStart::new("c:numFmt");
            nf.push_attribute(xml_attr("formatCode", fmt_code));
            nf.push_attribute(("sourceLinked", src_linked));
            writer
                .write_event(Event::Empty(nf))
                .map_err(xml_write_err)?;
        }

        // Tick marks
        if let Some(tm) = axis.major_tick_mark {
            let mut el = BytesStart::new("c:majorTickMark");
            el.push_attribute(("val", tm.xml_val()));
            writer
                .write_event(Event::Empty(el))
                .map_err(xml_write_err)?;
        }
        if let Some(tm) = axis.minor_tick_mark {
            let mut el = BytesStart::new("c:minorTickMark");
            el.push_attribute(("val", tm.xml_val()));
            writer
                .write_event(Event::Empty(el))
                .map_err(xml_write_err)?;
        }

        // Tick label position
        if let Some(tlp) = axis.tick_lbl_pos {
            let mut el = BytesStart::new("c:tickLblPos");
            el.push_attribute(("val", tlp.xml_val()));
            writer
                .write_event(Event::Empty(el))
                .map_err(xml_write_err)?;
        }

        // <c:txPr> — axis tick label font size
        if let Some(sz) = axis.font_size {
            writer
                .write_event(Event::Start(BytesStart::new("c:txPr")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::Empty(BytesStart::new("a:bodyPr")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::Empty(BytesStart::new("a:lstStyle")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("a:p")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::Start(BytesStart::new("a:pPr")))
                .map_err(xml_write_err)?;
            let mut def_rpr = BytesStart::new("a:defRPr");
            def_rpr.push_attribute(("sz", ibuf.format(sz)));
            writer
                .write_event(Event::Empty(def_rpr))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:pPr")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:p")))
                .map_err(xml_write_err)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:txPr")))
                .map_err(xml_write_err)?;
        }

        // <c:crossAx>
        let mut cross = BytesStart::new("c:crossAx");
        cross.push_attribute(("val", ibuf.format(axis.cross_ax)));
        writer
            .write_event(Event::Empty(cross))
            .map_err(xml_write_err)?;

        // Crosses at
        if let Some(crosses_at) = axis.crosses_at {
            Self::write_f64_element(writer, "c:crossesAt", crosses_at)?;
        }

        // crossBetween, majorUnit and minorUnit exist on CT_ValAx only.
        if tag == "c:valAx" {
            if let Some(cb) = axis.cross_between {
                let mut el = BytesStart::new("c:crossBetween");
                el.push_attribute(("val", cb.xml_val()));
                writer
                    .write_event(Event::Empty(el))
                    .map_err(xml_write_err)?;
            }
            if let Some(major) = axis.major_unit {
                Self::write_f64_element(writer, "c:majorUnit", major)?;
            }
            if let Some(minor) = axis.minor_unit {
                Self::write_f64_element(writer, "c:minorUnit", minor)?;
            }
        }

        writer
            .write_event(Event::End(BytesEnd::new(tag)))
            .map_err(xml_write_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Legend
    // -----------------------------------------------------------------------

    fn write_legend(writer: &mut Writer<&mut Vec<u8>>, legend: &ChartLegend) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:legend")))
            .map_err(xml_write_err)?;

        let mut pos = BytesStart::new("c:legendPos");
        pos.push_attribute(("val", legend.position.xml_val()));
        writer
            .write_event(Event::Empty(pos))
            .map_err(xml_write_err)?;

        let mut overlay = BytesStart::new("c:overlay");
        overlay.push_attribute(("val", if legend.overlay { "1" } else { "0" }));
        writer
            .write_event(Event::Empty(overlay))
            .map_err(xml_write_err)?;

        writer
            .write_event(Event::End(BytesEnd::new("c:legend")))
            .map_err(xml_write_err)?;

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
            .map_err(xml_write_err)?;

        let mut tt = BytesStart::new("c:trendlineType");
        tt.push_attribute(("val", tl.trend_type.xml_val()));
        writer
            .write_event(Event::Empty(tt))
            .map_err(xml_write_err)?;

        if let Some(order) = tl.order {
            let mut el = BytesStart::new("c:order");
            el.push_attribute(("val", ibuf.format(order)));
            writer
                .write_event(Event::Empty(el))
                .map_err(xml_write_err)?;
        }
        if let Some(period) = tl.period {
            let mut el = BytesStart::new("c:period");
            el.push_attribute(("val", ibuf.format(period)));
            writer
                .write_event(Event::Empty(el))
                .map_err(xml_write_err)?;
        }
        if let Some(fwd) = tl.forward {
            Self::write_f64_element(writer, "c:forward", fwd)?;
        }
        if let Some(bwd) = tl.backward {
            Self::write_f64_element(writer, "c:backward", bwd)?;
        }
        if tl.display_r_sqr {
            Self::write_bool_element(writer, "c:dispRSqr", true)?;
        }
        if tl.display_eq {
            Self::write_bool_element(writer, "c:dispEq", true)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:trendline")))
            .map_err(xml_write_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Error bars
    // -----------------------------------------------------------------------

    fn write_error_bars(writer: &mut Writer<&mut Vec<u8>>, eb: &ErrorBars) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:errBars")))
            .map_err(xml_write_err)?;

        // CT_ErrBars order: errDir, errBarType, errValType, plus, minus, val.
        if let Some(axis) = eb.err_dir {
            let mut el = BytesStart::new("c:errDir");
            el.push_attribute(("val", axis.xml_val()));
            writer
                .write_event(Event::Empty(el))
                .map_err(xml_write_err)?;
        }

        let mut dir = BytesStart::new("c:errBarType");
        dir.push_attribute(("val", eb.direction.xml_val()));
        writer
            .write_event(Event::Empty(dir))
            .map_err(xml_write_err)?;

        let mut vt = BytesStart::new("c:errValType");
        vt.push_attribute(("val", eb.err_type.xml_val()));
        writer
            .write_event(Event::Empty(vt))
            .map_err(xml_write_err)?;

        if let Some(ref plus) = eb.plus_ref {
            Self::write_ref_element(writer, "c:plus", "c:numRef", plus)?;
        }
        if let Some(ref minus) = eb.minus_ref {
            Self::write_ref_element(writer, "c:minus", "c:numRef", minus)?;
        }

        if let Some(val) = eb.value {
            Self::write_f64_element(writer, "c:val", val)?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("c:errBars")))
            .map_err(xml_write_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // View 3D
    // -----------------------------------------------------------------------

    fn write_view_3d(writer: &mut Writer<&mut Vec<u8>>, v: &View3D) -> Result<()> {
        let mut ibuf = itoa::Buffer::new();

        writer
            .write_event(Event::Start(BytesStart::new("c:view3D")))
            .map_err(xml_write_err)?;

        if let Some(rx) = v.rot_x {
            let mut el = BytesStart::new("c:rotX");
            el.push_attribute(("val", ibuf.format(rx)));
            writer
                .write_event(Event::Empty(el))
                .map_err(xml_write_err)?;
        }
        if let Some(ry) = v.rot_y {
            let mut el = BytesStart::new("c:rotY");
            el.push_attribute(("val", ibuf.format(ry)));
            writer
                .write_event(Event::Empty(el))
                .map_err(xml_write_err)?;
        }
        if let Some(ra) = v.r_ang_ax {
            let mut el = BytesStart::new("c:rAngAx");
            el.push_attribute(("val", if ra { "1" } else { "0" }));
            writer
                .write_event(Event::Empty(el))
                .map_err(xml_write_err)?;
        }

        if let Some(p) = v.perspective {
            let mut el = BytesStart::new("c:perspective");
            el.push_attribute(("val", ibuf.format(p)));
            writer
                .write_event(Event::Empty(el))
                .map_err(xml_write_err)?;
        }
        writer
            .write_event(Event::End(BytesEnd::new("c:view3D")))
            .map_err(xml_write_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Print settings (always present)
    // -----------------------------------------------------------------------

    fn write_print_settings(writer: &mut Writer<&mut Vec<u8>>) -> Result<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:printSettings")))
            .map_err(xml_write_err)?;

        // <c:headerFooter/>
        writer
            .write_event(Event::Empty(BytesStart::new("c:headerFooter")))
            .map_err(xml_write_err)?;

        // <c:pageMargins>
        let mut pm = BytesStart::new("c:pageMargins");
        pm.push_attribute(("b", "0.75"));
        pm.push_attribute(("l", "0.7"));
        pm.push_attribute(("r", "0.7"));
        pm.push_attribute(("t", "0.75"));
        pm.push_attribute(("header", "0.3"));
        pm.push_attribute(("footer", "0.3"));
        writer
            .write_event(Event::Empty(pm))
            .map_err(xml_write_err)?;

        // <c:pageSetup/>
        writer
            .write_event(Event::Empty(BytesStart::new("c:pageSetup")))
            .map_err(xml_write_err)?;

        writer
            .write_event(Event::End(BytesEnd::new("c:printSettings")))
            .map_err(xml_write_err)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Utility helpers
    // -----------------------------------------------------------------------

    fn write_bool_element(writer: &mut Writer<&mut Vec<u8>>, tag: &str, value: bool) -> Result<()> {
        let mut el = BytesStart::new(tag);
        el.push_attribute(("val", if value { "1" } else { "0" }));
        writer
            .write_event(Event::Empty(el))
            .map_err(xml_write_err)?;
        Ok(())
    }

    fn write_f64_element(writer: &mut Writer<&mut Vec<u8>>, tag: &str, value: f64) -> Result<()> {
        if !value.is_finite() {
            return Ok(()); // Skip NaN/Infinity — not valid xsd:double in OOXML
        }
        let mut el = BytesStart::new(tag);
        let mut xsd = XsdBuf::new();
        el.push_attribute(("val", fmt_xsd_double(value, &mut xsd, tag)?));
        writer
            .write_event(Event::Empty(el))
            .map_err(xml_write_err)?;
        Ok(())
    }
}
