//! Chart definitions — `xl/charts/chart{n}.xml` and the drawing anchors that place
//! charts on a worksheet.
//!
//! - [`ChartData::parse`] reads a chart part and [`ChartData::to_xml`] writes one.
//! - [`parse_drawing_anchors`] reads the chart anchors of `xl/drawings/drawing{n}.xml`
//!   into [`DrawingAnchors`] and [`ChartAnchor::generate_drawing_xml`] writes a
//!   worksheet's drawing part.
//! - [`WorksheetChart`] pairs a [`ChartData`] with its [`ChartAnchor`].
//!
//! The model covers the bar, column, line, pie, doughnut, scatter, area, radar,
//! bubble and stock chart types, with at most one secondary chart-type element
//! (a combo chart). A part that holds a 3-D, surface or bar/pie-of-pie chart
//! type, a third chart-type element, a date or series axis, an axis beyond the
//! modelled ones, `c:userShapes` or `c:externalData` still parses; the parser
//! keeps it verbatim in [`ChartData::raw_xml`] and the writer emits those bytes.
//! Other elements without a model field (e.g. `c:gapWidth`, `c:varyColors`,
//! `c:upDownBars`) are not read or written back.

mod parser;
mod types;
mod writer;

pub use parser::{DrawingAnchors, parse_drawing_anchors};
pub use types::*;

use serde::{Deserialize, Serialize};

/// A chart embedded in a worksheet, with its anchor position.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorksheetChart {
    pub chart: ChartData,
    pub anchor: ChartAnchor,
}

#[cfg(test)]
mod tests;
