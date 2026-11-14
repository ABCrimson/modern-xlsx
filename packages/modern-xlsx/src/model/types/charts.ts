// ---------------------------------------------------------------------------
// Chart definitions
// ---------------------------------------------------------------------------

/**
 * Chart kind, written as the chart-type element `c:{kind}Chart`, except that `'bar'` and
 * `'column'` are both `c:barChart`: horizontal bars for `'bar'`, vertical bars for
 * `'column'` unless {@link ChartDataModel.barDirHorizontal} is `true`.
 */
export type ChartType =
  | 'bar'
  | 'column'
  | 'line'
  | 'pie'
  | 'doughnut'
  | 'scatter'
  | 'area'
  | 'radar'
  | 'bubble'
  | 'stock';
/**
 * `c:grouping` of a bar, column, line or area chart. `'clustered'` is valid for bar and
 * column charts only.
 */
export type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';
/** `c:scatterStyle` of a scatter chart. */
export type ScatterStyle = 'lineMarker' | 'line' | 'marker' | 'smooth' | 'smoothMarker';
/** `c:radarStyle` of a radar chart. */
export type RadarStyle = 'standard' | 'marker' | 'filled';
/** `c:symbol` of a series marker. */
export type MarkerStyleType =
  | 'circle'
  | 'square'
  | 'diamond'
  | 'triangle'
  | 'star'
  | 'x'
  | 'plus'
  | 'dash'
  | 'dot'
  | 'none';
/** `c:tickLblPos` of an axis. */
export type TickLabelPosition = 'high' | 'low' | 'nextTo' | 'none';
/** `c:majorTickMark` or `c:minorTickMark` of an axis. */
export type TickMark = 'cross' | 'in' | 'out' | 'none';
/** `c:axPos` of an axis. */
export type AxisPosition = 'bottom' | 'top' | 'left' | 'right';
/** `c:legendPos` of a legend; `'topRight'` is written as `tr`. */
export type LegendPosition = 'top' | 'bottom' | 'left' | 'right' | 'topRight';

/** `c:trendlineType` of a trendline. */
export type TrendlineType =
  | 'linear'
  | 'exponential'
  | 'logarithmic'
  | 'polynomial'
  | 'power'
  | 'movingAverage';
/** `c:errValType` of error bars. */
export type ErrorBarType = 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr' | 'custom';
/** `c:errBarType` of error bars: which side of the value they extend to. */
export type ErrorBarDirection = 'both' | 'plus' | 'minus';

/**
 * A series `c:trendline`. Written for every chart type except pie, doughnut and radar,
 * which drop it.
 */
export interface TrendlineData {
  trendType: TrendlineType;
  /** Polynomial order, 2 to 6; `validateChartData` checks it on the top-level chart's series. */
  order?: number | null;
  /** Moving-average period, 2 or more; `validateChartData` checks it on the top-level chart's series. */
  period?: number | null;
  /** Periods to forecast forward. */
  forward?: number | null;
  /** Periods to forecast backward. */
  backward?: number | null;
  /** Show the trendline's equation on the chart. */
  displayEq?: boolean;
  /** Show the R-squared value on the chart. */
  displayRSqr?: boolean;
}

/**
 * Series `c:errBars`. Written for every chart type except pie, doughnut and radar, which
 * drop them.
 */
export interface ErrorBarsData {
  /** `c:errDir`: the axis the bars follow (`x` or `y`). */
  errDir?: 'x' | 'y' | null;
  errType: ErrorBarType;
  /** Omitted means `'both'`. */
  direction?: ErrorBarDirection;
  /** `c:val`, written whenever set: the amount for `errType` `'fixedVal'` or `'percentage'`. */
  value?: number | null;
  /** Range of the positive custom error amounts (`c:plus`), used with `errType: 'custom'`. */
  plusRef?: string | null;
  /** Range of the negative custom error amounts (`c:minus`), used with `errType: 'custom'`. */
  minusRef?: string | null;
}

/** The chart's `c:view3D` rotation and perspective. */
export interface View3DData {
  /** X rotation, -90 to 90 (checked by `validateChartData`). */
  rotX?: number | null;
  /** Y rotation, 0 to 360 (checked by `validateChartData`). */
  rotY?: number | null;
  /** Perspective, 0 to 240 (checked by `validateChartData`). */
  perspective?: number | null;
  /** Right-angle axes. */
  rAngAx?: boolean | null;
}

/**
 * A chart or axis title (`c:title`), written as rich text with one run. On read the
 * text of all its runs is joined, and a title without text reads as absent.
 */
export interface ChartTitleData {
  text: string;
  /** `c:overlay`: `true` lets the title overlap the plot area. Omitted means `false`. */
  overlay?: boolean;
  /**
   * Font size in hundredths of a point (1400 = 14 pt), written verbatim as the DrawingML
   * `sz` of the title text. `validateChartData` checks the chart title and the `catAxis`
   * and `valAxis` titles against 100 to 400000.
   */
  fontSize?: number | null;
  /** Bold title text, written as `b`; unset writes no `b`. */
  bold?: boolean | null;
  /**
   * Text colour as six hex digits `RRGGBB` without `#` (e.g. `'FF0000'`), written verbatim
   * as the DrawingML `a:srgbClr val`.
   */
  color?: string | null;
}

/** A chart axis, `c:catAx` or `c:valAx`. */
export interface ChartAxisData {
  /** The axis id (`c:axId`). */
  id: number;
  /** The `id` of the axis this one crosses (`c:crossAx`). */
  crossAx: number;
  title?: ChartTitleData | null;
  numFmt?: string | null;
  sourceLinked?: boolean;
  min?: number | null;
  max?: number | null;
  majorUnit?: number | null;
  minorUnit?: number | null;
  logBase?: number | null;
  reversed?: boolean;
  tickLblPos?: TickLabelPosition | null;
  majorTickMark?: TickMark | null;
  minorTickMark?: TickMark | null;
  majorGridlines?: boolean;
  minorGridlines?: boolean;
  delete?: boolean;
  position?: AxisPosition | null;
  crossesAt?: number | null;
  /** Font size for tick labels in hundredths of a point (1400 = 14pt). */
  fontSize?: number | null;
  /** `c:crossBetween` (`c:valAx` only): cross between categories or at their midpoint. */
  crossBetween?: 'between' | 'midCat' | null;
}

/** The chart's `c:legend`. */
export interface ChartLegendData {
  position: LegendPosition;
  /** `c:overlay`: `true` lets the legend overlap the plot area. Omitted means `false`. */
  overlay?: boolean;
}

/** Data labels (`c:dLbls`) of a chart or a series. */
export interface DataLabelsData {
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  numFmt?: string | null;
  showLeaderLines?: boolean;
  /** `c:showLegendKey`: show the legend key beside each label. */
  showLegendKey?: boolean | null;
  /** `c:showBubbleSize`: show the bubble size in each label. */
  showBubbleSize?: boolean | null;
}

/** A chart series (`c:ser`): its data references, name and formatting. */
export interface ChartSeriesData {
  /**
   * `c:idx`, written and read as given. Excel opens a chart whose series share it;
   * `ChartBuilder` still numbers it across the primary and secondary chart.
   */
  idx: number;
  /**
   * `c:order`, the series' place in the plotting order. It must be unique across the
   * series of the chart and its `secondaryChart`: Excel refuses a file where two share it,
   * so `validateChartData` throws on such a chart, `Workbook.validate()` reports it and the
   * write fails with `INVALID_WORKBOOK`. Written and read as given, never renumbered.
   */
  order: number;
  /**
   * Series name as a cell-reference formula (`c:tx/c:strRef/c:f`, for example
   * `'Sheet1!$B$1'`), never a literal. Written in preference to `nameLiteral`.
   */
  name?: string | null;
  /** Series name as literal text (`c:tx/c:v`), written when `name` is unset. */
  nameLiteral?: string | null;
  catRef?: string | null;
  valRef: string;
  xValRef?: string | null;
  bubbleSizeRef?: string | null;
  /** Fill colour as six hex digits `RRGGBB` without `#`, written verbatim as `a:srgbClr val`. */
  fillColor?: string | null;
  /** Line colour as six hex digits `RRGGBB` without `#`, written verbatim as `a:srgbClr val`. */
  lineColor?: string | null;
  /** Line width in EMU (12700 per point), written only with `lineColor`. */
  lineWidth?: number | null;
  /** Written for line, stock, scatter and radar charts only; other types drop it. */
  marker?: MarkerStyleType | null;
  /** Written for line, stock and scatter charts only; other types drop it. */
  smooth?: boolean | null;
  /** `c:explosion`, written for pie and doughnut charts only; other types drop it. */
  explosion?: number | null;
  dataLabels?: DataLabelsData | null;
  trendline?: TrendlineData | null;
  errorBars?: ErrorBarsData | null;
}

/**
 * A `c:manualLayout` position and size, each a fraction (0.0-1.0) of the chart's width
 * (`x`, `w`) or height (`y`, `h`).
 */
export interface ManualLayoutData {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** A chart part, `xl/charts/chart{n}.xml`: its type, series, axes, legend and options. */
export interface ChartDataModel {
  chartType: ChartType;
  title?: ChartTitleData | null;
  readonly series: ChartSeriesData[];
  catAxis?: ChartAxisData | null;
  valAxis?: ChartAxisData | null;
  legend?: ChartLegendData | null;
  dataLabels?: DataLabelsData | null;
  grouping?: ChartGrouping | null;
  scatterStyle?: ScatterStyle | null;
  radarStyle?: RadarStyle | null;
  /** `c:holeSize` of a doughnut chart (ignored for other types), 1 to 90 (checked by `validateChartData`). */
  holeSize?: number | null;
  /** `true` writes a `'column'` chart with horizontal bars (`c:barDir` `bar`); a `'bar'` chart is always horizontal. */
  barDirHorizontal?: boolean | null;
  /** `c:style`, 1 to 48 (checked by `validateChartData`). */
  styleId?: number | null;
  /** Manual layout of the plot area. */
  plotAreaLayout?: ManualLayoutData | null;
  /**
   * Secondary chart of a combo chart: one chart-type element written after the
   * primary one, on `secondaryCatAxis` and `secondaryValAxis`. An unset slot
   * shares the primary chart's axis in that slot (`catAxis` / `valAxis`; for a
   * scatter or bubble primary read from a file, `valAxis` / `secondaryValAxis`),
   * else a default axis the writer adds, also under a pie or doughnut primary,
   * which stays axis-free. Its own axis fields are not written. It cannot nest: a
   * `secondaryChart` with a `secondaryChart` of its own makes the write throw
   * (`XML_WRITE`). Its series share one `order` numbering with the primary chart's
   * (see {@link ChartSeriesData.order}); `ChartBuilder.secondaryChart` continues it.
   */
  secondaryChart?: ChartDataModel | null;
  /** Secondary value axis of a combo chart. */
  secondaryValAxis?: ChartAxisData | null;
  /** Secondary category axis of a combo chart, written as a second `c:catAx`. */
  secondaryCatAxis?: ChartAxisData | null;
  /** `true` writes a data table (`c:dTable`, with legend keys) in the plot area. */
  showDataTable?: boolean;
  view3d?: View3DData | null;
  /**
   * `c:autoTitleDeleted`. When unset the writer emits `1` for a chart without a
   * title and `0` for a chart with one, so Excel adds no automatic title.
   */
  autoTitleDeleted?: boolean | null;
  /**
   * The chart part exactly as read, kept when the file holds what the model
   * cannot represent (a 3-D, surface or bar/pie-of-pie chart type, a third
   * chart-type element, an axis the model has no slot for, or a `c:userShapes`
   * or `c:externalData` reference). While it is set the model is ignored on
   * write and the bytes are written unchanged; delete `rawXml` to write the
   * chart from the model instead, and with it the `rawRels` relationships only
   * these bytes referenced.
   */
  rawXml?: string | null;
  /**
   * The chart part's `.rels` as read, for every chart that has one, modelled
   * or kept in `rawXml`. It is written beside the chart at whatever path the
   * chart gets, so the parts it references (styles, colours, user shapes,
   * embedded data) stay linked to their own chart.
   */
  rawRels?: string | null;
}

/**
 * Where a chart sits on the sheet: the `xdr:from` and `xdr:to` cells of its drawing
 * anchor, written verbatim as `xdr:col` and `xdr:row`, so they are 0-based (`0` is column A
 * and row 1). Offsets within a cell are in EMUs; an omitted offset is `0`.
 */
export interface ChartAnchorData {
  fromCol: number;
  fromRow: number;
  fromColOff?: number;
  fromRowOff?: number;
  toCol: number;
  toRow: number;
  toColOff?: number;
  toRowOff?: number;
  /** Width in EMUs, for a oneCellAnchor or absoluteAnchor. With `extCy` also set the anchor is a oneCellAnchor, or an absoluteAnchor when `posX` and `posY` are set too. */
  extCx?: number | null;
  /** Height in EMUs. See `extCx`. */
  extCy?: number | null;
  /** Absolute X position in EMUs. With `posY`, `extCx` and `extCy` the anchor is an absoluteAnchor. */
  posX?: number | null;
  /** Absolute Y position in EMUs. See `posX`. */
  posY?: number | null;
}

/** A chart embedded in a worksheet: the chart and its anchor in the sheet's drawing. */
export interface WorksheetChartData {
  chart: ChartDataModel;
  anchor: ChartAnchorData;
}
