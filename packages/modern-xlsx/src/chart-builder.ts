import type {
  AxisPosition,
  ChartAnchorData,
  ChartAxisData,
  ChartDataModel,
  ChartGrouping,
  ChartLegendData,
  ChartSeriesData,
  ChartTitleData,
  ChartType,
  DataLabelsData,
  ErrorBarDirection,
  ErrorBarsData,
  ErrorBarType,
  LegendPosition,
  ManualLayoutData,
  MarkerStyleType,
  RadarStyle,
  ScatterStyle,
  TickLabelPosition,
  TickMark,
  TrendlineData,
  TrendlineType,
  View3DData,
  WorksheetChartData,
} from './types.js';

/**
 * Options for adding a data series to a chart via {@link ChartBuilder.addSeries}.
 */
export interface AddSeriesOptions {
  /** Display name for the series in the legend. */
  name?: string;
  /** Category axis data reference (e.g., `'Sheet1!$A$2:$A$5'`). */
  catRef?: string;
  /** Value axis data reference (e.g., `'Sheet1!$B$2:$B$5'`). Required. */
  valRef: string;
  /** X-value reference for scatter/bubble charts. */
  xValRef?: string;
  /** Bubble size reference for bubble charts. */
  bubbleSizeRef?: string;
  /** Fill color as hex RGB (e.g., `'4472C4'`). */
  fillColor?: string;
  /** Line/border color as hex RGB. */
  lineColor?: string;
  /** Line width in points. */
  lineWidth?: number;
  /** Marker symbol style (e.g., `'circle'`, `'square'`). */
  marker?: MarkerStyleType;
  /** Enable smooth line rendering for line/scatter charts. */
  smooth?: boolean;
  /** Pie/doughnut slice explosion distance (0-100). */
  explosion?: number;
  /** Per-series data label configuration. */
  dataLabels?: Partial<DataLabelsData>;
  /** Trendline configuration for this series. */
  trendline?: {
    trendType: TrendlineType;
    order?: number;
    period?: number;
    forward?: number;
    backward?: number;
    displayEq?: boolean;
    displayRSqr?: boolean;
  };
  /** Error bars configuration for this series. */
  errorBars?: {
    errType: ErrorBarType;
    direction?: ErrorBarDirection;
    value?: number;
  };
}

/**
 * Options for configuring a chart axis via {@link ChartBuilder.catAxis} or {@link ChartBuilder.valAxis}.
 */
export interface AxisOptions {
  /** Axis title text, or an object with text and formatting. */
  title?: string | { text: string; bold?: boolean; fontSize?: number; color?: string };
  /** Number format code for axis labels (e.g., `'#,##0'`). */
  numFmt?: string;
  /** Whether the number format is linked to the source data. */
  sourceLinked?: boolean;
  /** Minimum axis value (auto-scaled if omitted). */
  min?: number;
  /** Maximum axis value (auto-scaled if omitted). */
  max?: number;
  /** Major tick/gridline interval. */
  majorUnit?: number;
  /** Minor tick/gridline interval. */
  minorUnit?: number;
  /** Logarithmic base (e.g., 10 for log scale). */
  logBase?: number;
  /** Reverse the axis direction. */
  reversed?: boolean;
  /** Tick label position relative to the axis. */
  tickLblPos?: TickLabelPosition;
  /** Major tick mark style. */
  majorTickMark?: TickMark;
  /** Minor tick mark style. */
  minorTickMark?: TickMark;
  /** Show major gridlines. */
  majorGridlines?: boolean;
  /** Show minor gridlines. */
  minorGridlines?: boolean;
  /** Hide the axis entirely. */
  delete?: boolean;
  /** Axis position (left, right, top, bottom). */
  position?: AxisPosition;
  /** Value where the perpendicular axis crosses this axis. */
  crossesAt?: number;
  /** Font size for tick labels in hundredths of a point (1400 = 14pt). */
  fontSize?: number;
}

/**
 * Fluent builder for creating chart definitions.
 *
 * @example
 * ```ts
 * const chart = new ChartBuilder('bar')
 *   .title('Monthly Sales')
 *   .addSeries({ name: 'Revenue', catRef: 'Sheet1!$A$2:$A$5', valRef: 'Sheet1!$B$2:$B$5' })
 *   .legend('bottom')
 *   .anchor({ col: 0, row: 0 }, { col: 10, row: 20 })
 *   .build();
 * ```
 */
export class ChartBuilder {
  private chartType: ChartType;
  private titleData: ChartTitleData | null = null;
  private seriesList: ChartSeriesData[] = [];
  private catAxisData: ChartAxisData | null = null;
  private valAxisData: ChartAxisData | null = null;
  private legendData: ChartLegendData | null = null;
  private dataLabelsData: DataLabelsData | null = null;
  private groupingValue: ChartGrouping | null = null;
  private scatterStyleValue: ScatterStyle | null = null;
  private radarStyleValue: RadarStyle | null = null;
  private holeSizeValue: number | null = null;
  private barDirValue: boolean | null = null;
  private styleIdValue: number | null = null;
  private layoutValue: ManualLayoutData | null = null;
  private anchorData: ChartAnchorData = { fromCol: 0, fromRow: 0, toCol: 10, toRow: 15 };
  private showDataTableValue = false;
  private view3dData: View3DData | null = null;

  /**
   * Create a new ChartBuilder for the given chart type.
   *
   * @param type - The chart type (e.g., `'bar'`, `'line'`, `'pie'`, `'scatter'`).
   */
  constructor(type: ChartType) {
    this.chartType = type;
  }

  /**
   * Set the chart title.
   *
   * @param text - The title text.
   * @param opts - Optional formatting (bold, fontSize, color).
   * @returns `this` for chaining.
   */
  title(text: string, opts?: { bold?: boolean; fontSize?: number; color?: string }): this {
    this.titleData = {
      text,
      overlay: false,
      fontSize: opts?.fontSize ?? null,
      bold: opts?.bold ?? null,
      color: opts?.color ?? null,
    };
    return this;
  }

  /**
   * Add a data series to the chart.
   *
   * @param opts - Series configuration including data references and styling.
   * @returns `this` for chaining.
   *
   * @example
   * ```ts
   * builder.addSeries({
   *   name: 'Revenue',
   *   catRef: 'Sheet1!$A$2:$A$5',
   *   valRef: 'Sheet1!$B$2:$B$5',
   * });
   * ```
   */
  addSeries(opts: AddSeriesOptions): this {
    const idx = this.seriesList.length;
    this.seriesList.push({
      idx,
      order: idx,
      name: opts.name ?? null,
      catRef: opts.catRef ?? null,
      valRef: opts.valRef,
      xValRef: opts.xValRef ?? null,
      bubbleSizeRef: opts.bubbleSizeRef ?? null,
      fillColor: opts.fillColor ?? null,
      lineColor: opts.lineColor ?? null,
      lineWidth: opts.lineWidth ?? null,
      marker: opts.marker ?? null,
      smooth: opts.smooth ?? null,
      explosion: opts.explosion ?? null,
      dataLabels: buildDataLabels(opts.dataLabels),
      trendline: buildTrendline(opts.trendline),
      errorBars: buildErrorBars(opts.errorBars),
    });
    return this;
  }

  /**
   * Configure the category axis (axis id 0).
   *
   * @param opts - Axis title, scale, tick, and gridline options.
   * @returns `this` for chaining.
   */
  catAxis(opts?: AxisOptions): this {
    this.catAxisData = buildAxisData(0, 1, opts);
    return this;
  }

  /**
   * Configure the value axis (axis id 1).
   *
   * @param opts - Axis title, scale, tick, and gridline options.
   * @returns `this` for chaining.
   */
  valAxis(opts?: AxisOptions): this {
    this.valAxisData = buildAxisData(1, 0, opts);
    return this;
  }

  /**
   * Configure the legend.
   *
   * @param position - Legend position (default: `'bottom'`).
   * @param overlay - Whether the legend overlaps the plot area (default: `false`).
   * @returns `this` for chaining.
   */
  legend(position: LegendPosition = 'bottom', overlay = false): this {
    this.legendData = { position, overlay };
    return this;
  }

  /**
   * Configure chart-level data labels.
   *
   * @param opts - Data label visibility options (showVal, showCatName, showPercent, etc.).
   * @returns `this` for chaining.
   */
  dataLabels(opts: Partial<DataLabelsData>): this {
    this.dataLabelsData = normalizeDataLabels(opts);
    return this;
  }

  /**
   * Set the grouping mode (bar/column/line/area).
   *
   * @param g - Grouping type: `'clustered'`, `'stacked'`, `'percentStacked'`, or `'standard'`.
   * @returns `this` for chaining.
   */
  grouping(g: ChartGrouping): this {
    this.groupingValue = g;
    return this;
  }

  /**
   * Set scatter chart style.
   *
   * @param s - Scatter style: `'lineMarker'`, `'line'`, `'marker'`, `'smooth'`, `'smoothMarker'`.
   * @returns `this` for chaining.
   */
  scatterStyle(s: ScatterStyle): this {
    this.scatterStyleValue = s;
    return this;
  }

  /**
   * Set radar chart style.
   *
   * @param s - Radar style: `'standard'`, `'marker'`, `'filled'`.
   * @returns `this` for chaining.
   */
  radarStyle(s: RadarStyle): this {
    this.radarStyleValue = s;
    return this;
  }

  /**
   * Set doughnut hole size (0-90).
   *
   * @param percent - Hole size as a percentage of the chart radius.
   * @returns `this` for chaining.
   */
  holeSize(percent: number): this {
    this.holeSizeValue = percent;
    return this;
  }

  /**
   * Set bar direction. `true` = horizontal bars, `false` = vertical columns.
   *
   * @param horizontal - Whether bars are horizontal.
   * @returns `this` for chaining.
   */
  barDirection(horizontal: boolean): this {
    this.barDirValue = horizontal;
    return this;
  }

  /**
   * Set the chart style ID (1-48, matches Excel built-in styles).
   *
   * @param id - The style ID.
   * @returns `this` for chaining.
   */
  style(id: number): this {
    this.styleIdValue = id;
    return this;
  }

  /**
   * Enable showing the data table below the chart.
   *
   * @param show - Whether to display the data table (default: `true`).
   * @returns `this` for chaining.
   */
  showDataTable(show = true): this {
    this.showDataTableValue = show;
    return this;
  }

  /**
   * Set 3D rotation settings.
   *
   * @param opts - Rotation angles and perspective options.
   * @returns `this` for chaining.
   */
  view3d(opts: { rotX?: number; rotY?: number; perspective?: number; rAngAx?: boolean }): this {
    this.view3dData = {
      rotX: opts.rotX ?? null,
      rotY: opts.rotY ?? null,
      perspective: opts.perspective ?? null,
      rAngAx: opts.rAngAx ?? null,
    };
    return this;
  }

  /**
   * Set the plot area manual layout (fractional 0.0-1.0).
   *
   * @param x - Left offset as a fraction of chart width.
   * @param y - Top offset as a fraction of chart height.
   * @param w - Width as a fraction of chart width.
   * @param h - Height as a fraction of chart height.
   * @returns `this` for chaining.
   */
  plotLayout(x: number, y: number, w: number, h: number): this {
    this.layoutValue = { x, y, w, h };
    return this;
  }

  /**
   * Set the anchor position (cell coordinates).
   *
   * @param from - Top-left anchor cell (col, row, optional EMU offsets).
   * @param to - Bottom-right anchor cell (col, row, optional EMU offsets).
   * @returns `this` for chaining.
   *
   * @example
   * ```ts
   * builder.anchor({ col: 0, row: 0 }, { col: 10, row: 20 });
   * ```
   */
  anchor(
    from: { col: number; row: number; colOff?: number; rowOff?: number },
    to: { col: number; row: number; colOff?: number; rowOff?: number },
  ): this {
    this.anchorData = {
      fromCol: from.col,
      fromRow: from.row,
      fromColOff: from.colOff ?? 0,
      fromRowOff: from.rowOff ?? 0,
      toCol: to.col,
      toRow: to.row,
      toColOff: to.colOff ?? 0,
      toRowOff: to.rowOff ?? 0,
    };
    return this;
  }

  /**
   * Build the final chart data object ready for insertion into a worksheet.
   *
   * @returns A WorksheetChartData containing the chart model and anchor position.
   */
  build(): WorksheetChartData {
    const chart: ChartDataModel = {
      chartType: this.chartType,
      title: this.titleData,
      series: this.seriesList,
      catAxis: this.catAxisData,
      valAxis: this.valAxisData,
      legend: this.legendData,
      dataLabels: this.dataLabelsData,
      grouping: this.groupingValue,
      scatterStyle: this.scatterStyleValue,
      radarStyle: this.radarStyleValue,
      holeSize: this.holeSizeValue,
      barDirHorizontal: this.barDirValue,
      styleId: this.styleIdValue,
      plotAreaLayout: this.layoutValue,
      showDataTable: this.showDataTableValue,
      view3d: this.view3dData,
    };
    return { chart, anchor: this.anchorData };
  }
}

function normalizeDataLabels(opts: Partial<DataLabelsData>): DataLabelsData {
  return {
    showVal: opts.showVal ?? false,
    showCatName: opts.showCatName ?? false,
    showSerName: opts.showSerName ?? false,
    showPercent: opts.showPercent ?? false,
    numFmt: opts.numFmt ?? null,
    showLeaderLines: opts.showLeaderLines ?? false,
  };
}

function buildDataLabels(opts?: Partial<DataLabelsData>): DataLabelsData | null {
  return opts ? normalizeDataLabels(opts) : null;
}

function buildTrendline(opts?: AddSeriesOptions['trendline']): TrendlineData | null {
  if (!opts) return null;
  return {
    trendType: opts.trendType,
    order: opts.order ?? null,
    period: opts.period ?? null,
    forward: opts.forward ?? null,
    backward: opts.backward ?? null,
    displayEq: opts.displayEq ?? false,
    displayRSqr: opts.displayRSqr ?? false,
  };
}

function buildErrorBars(opts?: AddSeriesOptions['errorBars']): ErrorBarsData | null {
  if (!opts) return null;
  return {
    errType: opts.errType,
    direction: opts.direction ?? 'both',
    value: opts.value ?? null,
  };
}

function buildAxisTitle(title: AxisOptions['title']): ChartTitleData | null {
  if (title == null) return null;
  if (typeof title === 'string') return { text: title };
  return {
    text: title.text,
    bold: title.bold ?? null,
    fontSize: title.fontSize ?? null,
    color: title.color ?? null,
  };
}

const AXIS_DEFAULTS: Omit<ChartAxisData, 'id' | 'crossAx'> = {
  title: null,
  numFmt: null,
  sourceLinked: false,
  min: null,
  max: null,
  majorUnit: null,
  minorUnit: null,
  logBase: null,
  reversed: false,
  tickLblPos: null,
  majorTickMark: null,
  minorTickMark: null,
  majorGridlines: false,
  minorGridlines: false,
  delete: false,
  position: null,
  crossesAt: null,
};

function buildAxisScaleProps(
  opts: AxisOptions,
): Pick<
  ChartAxisData,
  'numFmt' | 'sourceLinked' | 'min' | 'max' | 'majorUnit' | 'minorUnit' | 'logBase' | 'reversed'
> {
  return {
    numFmt: opts.numFmt ?? null,
    sourceLinked: opts.sourceLinked ?? false,
    min: opts.min ?? null,
    max: opts.max ?? null,
    majorUnit: opts.majorUnit ?? null,
    minorUnit: opts.minorUnit ?? null,
    logBase: opts.logBase ?? null,
    reversed: opts.reversed ?? false,
  };
}

function buildAxisTickProps(
  opts: AxisOptions,
): Pick<
  ChartAxisData,
  | 'tickLblPos'
  | 'majorTickMark'
  | 'minorTickMark'
  | 'majorGridlines'
  | 'minorGridlines'
  | 'delete'
  | 'position'
  | 'crossesAt'
  | 'fontSize'
> {
  return {
    tickLblPos: opts.tickLblPos ?? null,
    majorTickMark: opts.majorTickMark ?? null,
    minorTickMark: opts.minorTickMark ?? null,
    majorGridlines: opts.majorGridlines ?? false,
    minorGridlines: opts.minorGridlines ?? false,
    delete: opts.delete ?? false,
    position: opts.position ?? null,
    crossesAt: opts.crossesAt ?? null,
    fontSize: opts.fontSize ?? null,
  };
}

function buildAxisData(id: number, crossAx: number, opts?: AxisOptions): ChartAxisData {
  if (!opts) return { ...AXIS_DEFAULTS, id, crossAx };
  return {
    id,
    crossAx,
    title: buildAxisTitle(opts.title),
    ...buildAxisScaleProps(opts),
    ...buildAxisTickProps(opts),
  };
}
