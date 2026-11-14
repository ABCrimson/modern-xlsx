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
} from '../model/types.js';

/**
 * Options for adding a data series to a chart via {@link ChartBuilder.addSeries}.
 */
export interface AddSeriesOptions {
  /**
   * `c:idx`. Omitted, {@link ChartBuilder.build} gives the series the next value from 0 up
   * that no series of this chart or its {@link ChartBuilder.secondaryChart} uses.
   */
  idx?: number;
  /**
   * `c:order`, which must be unique across this chart and its
   * {@link ChartBuilder.secondaryChart} (Excel refuses a file where two series share it).
   * Omitted, it is numbered like `idx`.
   */
  order?: number;
  /**
   * Series name as a cell-reference formula (e.g., `'Sheet1!$B$1'`), written as
   * `c:tx/c:strRef/c:f`. It is not display text: pass literal display names as
   * {@link AddSeriesOptions.nameLiteral}. Written in preference to `nameLiteral`.
   */
  name?: string;
  /** Series name as literal text (e.g., `'Revenue'`), written as `c:tx/c:v` when `name` is unset. */
  nameLiteral?: string;
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
  /** Line width in points (finite, 0-1584); stored as EMU (12700 per point). Written only with `lineColor`. */
  lineWidth?: number;
  /** Marker symbol style (e.g., `'circle'`, `'square'`). */
  marker?: MarkerStyleType;
  /** Enable smooth line rendering for line/scatter charts. */
  smooth?: boolean;
  /** Pie/doughnut slice explosion as a percentage of the radius (non-negative integer; xsd:unsignedInt, no spec maximum). */
  explosion?: number;
  /** Per-series data label configuration. */
  dataLabels?: Partial<DataLabelsData>;
  /**
   * Trendline configuration for this series. `order` (polynomial trendlines) is an integer
   * 2-6 and `period` (moving-average trendlines) an integer of at least 2; both are validated.
   */
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
  /**
   * Axis title text, or an object with text and formatting. `fontSize` is in hundredths of a
   * point (1400 = 14pt), written verbatim as DrawingML `sz`; the validated range is 100-400000.
   */
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
  /** Font size for tick labels in hundredths of a point (1400 = 14pt); the validated range is 100-400000. */
  fontSize?: number;
}

/**
 * Fluent builder for creating chart definitions.
 *
 * @example
 * ```ts
 * const chart = new ChartBuilder('bar')
 *   .title('Monthly Sales')
 *   .addSeries({ nameLiteral: 'Revenue', catRef: 'Sheet1!$A$2:$A$5', valRef: 'Sheet1!$B$2:$B$5' })
 *   .legend('bottom')
 *   .anchor({ col: 0, row: 0 }, { col: 10, row: 20 })
 *   .build();
 * ```
 */
export class ChartBuilder {
  #chartType: ChartType;
  #titleData: ChartTitleData | null = null;
  #seriesList: SeriesEntry[] = [];
  #secondary: ChartBuilder | null = null;
  #catAxisData: ChartAxisData | null = null;
  #valAxisData: ChartAxisData | null = null;
  #legendData: ChartLegendData | null = null;
  #dataLabelsData: DataLabelsData | null = null;
  #groupingValue: ChartGrouping | null = null;
  #scatterStyleValue: ScatterStyle | null = null;
  #radarStyleValue: RadarStyle | null = null;
  #holeSizeValue: number | null = null;
  #barDirValue: boolean | null = null;
  #styleIdValue: number | null = null;
  #layoutValue: ManualLayoutData | null = null;
  #anchorData: ChartAnchorData = { fromCol: 0, fromRow: 0, toCol: 10, toRow: 15 };
  #showDataTableValue = false;
  #view3dData: View3DData | null = null;

  /**
   * Create a new ChartBuilder for the given chart type.
   *
   * @param type - The chart type (e.g., `'bar'`, `'line'`, `'pie'`, `'scatter'`).
   */
  constructor(type: ChartType) {
    this.#chartType = type;
  }

  /**
   * Set the chart title.
   *
   * @param text - The title text.
   * @param opts - Optional formatting: `bold`, `fontSize` in hundredths of a point (1400 = 14pt,
   *   written verbatim as DrawingML `sz`; the validated range is 100-400000) and `color`.
   * @returns `this` for chaining.
   */
  title(text: string, opts?: { bold?: boolean; fontSize?: number; color?: string }): this {
    this.#titleData = {
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
   * Unless `opts` sets them, the series' `idx` and `order` are given by {@link ChartBuilder.build}:
   * the series of this chart and then those of its {@link ChartBuilder.secondaryChart} are
   * numbered 0, 1, 2, … in the order they were added, skipping every value a caller set.
   *
   * @param opts - Series configuration including data references and styling.
   * @returns `this` for chaining.
   *
   * @example
   * ```ts
   * builder.addSeries({
   *   nameLiteral: 'Revenue',
   *   catRef: 'Sheet1!$A$2:$A$5',
   *   valRef: 'Sheet1!$B$2:$B$5',
   * });
   * ```
   */
  addSeries(opts: AddSeriesOptions): this {
    this.#seriesList.push({
      idx: opts.idx,
      order: opts.order,
      data: {
        name: opts.name ?? null,
        nameLiteral: opts.nameLiteral ?? null,
        catRef: opts.catRef ?? null,
        valRef: opts.valRef,
        xValRef: opts.xValRef ?? null,
        bubbleSizeRef: opts.bubbleSizeRef ?? null,
        fillColor: opts.fillColor ?? null,
        lineColor: opts.lineColor ?? null,
        lineWidth: toLineWidthEmu(opts.lineWidth),
        marker: opts.marker ?? null,
        smooth: opts.smooth ?? null,
        explosion: opts.explosion ?? null,
        dataLabels: buildDataLabels(opts.dataLabels),
        trendline: buildTrendline(opts.trendline),
        errorBars: buildErrorBars(opts.errorBars),
      },
    });
    return this;
  }

  /**
   * Make this a combo chart: the chart `secondary` builds becomes
   * {@link ChartDataModel.secondaryChart}, a second chart-type element drawn on this chart's
   * axes (or on `secondaryCatAxis` / `secondaryValAxis` when they are set on the built model)
   * (under a pie or doughnut primary, on the default axes the writer adds).
   * Only its chart type, series and chart-type options (bar direction, grouping, scatter or
   * radar style, data labels, hole size) are written; its own axes, title, legend and anchor
   * are not.
   *
   * `secondary` is read when {@link ChartBuilder.build} runs. Its series continue this chart's
   * numbering: with no `idx` or `order` set, a column chart with two series and a line chart
   * with one get 0 and 1, then 2. Excel refuses a chart whose series share an `order`.
   *
   * @param secondary - The builder of the secondary chart. It must not have a secondary chart
   *   of its own, or `build()` throws a `RangeError`.
   * @returns `this` for chaining.
   *
   * @example
   * ```ts
   * const combo = new ChartBuilder('column')
   *   .addSeries({ nameLiteral: 'Sales', catRef: 'Sheet1!$A$2:$A$5', valRef: 'Sheet1!$B$2:$B$5' })
   *   .secondaryChart(
   *     new ChartBuilder('line').addSeries({ nameLiteral: 'Trend', valRef: 'Sheet1!$C$2:$C$5' }),
   *   )
   *   .build();
   * ```
   */
  secondaryChart(secondary: ChartBuilder): this {
    this.#secondary = secondary;
    return this;
  }

  /**
   * Configure the category axis (axis id 0).
   *
   * @param opts - Axis title, scale, tick, and gridline options.
   * @returns `this` for chaining.
   */
  catAxis(opts?: AxisOptions): this {
    this.#catAxisData = buildAxisData(0, 1, opts);
    return this;
  }

  /**
   * Configure the value axis (axis id 1).
   *
   * @param opts - Axis title, scale, tick, and gridline options.
   * @returns `this` for chaining.
   */
  valAxis(opts?: AxisOptions): this {
    this.#valAxisData = buildAxisData(1, 0, opts);
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
    this.#legendData = { position, overlay };
    return this;
  }

  /**
   * Configure chart-level data labels.
   *
   * @param opts - Data label visibility options (showVal, showCatName, showPercent, etc.).
   * @returns `this` for chaining.
   */
  dataLabels(opts: Partial<DataLabelsData>): this {
    this.#dataLabelsData = normalizeDataLabels(opts);
    return this;
  }

  /**
   * Set the grouping mode (bar/column/line/area).
   *
   * @param g - Grouping type: `'clustered'`, `'stacked'`, `'percentStacked'`, or `'standard'`.
   * @returns `this` for chaining.
   */
  grouping(g: ChartGrouping): this {
    this.#groupingValue = g;
    return this;
  }

  /**
   * Set scatter chart style.
   *
   * @param s - Scatter style: `'lineMarker'`, `'line'`, `'marker'`, `'smooth'`, `'smoothMarker'`.
   * @returns `this` for chaining.
   */
  scatterStyle(s: ScatterStyle): this {
    this.#scatterStyleValue = s;
    return this;
  }

  /**
   * Set radar chart style.
   *
   * @param s - Radar style: `'standard'`, `'marker'`, `'filled'`.
   * @returns `this` for chaining.
   */
  radarStyle(s: RadarStyle): this {
    this.#radarStyleValue = s;
    return this;
  }

  /**
   * Set doughnut hole size (an integer 1-90).
   *
   * @param percent - Hole size as a percentage of the chart radius.
   * @returns `this` for chaining.
   */
  holeSize(percent: number): this {
    this.#holeSizeValue = percent;
    return this;
  }

  /**
   * Set bar direction. `true` = horizontal bars, `false` = vertical columns.
   *
   * @param horizontal - Whether bars are horizontal.
   * @returns `this` for chaining.
   */
  barDirection(horizontal: boolean): this {
    this.#barDirValue = horizontal;
    return this;
  }

  /**
   * Set the chart style ID (an integer 1-48, matches Excel built-in styles).
   *
   * @param id - The style ID.
   * @returns `this` for chaining.
   */
  style(id: number): this {
    this.#styleIdValue = id;
    return this;
  }

  /**
   * Enable showing the data table below the chart.
   *
   * @param show - Whether to display the data table (default: `true`).
   * @returns `this` for chaining.
   */
  showDataTable(show = true): this {
    this.#showDataTableValue = show;
    return this;
  }

  /**
   * Set 3D rotation settings.
   *
   * @param opts - Rotation angles and perspective options.
   * @returns `this` for chaining.
   */
  view3d(opts: { rotX?: number; rotY?: number; perspective?: number; rAngAx?: boolean }): this {
    this.#view3dData = {
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
    this.#layoutValue = { x, y, w, h };
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
    this.#anchorData = {
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
   * Returns an independent deep copy; later builder calls do not affect it.
   *
   * Every chart type except `'pie'` and `'doughnut'` has axes, and the OOXML schema
   * requires two axis ids for them. An axis that was not configured with
   * {@link ChartBuilder.catAxis} or {@link ChartBuilder.valAxis} is therefore filled with its
   * default in the returned model (category id 0, value id 1). Pie and doughnut charts keep
   * `catAxis` and `valAxis` as `null`. For scatter and bubble charts the category slot is
   * written as the bottom value axis. A {@link ChartBuilder.secondaryChart} gets no default
   * axes: it is drawn on this chart's (under a pie or doughnut primary, on the default axes
   * the writer adds).
   *
   * Series without an explicit `idx` or `order` are numbered here, across this chart and then
   * its secondary chart (see {@link ChartBuilder.addSeries}).
   *
   * @returns A WorksheetChartData containing the chart model and anchor position.
   * @throws {RangeError} When the secondary chart has a secondary chart of its own.
   */
  build(): WorksheetChartData {
    const secondary = this.#secondary;
    // TS18030: an optional chain cannot reach a private field, so test for null first.
    if (secondary !== null && secondary.#secondary !== null) {
      throw new RangeError('A secondary chart cannot have a secondary chart of its own');
    }
    const own = this.#seriesList.length;
    const series = numberSeries(
      secondary ? [...this.#seriesList, ...secondary.#seriesList] : this.#seriesList,
    );
    const chart = this.#model(series.slice(0, own), true);
    if (secondary) chart.secondaryChart = secondary.#model(series.slice(own), false);
    return structuredClone({ chart, anchor: this.#anchorData });
  }

  /**
   * This builder's chart model with `series`. A primary chart of a type with axes gets the
   * default axes for unconfigured slots; a secondary chart keeps only what was configured.
   */
  #model(series: ChartSeriesData[], primary: boolean): ChartDataModel {
    const hasAxes = primary && this.#chartType !== 'pie' && this.#chartType !== 'doughnut';
    return {
      chartType: this.#chartType,
      title: this.#titleData,
      series,
      catAxis: hasAxes ? (this.#catAxisData ?? buildAxisData(0, 1, undefined)) : this.#catAxisData,
      valAxis: hasAxes ? (this.#valAxisData ?? buildAxisData(1, 0, undefined)) : this.#valAxisData,
      legend: this.#legendData,
      dataLabels: this.#dataLabelsData,
      grouping: this.#groupingValue,
      scatterStyle: this.#scatterStyleValue,
      radarStyle: this.#radarStyleValue,
      holeSize: this.#holeSizeValue,
      barDirHorizontal: this.#barDirValue,
      styleId: this.#styleIdValue,
      plotAreaLayout: this.#layoutValue,
      showDataTable: this.#showDataTableValue,
      view3d: this.#view3dData,
    };
  }
}

/** A series as added: its model without `idx` and `order`, and the `idx` and `order` the caller set. */
interface SeriesEntry {
  readonly idx: number | undefined;
  readonly order: number | undefined;
  readonly data: Omit<ChartSeriesData, 'idx' | 'order'>;
}

/**
 * The series of `entries` with their `idx` and `order`: the caller's value when set, else the
 * next value from 0 up that no entry sets, in entry order.
 */
function numberSeries(entries: readonly SeriesEntry[]): ChartSeriesData[] {
  const counter = (key: 'idx' | 'order'): ((set: number | undefined) => number) => {
    const taken = new Set(entries.flatMap((e) => e[key] ?? []));
    let next = 0;
    return (set) => {
      if (set != null) return set;
      while (taken.has(next)) next++;
      return next++;
    };
  };
  const idx = counter('idx');
  const order = counter('order');
  return entries.map((e) => ({ idx: idx(e.idx), order: order(e.order), ...e.data }));
}

const EMU_PER_POINT = 12700;
const MAX_LINE_WIDTH_POINTS = 1584;

function toLineWidthEmu(points: number | undefined): number | null {
  if (points == null) return null;
  if (!Number.isFinite(points) || points < 0 || points > MAX_LINE_WIDTH_POINTS) {
    throw new RangeError(
      `lineWidth must be a finite number of points between 0 and ${MAX_LINE_WIDTH_POINTS}, got ${points}`,
    );
  }
  return Math.round(points * EMU_PER_POINT);
}

function normalizeDataLabels(opts: Partial<DataLabelsData>): DataLabelsData {
  return {
    showVal: opts.showVal ?? false,
    showCatName: opts.showCatName ?? false,
    showSerName: opts.showSerName ?? false,
    showPercent: opts.showPercent ?? false,
    numFmt: opts.numFmt ?? null,
    showLeaderLines: opts.showLeaderLines ?? false,
    showLegendKey: opts.showLegendKey ?? null,
    showBubbleSize: opts.showBubbleSize ?? null,
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
