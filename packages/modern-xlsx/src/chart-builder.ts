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
  LegendPosition,
  ManualLayoutData,
  MarkerStyleType,
  RadarStyle,
  ScatterStyle,
  TickLabelPosition,
  TickMark,
  WorksheetChartData,
} from './types.js';

export interface AddSeriesOptions {
  name?: string;
  catRef?: string;
  valRef: string;
  xValRef?: string;
  bubbleSizeRef?: string;
  fillColor?: string;
  lineColor?: string;
  lineWidth?: number;
  marker?: MarkerStyleType;
  smooth?: boolean;
  explosion?: number;
  dataLabels?: Partial<DataLabelsData>;
}

export interface AxisOptions {
  title?: string | { text: string; bold?: boolean; fontSize?: number; color?: string };
  numFmt?: string;
  sourceLinked?: boolean;
  min?: number;
  max?: number;
  majorUnit?: number;
  minorUnit?: number;
  logBase?: number;
  reversed?: boolean;
  tickLblPos?: TickLabelPosition;
  majorTickMark?: TickMark;
  minorTickMark?: TickMark;
  majorGridlines?: boolean;
  minorGridlines?: boolean;
  delete?: boolean;
  position?: AxisPosition;
  crossesAt?: number;
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

  constructor(type: ChartType) {
    this.chartType = type;
  }

  /** Set the chart title. */
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

  /** Add a data series. */
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
    });
    return this;
  }

  /** Configure the category axis (axis id 0). */
  catAxis(opts?: AxisOptions): this {
    this.catAxisData = buildAxisData(0, 1, opts);
    return this;
  }

  /** Configure the value axis (axis id 1). */
  valAxis(opts?: AxisOptions): this {
    this.valAxisData = buildAxisData(1, 0, opts);
    return this;
  }

  /** Configure the legend. */
  legend(position: LegendPosition = 'bottom', overlay = false): this {
    this.legendData = { position, overlay };
    return this;
  }

  /** Configure chart-level data labels. */
  dataLabels(opts: Partial<DataLabelsData>): this {
    this.dataLabelsData = normalizeDataLabels(opts);
    return this;
  }

  /** Set the grouping mode (bar/column/line/area). */
  grouping(g: ChartGrouping): this {
    this.groupingValue = g;
    return this;
  }

  /** Set scatter chart style. */
  scatterStyle(s: ScatterStyle): this {
    this.scatterStyleValue = s;
    return this;
  }

  /** Set radar chart style. */
  radarStyle(s: RadarStyle): this {
    this.radarStyleValue = s;
    return this;
  }

  /** Set doughnut hole size (0-90). */
  holeSize(percent: number): this {
    this.holeSizeValue = percent;
    return this;
  }

  /** Set bar direction. true = horizontal bars, false = vertical columns. */
  barDirection(horizontal: boolean): this {
    this.barDirValue = horizontal;
    return this;
  }

  /** Set the chart style ID (1-48, matches Excel built-in styles). */
  style(id: number): this {
    this.styleIdValue = id;
    return this;
  }

  /** Set the plot area manual layout (fractional 0.0-1.0). */
  plotLayout(x: number, y: number, w: number, h: number): this {
    this.layoutValue = { x, y, w, h };
    return this;
  }

  /** Set the anchor position (cell coordinates). */
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

  /** Build the final chart data. */
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

function buildAxisTitle(title: AxisOptions['title']): ChartTitleData | null {
  if (!title) return null;
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
  fontSize: null,
};

function buildAxisScaleProps(
  opts: AxisOptions,
): Pick<
  ChartAxisData,
  | 'numFmt'
  | 'sourceLinked'
  | 'min'
  | 'max'
  | 'majorUnit'
  | 'minorUnit'
  | 'logBase'
  | 'reversed'
  | 'fontSize'
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
    fontSize: opts.fontSize ?? null,
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
