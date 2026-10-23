/**
 * Chart part (`chartN.xml`, root `<c:chartSpace>`) -> {@link Chart} model.
 *
 * We read only the cached values baked into the chart (`<c:numCache>` /
 * `<c:strCache>`), not the live spreadsheet link — that snapshot is exactly what
 * PowerPoint and Word draw when the workbook isn't available. Our XML helpers match by
 * local name, so the `c:` prefix is transparent here.
 *
 * Series colors come from each series' `<c:spPr>` (or per-point `<c:dPt>`); when
 * absent we cycle the theme accent palette, matching PowerPoint's auto colors.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../xml.js';
import type { Chart, ChartKind, ChartSeries, ChartGrouping } from './model.js';
import { resolveContainerColor, type Color, type ColorContext } from '../color.js';

/** plotArea child element -> chart kind. 3D variants collapse to their 2D kind. */
const TYPE_ELEMENTS: Array<[string, ChartKind]> = [
  ['barChart', 'bar'],
  ['bar3DChart', 'bar'],
  ['lineChart', 'line'],
  ['line3DChart', 'line'],
  ['pieChart', 'pie'],
  ['pie3DChart', 'pie'],
  ['ofPieChart', 'pie'],
  ['doughnutChart', 'doughnut'],
  ['areaChart', 'area'],
  ['area3DChart', 'area'],
  ['scatterChart', 'scatter'],
  ['bubbleChart', 'scatter'],
  ['radarChart', 'line'],
];

export function parseChart(chartSpace: XmlNode, colors: ColorContext): Chart | undefined {
  const plotArea = child(child(chartSpace, 'chart'), 'plotArea');
  if (!plotArea) return undefined;

  let typeEl: XmlNode | undefined;
  let kind: ChartKind | undefined;
  for (const c of plotArea.children) {
    const ln = localName(c.name);
    const hit = TYPE_ELEMENTS.find(([el]) => el === ln);
    if (hit) {
      typeEl = c;
      kind = hit[1];
      break;
    }
  }
  if (!typeEl || !kind) return undefined;

  const palette = accentPalette(colors);
  const series: ChartSeries[] = [];
  let categories: string[] = [];

  const serEls = children(typeEl, 'ser');
  // Every scatter style but "marker" joins the points; a series can still
  // switch its own line off (<c:spPr><a:ln><a:noFill/>).
  const scatterStyle = attr(child(typeEl, 'scatterStyle'), 'val') ?? 'marker';
  const scatterLines = kind === 'scatter' && scatterStyle !== 'marker' && scatterStyle !== 'none';
  serEls.forEach((ser, i) => {
    const isScatter = kind === 'scatter';
    const values = isScatter ? numValues(child(ser, 'yVal')) : numValues(child(ser, 'val'));
    if (!values.length) return;

    const s: ChartSeries = {
      name: seriesName(ser),
      values,
      color: seriesColor(ser, colors) ?? palette[i % palette.length],
    };
    if (isScatter) {
      const xv = numValues(child(ser, 'xVal'));
      if (xv.length) s.xValues = xv;
      if (scatterLines && !child(child(child(ser, 'spPr'), 'ln'), 'noFill')) s.line = true;
    }
    if (attr(child(child(ser, 'marker'), 'symbol'), 'val') === 'none') s.noMarkers = true;
    const pts = pointColors(ser, colors);
    if (pts.length) s.pointColors = pts;
    series.push(s);

    // Categories are repeated identically on each series; take the longest.
    const cats = strValues(child(ser, 'cat'));
    if (cats.length > categories.length) categories = cats;
  });

  if (!series.length) return undefined;

  // Synthesize category labels for pie/single-series when none are cached.
  if (!categories.length) {
    const n = Math.max(...series.map((s) => s.values.length));
    categories = Array.from({ length: n }, (_, i) => `${i + 1}`);
  }

  const chart: Chart = { kind, categories, series };

  if (kind === 'bar') {
    chart.barHorizontal = attr(child(typeEl, 'barDir'), 'val') === 'bar';
  }
  const grouping = attr(child(typeEl, 'grouping'), 'val') as ChartGrouping | undefined;
  if (grouping) chart.grouping = grouping;

  if (kind === 'doughnut') {
    const hole = attrNum(child(typeEl, 'holeSize'), 'val');
    chart.holeSize = hole !== undefined ? hole / 100 : 0.5;
  }

  const title = chartTitle(child(chartSpace, 'chart'), series);
  if (title) chart.title = title;

  if (kind === 'scatter') {
    // The X axis is the value axis stated first in the type's <c:axId> pair.
    const xId = attr(children(typeEl, 'axId')[0], 'val');
    const xAx = children(plotArea, 'valAx').find((ax) => attr(child(ax, 'axId'), 'val') === xId);
    if (child(xAx, 'majorGridlines')) chart.xGridlines = true;
  }

  const legendPos = attr(child(child(child(chartSpace, 'chart'), 'legend'), 'legendPos'), 'val');
  if (child(child(chartSpace, 'chart'), 'legend')) {
    chart.legend = (legendPos as Chart['legend']) ?? 'r';
  }

  if (hasDataLabels(typeEl)) chart.showValueLabels = true;

  return chart;
}

/** Read `<c:pt idx><c:v>` numeric values from a `<c:val>`/`<c:xVal>` container, ordered by idx. */
function numValues(container: XmlNode | undefined): number[] {
  const cache = child(container, 'numRef') ?? container;
  const numCache = child(cache, 'numCache') ?? child(container, 'numCache');
  const pts = points(numCache ?? child(child(container, 'numRef'), 'numCache'));
  return pts.map((v) => {
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  });
}

/** Read string values from a `<c:cat>`/`<c:tx>` container (str or num cache). */
function strValues(container: XmlNode | undefined): string[] {
  const ref = child(container, 'strRef') ?? child(container, 'numRef') ?? container;
  const cache = child(ref, 'strCache') ?? child(ref, 'numCache');
  return points(cache);
}

/** Ordered `<c:v>` strings from a cache element, honoring `<c:pt idx>` gaps. */
function points(cache: XmlNode | undefined): string[] {
  if (!cache) return [];
  const out: string[] = [];
  for (const pt of children(cache, 'pt')) {
    const idx = attrNum(pt, 'idx') ?? out.length;
    out[idx] = child(pt, 'v')?.text ?? '';
  }
  for (let i = 0; i < out.length; i++) if (out[i] === undefined) out[i] = '';
  return out;
}

function seriesName(ser: XmlNode): string | undefined {
  const tx = child(ser, 'tx');
  const cached = strValues(tx)[0];
  if (cached) return cached;
  // Rich-text title: concatenate run text.
  const rich = child(tx, 'rich') ?? child(child(tx, 'strRef'), 'rich');
  if (rich) {
    const text = children(rich, 'p')
      .flatMap((p) => children(p, 'r').map((r) => child(r, 't')?.text ?? ''))
      .join('');
    if (text) return text;
  }
  return undefined;
}

/** A solid fill on a series or point's <c:spPr>; other fills fall back to the palette. */
function solidColor(spPr: XmlNode | undefined, colors: ColorContext): Color | undefined {
  return resolveContainerColor(child(spPr, 'solidFill'), colors);
}

function seriesColor(ser: XmlNode, colors: ColorContext): Color | undefined {
  return solidColor(child(ser, 'spPr'), colors);
}

/** Per-point override colors from `<c:dPt>` elements, indexed by `<c:idx>`. */
function pointColors(ser: XmlNode, colors: ColorContext): (Color | undefined)[] {
  const out: (Color | undefined)[] = [];
  for (const dPt of children(ser, 'dPt')) {
    const idx = attrNum(child(dPt, 'idx'), 'val');
    const color = solidColor(child(dPt, 'spPr'), colors);
    if (idx !== undefined && color) out[idx] = color;
  }
  return out;
}

/**
 * The chart title's text. A <c:title> with no text of its own is an automatic
 * title, which Office fills with the series name when there is one series.
 */
function chartTitle(chartEl: XmlNode | undefined, series: ChartSeries[]): string | undefined {
  const title = child(chartEl, 'title');
  if (!title) return undefined;
  const tx = child(title, 'tx');
  const rich = child(tx, 'rich');
  if (!rich) {
    const ref = strValues(tx)[0];
    if (ref) return ref;
    return series.length === 1 ? series[0]!.name : undefined;
  }
  const text = children(rich, 'p')
    .flatMap((p) => children(p, 'r').map((r) => child(r, 't')?.text ?? ''))
    .join('');
  return text || undefined;
}

function hasDataLabels(typeEl: XmlNode): boolean {
  const dLbls = child(typeEl, 'dLbls');
  if (!dLbls) return false;
  return attr(child(dLbls, 'showVal'), 'val') === '1';
}

/** Theme accent1..6 as the default series-color cycle. */
function accentPalette(ctx: ColorContext): Color[] {
  const colors = ctx.theme.colors;
  return ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
    .map((k) => colors[k])
    .filter((hex): hex is string => !!hex)
    .map((hex) => ({ hex }));
}
