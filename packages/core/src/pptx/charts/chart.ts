/**
 * Chart part (`chartN.xml`, root `<c:chartSpace>`) -> {@link Chart} model.
 *
 * We read only the cached values baked into the chart (`<c:numCache>` /
 * `<c:strCache>`), not the live spreadsheet link — that snapshot is exactly what
 * PowerPoint draws when the workbook isn't available. Our XML helpers match by
 * local name, so the `c:` prefix is transparent here.
 *
 * Series colors come from each series' `<c:spPr>` (or per-point `<c:dPt>`); when
 * absent we cycle the theme accent palette, matching PowerPoint's auto colors.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import type { Chart, ChartKind, ChartSeries, ChartGrouping, Color } from '../model.js';
import type { ParseScope } from '../scope.js';
import { parseFill } from '../shapes/fill.js';

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

export function parseChart(chartSpace: XmlNode, scope: ParseScope): Chart | undefined {
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

  const palette = accentPalette(scope);
  const series: ChartSeries[] = [];
  let categories: string[] = [];

  const serEls = children(typeEl, 'ser');
  serEls.forEach((ser, i) => {
    const isScatter = kind === 'scatter';
    const values = isScatter ? numValues(child(ser, 'yVal')) : numValues(child(ser, 'val'));
    if (!values.length) return;

    const s: ChartSeries = {
      name: seriesName(ser),
      values,
      color: seriesColor(ser, scope) ?? palette[i % palette.length],
    };
    if (isScatter) {
      const xv = numValues(child(ser, 'xVal'));
      if (xv.length) s.xValues = xv;
    }
    const pts = pointColors(ser, scope);
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

  const title = chartTitle(child(chartSpace, 'chart'));
  if (title) chart.title = title;

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

function seriesColor(ser: XmlNode, scope: ParseScope): Color | undefined {
  const fill = parseFill(child(ser, 'spPr'), scope);
  return fill?.type === 'solid' ? fill.color : undefined;
}

/** Per-point override colors from `<c:dPt>` elements, indexed by `<c:idx>`. */
function pointColors(ser: XmlNode, scope: ParseScope): (Color | undefined)[] {
  const out: (Color | undefined)[] = [];
  for (const dPt of children(ser, 'dPt')) {
    const idx = attrNum(child(dPt, 'idx'), 'val');
    const fill = parseFill(child(dPt, 'spPr'), scope);
    if (idx !== undefined && fill?.type === 'solid') out[idx] = fill.color;
  }
  return out;
}

function chartTitle(chartEl: XmlNode | undefined): string | undefined {
  const title = child(chartEl, 'title');
  if (!title || attr(child(chartEl, 'autoTitleDeleted'), 'val') === '1') {
    const rich = child(child(title, 'tx'), 'rich');
    if (!rich) return undefined;
  }
  const rich = child(child(title, 'tx'), 'rich');
  if (!rich) return undefined;
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
function accentPalette(scope: ParseScope): Color[] {
  const colors = scope.colorCtx.theme.colors;
  return ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
    .map((k) => colors[k])
    .filter((hex): hex is string => !!hex)
    .map((hex) => ({ hex }));
}
