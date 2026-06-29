/**
 * Chart model -> SVG.
 *
 * Lays out a chart inside the frame's pixel box: an optional title band, an
 * optional legend strip, and the plot area filling the rest. All chart kinds
 * draw into the same plot rect, so axis/gridline code is shared between
 * bar/line/area/scatter, while pie/doughnut use a centered radial layout.
 *
 * This is a faithful-enough static renderer (the cached snapshot PowerPoint
 * shows when the workbook is detached) — not an interactive charting library.
 */
import type { Chart, ChartSeries, Color } from '../model.js';
import { colorToCss } from '../color.js';
import { SVG_NS } from '../render/primitives.js';

const AXIS_COLOR = '#bfbfbf';
const GRID_COLOR = '#e6e6e6';
const TEXT_COLOR = '#595959';
const FONT = '12px "Segoe UI", system-ui, sans-serif';

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export function renderChart(chart: Chart, w: number, h: number): SVGSVGElement {
  const svg = el('svg') as SVGSVGElement;
  svg.setAttribute('width', `${w}`);
  svg.setAttribute('height', `${h}`);
  svg.setAttribute('viewBox', `0 0 ${w} ${h}`);
  svg.style.position = 'absolute';
  svg.style.inset = '0';
  svg.style.font = FONT;

  let plot: Rect = { x: 0, y: 0, w, h };

  if (chart.title) {
    plot = { ...plot, y: 24, h: plot.h - 24 };
    svg.appendChild(text(w / 2, 16, chart.title, { anchor: 'middle', weight: 600, size: 14, fill: '#404040' }));
  }

  if (chart.legend) {
    plot = layoutLegend(svg, chart, plot);
  }

  switch (chart.kind) {
    case 'pie':
    case 'doughnut':
      renderPie(svg, chart, plot);
      break;
    case 'scatter':
      renderScatter(svg, chart, plot);
      break;
    case 'line':
      renderLineArea(svg, chart, plot, false);
      break;
    case 'area':
      renderLineArea(svg, chart, plot, true);
      break;
    case 'bar':
    default:
      renderBar(svg, chart, plot);
      break;
  }

  return svg;
}

// ---------------------------------------------------------------------------
// Legend
// ---------------------------------------------------------------------------

function legendEntries(chart: Chart): Array<{ label: string; color: Color }> {
  // Pie/doughnut legends list categories (one color per slice); other charts
  // list series.
  if ((chart.kind === 'pie' || chart.kind === 'doughnut') && chart.series[0]) {
    const s = chart.series[0];
    return chart.categories.map((label, i) => ({
      label,
      color: s.pointColors?.[i] ?? sliceColor(s, i),
    }));
  }
  return chart.series.map((s, i) => ({
    label: s.name ?? `Series ${i + 1}`,
    color: s.color ?? { hex: '4472C4' },
  }));
}

function layoutLegend(svg: SVGSVGElement, chart: Chart, plot: Rect): Rect {
  const entries = legendEntries(chart);
  const pos = chart.legend!;
  const swatch = 10;
  const rowH = 18;
  const pad = 8;

  if (pos === 'b' || pos === 't') {
    // Horizontal legend strip.
    const widths = entries.map((e) => swatch + 4 + measure(e.label) + 14);
    const total = widths.reduce((a, b) => a + b, 0);
    let x = plot.x + Math.max(0, (plot.w - total) / 2);
    const y = pos === 'b' ? plot.y + plot.h - rowH : plot.y;
    for (const e of entries) {
      legendItem(svg, x, y, swatch, e);
      x += swatch + 4 + measure(e.label) + 14;
    }
    return pos === 'b'
      ? { ...plot, h: plot.h - rowH - pad }
      : { ...plot, y: plot.y + rowH + pad, h: plot.h - rowH - pad };
  }

  // Vertical legend (right/left/top-right).
  const colW = Math.min(
    plot.w * 0.32,
    Math.max(...entries.map((e) => swatch + 6 + measure(e.label))) + 12,
  );
  const left = pos === 'l';
  const x = left ? plot.x : plot.x + plot.w - colW;
  let y = plot.y + Math.max(pad, (plot.h - entries.length * rowH) / 2);
  for (const e of entries) {
    legendItem(svg, x, y, swatch, e);
    y += rowH;
  }
  return left ? { ...plot, x: plot.x + colW, w: plot.w - colW } : { ...plot, w: plot.w - colW };
}

function legendItem(
  svg: SVGSVGElement,
  x: number,
  y: number,
  swatch: number,
  e: { label: string; color: Color },
): void {
  const r = el('rect');
  r.setAttribute('x', `${x}`);
  r.setAttribute('y', `${y + 3}`);
  r.setAttribute('width', `${swatch}`);
  r.setAttribute('height', `${swatch}`);
  r.setAttribute('fill', colorToCss(e.color));
  r.setAttribute('rx', '1');
  svg.appendChild(r);
  svg.appendChild(text(x + swatch + 4, y + 11, e.label, { fill: TEXT_COLOR }));
}

// ---------------------------------------------------------------------------
// Cartesian charts (bar / column / line / area / scatter)
// ---------------------------------------------------------------------------

/** Value range across series, honoring stacking; always includes 0. */
function valueRange(chart: Chart): { min: number; max: number } {
  const stacked = chart.grouping === 'stacked' || chart.grouping === 'percentStacked';
  let min = 0;
  let max = 0;
  if (chart.grouping === 'percentStacked') return { min: 0, max: 1 };
  const n = Math.max(...chart.series.map((s) => s.values.length), 0);
  if (stacked) {
    for (let i = 0; i < n; i++) {
      let pos = 0;
      let neg = 0;
      for (const s of chart.series) {
        const v = s.values[i] ?? 0;
        if (v >= 0) pos += v;
        else neg += v;
      }
      max = Math.max(max, pos);
      min = Math.min(min, neg);
    }
  } else {
    for (const s of chart.series) {
      for (const v of s.values) {
        max = Math.max(max, v);
        min = Math.min(min, v);
      }
    }
  }
  if (max === min) max = min + 1;
  return niceRange(min, max);
}

/** Round a range out to "nice" tick boundaries. */
function niceRange(min: number, max: number): { min: number; max: number } {
  const span = max - min;
  const step = niceStep(span / 5);
  return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step };
}

function niceStep(raw: number): number {
  const mag = Math.pow(10, Math.floor(Math.log10(raw || 1)));
  const norm = raw / mag;
  const nice = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10;
  return nice * mag;
}

function plotInsets(chart: Chart, plot: Rect): Rect {
  // Leave room for value-axis labels (left) and category labels (bottom).
  const left = chart.kind === 'scatter' ? 40 : 44;
  const bottom = 22;
  const top = 8;
  const right = 12;
  return { x: plot.x + left, y: plot.y + top, w: plot.w - left - right, h: plot.h - top - bottom };
}

function renderBar(svg: SVGSVGElement, chart: Chart, plot: Rect): void {
  const area = plotInsets(chart, plot);
  const { min, max } = valueRange(chart);
  const horizontal = chart.barHorizontal === true;
  const stacked = chart.grouping === 'stacked' || chart.grouping === 'percentStacked';
  const percent = chart.grouping === 'percentStacked';
  const cats = chart.categories.length;

  drawAxes(svg, area, chart, min, max, horizontal);

  const valToPx = (v: number) =>
    horizontal
      ? area.x + ((v - min) / (max - min)) * area.w
      : area.y + area.h - ((v - min) / (max - min)) * area.h;
  const zero = valToPx(0);

  const bandSize = (horizontal ? area.h : area.w) / Math.max(cats, 1);
  const groupPad = bandSize * 0.2;
  const slot = bandSize - groupPad;
  const nSer = chart.series.length;
  const barW = stacked ? slot : slot / nSer;

  for (let ci = 0; ci < cats; ci++) {
    const bandStart = (horizontal ? area.y : area.x) + ci * bandSize + groupPad / 2;
    let posAcc = 0;
    let negAcc = 0;
    let colTotal = 0;
    if (percent) {
      colTotal = chart.series.reduce((a, s) => a + Math.abs(s.values[ci] ?? 0), 0) || 1;
    }
    chart.series.forEach((s, si) => {
      let v = s.values[ci] ?? 0;
      if (percent) v = v / colTotal;
      const color = s.pointColors?.[ci] ?? s.color ?? { hex: '4472C4' };
      let rect: Rect;
      if (stacked) {
        const base = v >= 0 ? posAcc : negAcc;
        const top = base + v;
        if (v >= 0) posAcc = top;
        else negAcc = top;
        rect = horizontal
          ? boxH(area, bandStart, barW, valToPx(base), valToPx(top))
          : boxV(area, bandStart, barW, valToPx(base), valToPx(top));
      } else {
        const off = si * barW;
        rect = horizontal
          ? boxH(area, bandStart, barW, zero, valToPx(v), off)
          : boxV(area, bandStart, barW, zero, valToPx(v), off);
      }
      drawRect(svg, rect, color);
    });
  }
}

function boxV(_area: Rect, x: number, w: number, y0: number, y1: number, off = 0): Rect {
  const top = Math.min(y0, y1);
  return { x: x + off, y: top, w: w - 1, h: Math.abs(y1 - y0) };
}
function boxH(_area: Rect, y: number, h: number, x0: number, x1: number, off = 0): Rect {
  const left = Math.min(x0, x1);
  return { x: left, y: y + off, w: Math.abs(x1 - x0), h: h - 1 };
}

function renderLineArea(svg: SVGSVGElement, chart: Chart, plot: Rect, area: boolean): void {
  const a = plotInsets(chart, plot);
  const { min, max } = valueRange(chart);
  const cats = chart.categories.length;
  drawAxes(svg, a, chart, min, max, false);

  const xAt = (i: number) => a.x + (cats <= 1 ? a.w / 2 : (i / (cats - 1)) * a.w);
  const yAt = (v: number) => a.y + a.h - ((v - min) / (max - min)) * a.h;

  chart.series.forEach((s) => {
    const color = s.color ?? { hex: '4472C4' };
    const pts = s.values.map((v, i) => `${xAt(i)},${yAt(v)}`);
    if (!pts.length) return;
    if (area) {
      const poly = el('polygon');
      poly.setAttribute('points', `${xAt(0)},${yAt(min)} ${pts.join(' ')} ${xAt(s.values.length - 1)},${yAt(min)}`);
      poly.setAttribute('fill', colorToCss({ ...color, alpha: 0.6 }));
      svg.appendChild(poly);
    }
    const line = el('polyline');
    line.setAttribute('points', pts.join(' '));
    line.setAttribute('fill', 'none');
    line.setAttribute('stroke', colorToCss(color));
    line.setAttribute('stroke-width', area ? '1.5' : '2.25');
    line.setAttribute('stroke-linejoin', 'round');
    svg.appendChild(line);
    if (!area) {
      s.values.forEach((v, i) => svg.appendChild(dot(xAt(i), yAt(v), color)));
    }
  });
}

function renderScatter(svg: SVGSVGElement, chart: Chart, plot: Rect): void {
  const a = plotInsets(chart, plot);
  let xMin = Infinity;
  let xMax = -Infinity;
  let yMin = Infinity;
  let yMax = -Infinity;
  for (const s of chart.series) {
    s.values.forEach((y, i) => {
      const x = s.xValues?.[i] ?? i;
      xMin = Math.min(xMin, x);
      xMax = Math.max(xMax, x);
      yMin = Math.min(yMin, y);
      yMax = Math.max(yMax, y);
    });
  }
  if (!Number.isFinite(xMin)) return;
  const yr = niceRange(Math.min(0, yMin), yMax);
  const xr = niceRange(xMin, xMax);
  const xAt = (x: number) => a.x + ((x - xr.min) / (xr.max - xr.min || 1)) * a.w;
  const yAt = (y: number) => a.y + a.h - ((y - yr.min) / (yr.max - yr.min || 1)) * a.h;

  // Axes with numeric ticks on both dimensions.
  axisLine(svg, a.x, a.y, a.x, a.y + a.h);
  axisLine(svg, a.x, a.y + a.h, a.x + a.w, a.y + a.h);
  for (let i = 0; i <= 5; i++) {
    const yv = yr.min + ((yr.max - yr.min) / 5) * i;
    const yp = yAt(yv);
    gridLine(svg, a.x, yp, a.x + a.w, yp);
    svg.appendChild(text(a.x - 6, yp + 4, fmt(yv), { anchor: 'end', fill: TEXT_COLOR }));
  }
  chart.series.forEach((s) => {
    const color = s.color ?? { hex: '4472C4' };
    s.values.forEach((y, i) => svg.appendChild(dot(xAt(s.xValues?.[i] ?? i), yAt(y), color, 3.5)));
  });
}

function drawAxes(
  svg: SVGSVGElement,
  a: Rect,
  chart: Chart,
  min: number,
  max: number,
  horizontal: boolean,
): void {
  axisLine(svg, a.x, a.y, a.x, a.y + a.h); // y axis
  axisLine(svg, a.x, a.y + a.h, a.x + a.w, a.y + a.h); // x axis

  const ticks = 5;
  for (let i = 0; i <= ticks; i++) {
    const v = min + ((max - min) / ticks) * i;
    if (horizontal) {
      const xp = a.x + ((v - min) / (max - min)) * a.w;
      gridLine(svg, xp, a.y, xp, a.y + a.h);
      svg.appendChild(text(xp, a.y + a.h + 14, fmtAxis(v, chart), { anchor: 'middle', fill: TEXT_COLOR }));
    } else {
      const yp = a.y + a.h - ((v - min) / (max - min)) * a.h;
      gridLine(svg, a.x, yp, a.x + a.w, yp);
      svg.appendChild(text(a.x - 6, yp + 4, fmtAxis(v, chart), { anchor: 'end', fill: TEXT_COLOR }));
    }
  }

  // Category labels along the band axis.
  const cats = chart.categories;
  const band = (horizontal ? a.h : a.w) / Math.max(cats.length, 1);
  cats.forEach((c, i) => {
    const center = (horizontal ? a.y : a.x) + band * (i + 0.5);
    if (horizontal) {
      svg.appendChild(text(a.x - 6, center + 4, c, { anchor: 'end', fill: TEXT_COLOR }));
    } else {
      svg.appendChild(text(center, a.y + a.h + 14, ellipsize(c, band), { anchor: 'middle', fill: TEXT_COLOR }));
    }
  });
}

function fmtAxis(v: number, chart: Chart): string {
  return chart.grouping === 'percentStacked' ? `${Math.round(v * 100)}%` : fmt(v);
}

// ---------------------------------------------------------------------------
// Pie / doughnut
// ---------------------------------------------------------------------------

function sliceColor(s: ChartSeries, i: number): Color {
  // Even without per-point colors, pie slices must differ — derive a palette
  // by rotating the series hue when only one base color exists.
  const palette = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E'];
  return s.pointColors?.[i] ?? { hex: palette[i % palette.length]! };
}

function renderPie(svg: SVGSVGElement, chart: Chart, plot: Rect): void {
  const s = chart.series[0];
  if (!s) return;
  const total = s.values.reduce((a, v) => a + Math.abs(v), 0);
  if (total <= 0) return;

  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2;
  const r = Math.min(plot.w, plot.h) / 2 - 8;
  const inner = chart.kind === 'doughnut' ? r * (chart.holeSize ?? 0.5) : 0;

  let angle = -Math.PI / 2; // start at 12 o'clock
  s.values.forEach((v, i) => {
    const frac = Math.abs(v) / total;
    const end = angle + frac * Math.PI * 2;
    const color = s.pointColors?.[i] ?? sliceColor(s, i);
    svg.appendChild(arcSlice(cx, cy, r, inner, angle, end, color));
    angle = end;
  });
}

function arcSlice(
  cx: number,
  cy: number,
  r: number,
  inner: number,
  a0: number,
  a1: number,
  color: Color,
): SVGElement {
  const large = a1 - a0 > Math.PI ? 1 : 0;
  const x0 = cx + r * Math.cos(a0);
  const y0 = cy + r * Math.sin(a0);
  const x1 = cx + r * Math.cos(a1);
  const y1 = cy + r * Math.sin(a1);
  let d: string;
  if (inner > 0) {
    const ix0 = cx + inner * Math.cos(a0);
    const iy0 = cy + inner * Math.sin(a0);
    const ix1 = cx + inner * Math.cos(a1);
    const iy1 = cy + inner * Math.sin(a1);
    d = `M${x0},${y0} A${r},${r} 0 ${large} 1 ${x1},${y1} L${ix1},${iy1} A${inner},${inner} 0 ${large} 0 ${ix0},${iy0} Z`;
  } else {
    d = `M${cx},${cy} L${x0},${y0} A${r},${r} 0 ${large} 1 ${x1},${y1} Z`;
  }
  const p = el('path');
  p.setAttribute('d', d);
  p.setAttribute('fill', colorToCss(color));
  p.setAttribute('stroke', '#fff');
  p.setAttribute('stroke-width', '1');
  return p;
}

// ---------------------------------------------------------------------------
// Low-level SVG helpers
// ---------------------------------------------------------------------------

function el(name: string): SVGElement {
  return document.createElementNS(SVG_NS, name);
}

function drawRect(svg: SVGSVGElement, r: Rect, color: Color): void {
  if (r.w <= 0 || r.h <= 0) return;
  const rect = el('rect');
  rect.setAttribute('x', `${r.x}`);
  rect.setAttribute('y', `${r.y}`);
  rect.setAttribute('width', `${r.w}`);
  rect.setAttribute('height', `${r.h}`);
  rect.setAttribute('fill', colorToCss(color));
  svg.appendChild(rect);
}

function dot(cx: number, cy: number, color: Color, r = 3): SVGCircleElement {
  const c = el('circle') as SVGCircleElement;
  c.setAttribute('cx', `${cx}`);
  c.setAttribute('cy', `${cy}`);
  c.setAttribute('r', `${r}`);
  c.setAttribute('fill', colorToCss(color));
  return c;
}

function axisLine(svg: SVGSVGElement, x1: number, y1: number, x2: number, y2: number): void {
  svg.appendChild(seg(x1, y1, x2, y2, AXIS_COLOR));
}
function gridLine(svg: SVGSVGElement, x1: number, y1: number, x2: number, y2: number): void {
  svg.appendChild(seg(x1, y1, x2, y2, GRID_COLOR));
}
function seg(x1: number, y1: number, x2: number, y2: number, color: string): SVGLineElement {
  const l = el('line') as SVGLineElement;
  l.setAttribute('x1', `${x1}`);
  l.setAttribute('y1', `${y1}`);
  l.setAttribute('x2', `${x2}`);
  l.setAttribute('y2', `${y2}`);
  l.setAttribute('stroke', color);
  return l;
}

interface TextOpts {
  anchor?: 'start' | 'middle' | 'end';
  weight?: number;
  size?: number;
  fill?: string;
}
function text(x: number, y: number, content: string, opts: TextOpts = {}): SVGTextElement {
  const t = el('text') as SVGTextElement;
  t.setAttribute('x', `${x}`);
  t.setAttribute('y', `${y}`);
  if (opts.anchor) t.setAttribute('text-anchor', opts.anchor);
  if (opts.weight) t.setAttribute('font-weight', `${opts.weight}`);
  t.setAttribute('font-size', `${opts.size ?? 11}`);
  t.setAttribute('fill', opts.fill ?? TEXT_COLOR);
  t.textContent = content;
  return t;
}

/** Approximate text width (no DOM measurement during build). */
function measure(s: string): number {
  return s.length * 6.2;
}
function ellipsize(s: string, maxPx: number): string {
  const max = Math.max(1, Math.floor(maxPx / 6.2));
  return s.length > max ? s.slice(0, Math.max(1, max - 1)) + '…' : s;
}

function fmt(v: number): string {
  if (v === 0) return '0';
  const abs = Math.abs(v);
  if (abs >= 1000) return `${(v / 1000).toFixed(abs >= 10000 ? 0 : 1)}k`;
  if (Number.isInteger(v)) return `${v}`;
  return v.toFixed(abs < 1 ? 2 : 1);
}
