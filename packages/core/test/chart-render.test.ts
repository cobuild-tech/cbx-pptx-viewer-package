/**
 * Renderer smoke test: drives renderChart against a tiny serializable SVG-DOM
 * shim (the package test env is `node`, so there's no real document). Asserts
 * the right primitives are emitted per chart kind, and writes a visual HTML
 * gallery to the scratchpad for eyeballing.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { Chart } from '../src/pptx/model.js';

const SCRATCH = tmpdir();

// --- minimal SVG-DOM shim -------------------------------------------------
class FakeNode {
  ns?: string;
  tag: string;
  attrs: Record<string, string> = {};
  style: Record<string, string> = {};
  children: FakeNode[] = [];
  textContent = '';
  constructor(tag: string, ns?: string) {
    this.tag = tag;
    this.ns = ns;
  }
  setAttribute(k: string, v: string) {
    this.attrs[k] = v;
  }
  setAttributeNS(_ns: string, k: string, v: string) {
    this.attrs[k] = v;
  }
  appendChild(c: FakeNode) {
    this.children.push(c);
    return c;
  }
  get outerHTML(): string {
    const a = Object.entries(this.attrs).map(([k, v]) => `${k}="${v}"`);
    const styleStr = Object.entries(this.style)
      .map(([k, v]) => `${k.replace(/[A-Z]/g, (m) => '-' + m.toLowerCase())}:${v}`)
      .join(';');
    if (styleStr) a.push(`style="${styleStr}"`);
    const open = `<${this.tag}${a.length ? ' ' + a.join(' ') : ''}>`;
    const inner = this.children.map((c) => c.outerHTML).join('') + escapeText(this.textContent);
    return `${open}${inner}</${this.tag}>`;
  }
  querySelector() {
    return null;
  }
}
function escapeText(s: string) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

const fakeDoc = {
  createElementNS: (ns: string, tag: string) => new FakeNode(tag, ns),
  createElement: (tag: string) => new FakeNode(tag),
};

let renderChart: typeof import('../src/pptx/charts/render.js').renderChart;

beforeAll(async () => {
  (globalThis as Record<string, unknown>).document = fakeDoc;
  renderChart = (await import('../src/pptx/charts/render.js')).renderChart;
});
afterAll(() => {
  delete (globalThis as Record<string, unknown>).document;
});

const C = (hex: string) => ({ hex });

const samples: Array<{ title: string; chart: Chart }> = [
  {
    title: 'Clustered column',
    chart: {
      kind: 'bar',
      grouping: 'clustered',
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      legend: 'b',
      title: 'Revenue vs Cost',
      series: [
        { name: 'Revenue', values: [120, 180, 150, 220], color: C('4472C4') },
        { name: 'Cost', values: [80, 110, 95, 130], color: C('ED7D31') },
      ],
    },
  },
  {
    title: 'Stacked bar (horizontal)',
    chart: {
      kind: 'bar',
      barHorizontal: true,
      grouping: 'stacked',
      categories: ['North', 'South', 'East', 'West'],
      legend: 'b',
      series: [
        { name: 'A', values: [30, 20, 25, 40], color: C('5B9BD5') },
        { name: 'B', values: [15, 25, 10, 20], color: C('70AD47') },
      ],
    },
  },
  {
    title: 'Line',
    chart: {
      kind: 'line',
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May'],
      legend: 'b',
      series: [
        { name: 'Visitors', values: [10, 25, 18, 32, 28], color: C('4472C4') },
        { name: 'Signups', values: [4, 9, 7, 14, 16], color: C('ED7D31') },
      ],
    },
  },
  {
    title: 'Area',
    chart: {
      kind: 'area',
      categories: ['2019', '2020', '2021', '2022', '2023'],
      legend: 'b',
      series: [{ name: 'Total', values: [40, 55, 70, 65, 90], color: C('70AD47') }],
    },
  },
  {
    title: 'Pie',
    chart: {
      kind: 'pie',
      categories: ['Chrome', 'Safari', 'Edge', 'Firefox'],
      legend: 'r',
      series: [{ name: 'Share', values: [62, 19, 11, 8] }],
    },
  },
  {
    title: 'Doughnut',
    chart: {
      kind: 'doughnut',
      holeSize: 0.55,
      categories: ['A', 'B', 'C'],
      legend: 'r',
      series: [{ name: 'Mix', values: [5, 3, 2] }],
    },
  },
  {
    title: 'Scatter',
    chart: {
      kind: 'scatter',
      categories: [],
      legend: 'b',
      series: [
        { name: 'Set 1', values: [4, 8, 6, 10, 7], xValues: [1, 2, 3, 4, 5], color: C('4472C4') },
      ],
    },
  },
];

describe('chart rendering', () => {
  it('emits an SVG with the right primitives per kind', () => {
    const bar = renderChart(samples[0]!.chart, 400, 260) as unknown as FakeNode;
    const html = bar.outerHTML;
    expect(bar.tag).toBe('svg');
    expect(html).toContain('<rect'); // bars
    expect(html).toContain('Revenue vs Cost'); // title
    expect(html).toContain('Q1'); // category label

    const pie = renderChart(samples[4]!.chart, 400, 260) as unknown as FakeNode;
    expect(pie.outerHTML).toContain('<path'); // slices
    expect(pie.outerHTML).toContain('Chrome'); // legend entry

    const line = renderChart(samples[2]!.chart, 400, 260) as unknown as FakeNode;
    expect(line.outerHTML).toContain('<polyline');

    const area = renderChart(samples[3]!.chart, 400, 260) as unknown as FakeNode;
    expect(area.outerHTML).toContain('<polygon');
  });

  it('writes a visual gallery to the scratchpad', () => {
    const cards = samples
      .map((s) => {
        const svg = (renderChart(s.chart, 380, 240) as unknown as FakeNode).outerHTML;
        return `<figure><figcaption>${s.title}</figcaption><div class="box">${svg}</div></figure>`;
      })
      .join('\n');
    const doc = `<!doctype html><meta charset="utf-8"><title>Chart renderer output</title>
<style>
  body{font:14px system-ui;margin:24px;background:#fafafa;color:#222}
  h1{font-size:18px}
  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(400px,1fr));gap:20px}
  figure{margin:0;background:#fff;border:1px solid #e3e3e3;border-radius:8px;padding:12px}
  figcaption{font-weight:600;margin-bottom:8px}
  .box{width:380px;height:240px;position:relative;border:1px solid #f0f0f0}
  .box svg{display:block}
</style>
<h1>PPT-viewer chart renderer — sample output</h1>
<div class="grid">${cards}</div>`;
    try {
      writeFileSync(join(SCRATCH, 'chart-gallery.html'), doc);
    } catch {
      /* gallery output is a dev aid; never fail the suite on a write error */
    }
    expect(doc.length).toBeGreaterThan(1000);
  });
});
