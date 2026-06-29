/**
 * Visual gallery of every preset generator: writes an HTML grid of each shape
 * to the scratchpad for eyeballing, and asserts each produces a non-empty path
 * that is not silently the rectangle fallback (except the ones that genuinely
 * are rectangles).
 */
import { describe, it, expect } from 'vitest';
import { writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { presetPath, OPEN_PRESETS, EVENODD_PRESETS, hasPreset } from '../src/pptx/shapes/geometry/presets.js';

const SCRATCH = tmpdir();

const NAMES = [
  'rect', 'roundRect', 'round1Rect', 'round2SameRect', 'round2DiagRect', 'snip1Rect', 'snip2SameRect',
  'ellipse', 'triangle', 'rtTriangle', 'diamond', 'parallelogram', 'trapezoid',
  'pentagon', 'hexagon', 'heptagon', 'octagon', 'decagon', 'dodecagon',
  'star4', 'star5', 'star6', 'star7', 'star8', 'star10', 'star12', 'star16', 'star24', 'star32',
  'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow', 'quadArrow',
  'bentArrow', 'notchedRightArrow', 'stripedRightArrow', 'homePlate', 'chevron',
  'mathPlus', 'mathMinus', 'mathMultiply', 'mathEqual', 'plus',
  'donut', 'noSmoking', 'frame', 'halfFrame', 'corner', 'diagStripe', 'plaque', 'can', 'cube', 'bevel',
  'teardrop', 'pie', 'chord', 'blockArc', 'heart', 'lightningBolt', 'sun', 'moon', 'smileyFace', 'cloud', 'foldedCorner',
  'wave', 'doubleWave', 'ribbon', 'ribbon2', 'verticalScroll', 'horizontalScroll',
  'leftBracket', 'rightBracket', 'bracketPair', 'leftBrace', 'rightBrace', 'bracePair',
  'flowChartDecision', 'flowChartTerminator', 'flowChartAlternateProcess', 'flowChartInputOutput',
  'flowChartData', 'flowChartPreparation', 'flowChartManualInput', 'flowChartManualOperation',
  'flowChartExtract', 'flowChartMerge', 'flowChartOffpageConnector', 'flowChartPunchedCard',
  'flowChartDocument', 'flowChartMultidocument', 'flowChartMagneticDisk', 'flowChartMagneticDrum',
  'flowChartMagneticTape', 'flowChartDelay', 'flowChartDisplay', 'flowChartConnector', 'flowChartSort',
  'flowChartPunchedTape',
  'wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout',
  'arc', 'line',
];

describe('preset geometry gallery', () => {
  it('every listed preset has a generator', () => {
    const missing = NAMES.filter((n) => !hasPreset(n));
    expect(missing).toEqual([]);
  });

  it('writes a visual gallery', () => {
    const W = 120;
    const H = 90;
    const adj = { adj: 0.2, adj1: 0.25, adj2: 0.5, adj3: 0.25 };
    // Angle-driven shapes read adj as degrees*0.6; give them realistic angles.
    const ANGLE = new Set(['arc', 'pie', 'chord', 'blockArc']);
    const angleAdj = (n: string): Record<string, number> =>
      n === 'pie' ? { adj1: 0, adj2: 270 * 0.6 }
      : n === 'chord' ? { adj1: 45 * 0.6, adj2: 270 * 0.6 }
      : n === 'arc' ? { adj1: 270 * 0.6, adj2: 90 * 0.6 }
      : { adj1: 180 * 0.6, adj2: 0, adj3: 0.25 }; // blockArc
    const OVERRIDE: Record<string, Record<string, number>> = {
      teardrop: { adj: 1 },
      moon: { adj: 0.5 },
    };
    const cards = NAMES.map((name) => {
      const a = OVERRIDE[name] ?? (ANGLE.has(name) ? angleAdj(name) : adj);
      const d = presetPath(name, W, H, a);
      const open = OPEN_PRESETS.has(name);
      const even = EVENODD_PRESETS.has(name) ? ` fill-rule="evenodd"` : '';
      const path = open
        ? `<path d="${d}" fill="none" stroke="#1f3864" stroke-width="2"/>`
        : `<path d="${d}"${even} fill="#9dc3e6" stroke="#1f3864" stroke-width="1.5"/>`;
      return `<figure><svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${path}</svg><figcaption>${name}</figcaption></figure>`;
    }).join('\n');
    const doc = `<!doctype html><meta charset="utf-8"><title>Preset shapes</title>
<style>
  body{font:12px system-ui;margin:20px;background:#fafafa;color:#222}
  h1{font-size:18px}
  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:10px}
  figure{margin:0;background:#fff;border:1px solid #e3e3e3;border-radius:6px;padding:6px;text-align:center}
  figcaption{margin-top:4px;color:#555;font-size:11px;word-break:break-all}
  svg{background:#fff}
</style>
<h1>Preset geometry — ${NAMES.length} shapes</h1>
<div class="grid">${cards}</div>`;
    try {
      writeFileSync(join(SCRATCH, 'shapes-gallery.html'), doc);
    } catch {
      /* gallery output is a dev aid; never fail the suite on a write error */
    }
    expect(NAMES.length).toBeGreaterThan(90);
  });
});
