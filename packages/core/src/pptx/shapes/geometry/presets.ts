/**
 * Preset-geometry -> SVG path generator.
 *
 * OOXML defines ~187 preset shapes (`prstGeom`). We implement the common ones
 * exactly and fall back to a rectangle for the rest (tracked so we can grow the
 * table over time). Paths are generated in the shape's own pixel box (0..w, 0..h).
 */

/** Adjustment values are fractions 0..1 (already converted from guide units). */
export type Adjust = Record<string, number>;

function rect(w: number, h: number): string {
  return `M0,0 L${w},0 L${w},${h} L0,${h} Z`;
}

function ellipse(w: number, h: number): string {
  const rx = w / 2;
  const ry = h / 2;
  return `M0,${ry} A${rx},${ry} 0 1 0 ${w},${ry} A${rx},${ry} 0 1 0 0,${ry} Z`;
}

/** Rounded rectangle with independent per-corner radii (px), rounded with arcs. */
function cornerRect(w: number, h: number, tl: number, tr: number, br: number, bl: number): string {
  const cap = Math.min(w, h) / 2;
  const c = (v: number) => Math.max(0, Math.min(v, cap));
  tl = c(tl); tr = c(tr); br = c(br); bl = c(bl);
  return [
    `M${tl},0`,
    `L${w - tr},0`,
    tr ? `Q${w},0 ${w},${tr}` : '',
    `L${w},${h - br}`,
    br ? `Q${w},${h} ${w - br},${h}` : '',
    `L${bl},${h}`,
    bl ? `Q0,${h} 0,${h - bl}` : '',
    `L0,${tl}`,
    tl ? `Q0,0 ${tl},0` : '',
    'Z',
  ]
    .filter(Boolean)
    .join(' ');
}

/** Rectangle with independent per-corner straight-cut (chamfered) corners. */
function snipRect(w: number, h: number, tl: number, tr: number, br: number, bl: number): string {
  const cap = Math.min(w, h) / 2;
  const c = (v: number) => Math.max(0, Math.min(v, cap));
  tl = c(tl); tr = c(tr); br = c(br); bl = c(bl);
  return [
    `M${tl},0`,
    `L${w - tr},0`,
    tr ? `L${w},${tr}` : '',
    `L${w},${h - br}`,
    br ? `L${w - br},${h}` : '',
    `L${bl},${h}`,
    bl ? `L0,${h - bl}` : '',
    `L0,${tl}`,
    'Z',
  ]
    .filter(Boolean)
    .join(' ');
}

/** Corner radius in px from an adjust fraction of the short side (default 1/6). */
function cornerRadius(w: number, h: number, adj: number | undefined, def = 0.16667): number {
  return Math.min(w, h) * (adj ?? def);
}

function roundRect(w: number, h: number, adj: Adjust): string {
  const r = cornerRadius(w, h, adj['adj'] ?? adj['adj1']);
  return cornerRect(w, h, r, r, r, r);
}

function poly(points: Array<[number, number]>): string {
  return (
    points.map(([x, y], i) => `${i === 0 ? 'M' : 'L'}${x},${y}`).join(' ') + ' Z'
  );
}

const GENERATORS: Record<string, (w: number, h: number, adj: Adjust) => string> = {
  rect: rect,
  flowChartProcess: rect,
  ellipse: ellipse,
  flowChartConnector: ellipse,
  roundRect: roundRect,
  // round1Rect rounds only the top-right corner.
  round1Rect: (w, h, adj) => {
    const r = cornerRadius(w, h, adj['adj']);
    return cornerRect(w, h, 0, r, 0, 0);
  },
  // round2SameRect rounds the two top corners by adj1, the two bottom by adj2.
  round2SameRect: (w, h, adj) => {
    const r1 = cornerRadius(w, h, adj['adj1']);
    const r2 = cornerRadius(w, h, adj['adj2'], 0);
    return cornerRect(w, h, r1, r1, r2, r2);
  },
  // round2DiagRect rounds one diagonal pair (TL+BR) by adj1, the other by adj2.
  round2DiagRect: (w, h, adj) => {
    const r1 = cornerRadius(w, h, adj['adj1']);
    const r2 = cornerRadius(w, h, adj['adj2'], 0);
    return cornerRect(w, h, r1, r2, r1, r2);
  },
  // snip*Rect: corners cut straight instead of rounded (top-right for snip1Rect).
  snip1Rect: (w, h, adj) => snipRect(w, h, 0, cornerRadius(w, h, adj['adj']), 0, 0),
  snip2SameRect: (w, h, adj) => {
    const r1 = cornerRadius(w, h, adj['adj1']);
    const r2 = cornerRadius(w, h, adj['adj2'], 0);
    return snipRect(w, h, r1, r1, r2, r2);
  },
  triangle: (w, h) =>
    poly([
      [w / 2, 0],
      [w, h],
      [0, h],
    ]),
  rtTriangle: (w, h) =>
    poly([
      [0, 0],
      [0, h],
      [w, h],
    ]),
  diamond: (w, h) =>
    poly([
      [w / 2, 0],
      [w, h / 2],
      [w / 2, h],
      [0, h / 2],
    ]),
  parallelogram: (w, h, adj) => {
    const a = (adj['adj'] ?? 0.25) * w;
    return poly([
      [a, 0],
      [w, 0],
      [w - a, h],
      [0, h],
    ]);
  },
  trapezoid: (w, h, adj) => {
    const a = (adj['adj'] ?? 0.25) * w;
    return poly([
      [a, 0],
      [w - a, 0],
      [w, h],
      [0, h],
    ]);
  },
  pentagon: (w, h) => regularPolygon(w, h, 5, -90),
  hexagon: (w, h) => regularPolygon(w, h, 6, 0),
  octagon: (w, h) => regularPolygon(w, h, 8, 22.5),
  rightArrow: (w, h, adj) => {
    const aw = (adj['adj1'] ?? 0.5) * h; // head half-height fraction of h -> body inset
    const al = (adj['adj2'] ?? 0.5) * w; // arrow head length
    const bodyTop = (h - aw) / 2;
    return poly([
      [0, bodyTop],
      [w - al, bodyTop],
      [w - al, 0],
      [w, h / 2],
      [w - al, h],
      [w - al, h - bodyTop],
      [0, h - bodyTop],
    ]);
  },
  leftArrow: (w, h, adj) => {
    const aw = (adj['adj1'] ?? 0.5) * h;
    const al = (adj['adj2'] ?? 0.5) * w;
    const bodyTop = (h - aw) / 2;
    return poly([
      [w, bodyTop],
      [al, bodyTop],
      [al, 0],
      [0, h / 2],
      [al, h],
      [al, h - bodyTop],
      [w, h - bodyTop],
    ]);
  },
  // Pentagon/arrow callout (a rectangle with a pointed right edge).
  homePlate: (w, h, adj) => {
    const dx = Math.min(w, h) * (adj['adj'] ?? 0.5);
    const x1 = Math.max(0, w - dx);
    return poly([
      [0, 0],
      [x1, 0],
      [w, h / 2],
      [x1, h],
      [0, h],
    ]);
  },
  chevron: (w, h, adj) => {
    const dx = Math.min(w, h) * (adj['adj'] ?? 0.5);
    return poly([
      [0, 0],
      [w - dx, 0],
      [w, h / 2],
      [w - dx, h],
      [0, h],
      [dx, h / 2],
    ]);
  },
  // Open elliptical arc (a stroked curve, not a filled region). adj1/adj2 are
  // the start/end angles in 60000ths of a degree (clockwise from 3 o'clock).
  arc: (w, h, adj) => {
    const rx = w / 2;
    const ry = h / 2;
    const a1 = adjDeg(adj['adj1'], 270);
    const a2 = adjDeg(adj['adj2'], 0);
    const r1 = (a1 * Math.PI) / 180;
    const r2 = (a2 * Math.PI) / 180;
    const x1 = rx + rx * Math.cos(r1);
    const y1 = ry + ry * Math.sin(r1);
    const x2 = rx + rx * Math.cos(r2);
    const y2 = ry + ry * Math.sin(r2);
    const sweep = (((a2 - a1) % 360) + 360) % 360; // clockwise degrees swept
    const largeArc = sweep > 180 ? 1 : 0;
    return `M${x1},${y1} A${rx},${ry} 0 ${largeArc} 1 ${x2},${y2}`;
  },
  line: (w, h) => `M0,0 L${w},${h}`,
  straightConnector1: (w, h) => `M0,0 L${w},${h}`,
};

/**
 * Convert an angle-typed adjustment to degrees. Adjust values arrive already
 * divided by 100000 (the percentage convention), but angle guides are authored
 * in 60000ths of a degree, so undo that scaling: deg = adj * 100000 / 60000.
 */
function adjDeg(adj: number | undefined, defaultDeg: number): number {
  if (adj === undefined) return defaultDeg;
  return (adj * 100000) / 60000;
}

function regularPolygon(w: number, h: number, n: number, startDeg: number): string {
  const cx = w / 2;
  const cy = h / 2;
  const rx = w / 2;
  const ry = h / 2;
  const pts: Array<[number, number]> = [];
  for (let i = 0; i < n; i++) {
    const ang = ((startDeg + (360 / n) * i) * Math.PI) / 180;
    pts.push([cx + rx * Math.cos(ang), cy + ry * Math.sin(ang)]);
  }
  return poly(pts);
}

/** Preset names we render as just an outline stroke (no closed fill area). */
export const OPEN_PRESETS = new Set(['line', 'straightConnector1', 'arc']);

/** True if we have an exact generator for this preset. */
export function hasPreset(preset: string): boolean {
  return preset in GENERATORS;
}

/** Generate an SVG path `d` for a preset shape in a w×h pixel box. */
export function presetPath(preset: string, w: number, h: number, adj: Adjust): string {
  const gen = GENERATORS[preset] ?? rect;
  return gen(w, h, adj);
}
