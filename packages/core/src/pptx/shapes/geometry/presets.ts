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

function roundRect(w: number, h: number, adj: Adjust): string {
  const a = adj['adj'] ?? adj['adj1'] ?? 0.16667;
  const r = Math.min(w, h) * Math.min(0.5, a);
  return [
    `M${r},0`,
    `L${w - r},0`,
    `Q${w},0 ${w},${r}`,
    `L${w},${h - r}`,
    `Q${w},${h} ${w - r},${h}`,
    `L${r},${h}`,
    `Q0,${h} 0,${h - r}`,
    `L0,${r}`,
    `Q0,0 ${r},0`,
    'Z',
  ].join(' ');
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
  line: (w, h) => `M0,0 L${w},${h}`,
  straightConnector1: (w, h) => `M0,0 L${w},${h}`,
};

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
export const OPEN_PRESETS = new Set(['line', 'straightConnector1']);

/** True if we have an exact generator for this preset. */
export function hasPreset(preset: string): boolean {
  return preset in GENERATORS;
}

/** Generate an SVG path `d` for a preset shape in a w×h pixel box. */
export function presetPath(preset: string, w: number, h: number, adj: Adjust): string {
  const gen = GENERATORS[preset] ?? rect;
  return gen(w, h, adj);
}
