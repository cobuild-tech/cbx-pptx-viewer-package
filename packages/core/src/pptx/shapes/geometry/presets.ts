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

/** Polyline without an implicit close (for open shapes like brackets). */
function polyOpen(points: Array<[number, number]>): string {
  return points.map(([x, y], i) => `${i === 0 ? 'M' : 'L'}${x},${y}`).join(' ');
}

/**
 * Star polygon: `n` outer points on the box ellipse alternating with `n` inner
 * points at `innerRatio` of the radius. First point at 12 o'clock.
 */
function star(w: number, h: number, n: number, innerRatio: number): string {
  const cx = w / 2;
  const cy = h / 2;
  const pts: Array<[number, number]> = [];
  for (let i = 0; i < n * 2; i++) {
    const ang = ((-90 + (180 / n) * i) * Math.PI) / 180;
    const r = i % 2 === 0 ? 1 : innerRatio;
    pts.push([cx + cx * r * Math.cos(ang), cy + cy * r * Math.sin(ang)]);
  }
  return poly(pts);
}

/** Point on the box ellipse at `deg` (clockwise from 3 o'clock). */
function ellPt(w: number, h: number, deg: number): [number, number] {
  const r = (deg * Math.PI) / 180;
  return [w / 2 + (w / 2) * Math.cos(r), h / 2 + (h / 2) * Math.sin(r)];
}

/** Elliptical-arc command from `a0`->`a1` degrees on an rx/ry ellipse centered at cx,cy. */
function arcCmd(cx: number, cy: number, rx: number, ry: number, a0: number, a1: number): string {
  const p1 = [cx + rx * Math.cos((a1 * Math.PI) / 180), cy + ry * Math.sin((a1 * Math.PI) / 180)];
  const large = Math.abs(a1 - a0) > 180 ? 1 : 0;
  const sweep = a1 > a0 ? 1 : 0;
  return `A${rx},${ry} 0 ${large} ${sweep} ${p1[0]},${p1[1]}`;
}

/** A full circle/ellipse subpath as two arcs (so it can be combined for holes). */
function ellipsePath(cx: number, cy: number, rx: number, ry: number): string {
  return `M${cx - rx},${cy} A${rx},${ry} 0 1 0 ${cx + rx},${cy} A${rx},${ry} 0 1 0 ${cx - rx},${cy} Z`;
}

/** Vertical (up/down) block arrow filling the box. */
function vArrow(w: number, h: number, up: boolean): string {
  const sh = w * 0.2; // shaft half-width
  const hl = h * 0.45; // arrowhead length
  const cx = w / 2;
  const pts: Array<[number, number]> = up
    ? [[cx, 0], [w, hl], [cx + sh, hl], [cx + sh, h], [cx - sh, h], [cx - sh, hl], [0, hl]]
    : [[cx, h], [w, h - hl], [cx + sh, h - hl], [cx + sh, 0], [cx - sh, 0], [cx - sh, h - hl], [0, h - hl]];
  return poly(pts);
}

/** Plus/cross shape; `t` is the arm half-thickness as a fraction of the short side. */
function cross(w: number, h: number, t: number): string {
  const tx = w * t;
  const ty = h * t;
  const cx = w / 2;
  const cy = h / 2;
  return poly([
    [cx - tx, 0], [cx + tx, 0], [cx + tx, cy - ty], [w, cy - ty], [w, cy + ty],
    [cx + tx, cy + ty], [cx + tx, h], [cx - tx, h], [cx - tx, cy + ty], [0, cy + ty],
    [0, cy - ty], [cx - tx, cy - ty],
  ]);
}

/** Diagonal "X" (multiply) as two crossed bands. */
function xMark(w: number, h: number, t: number): string {
  return [
    poly([[0, t], [t, 0], [w, h - t], [w - t, h]]),
    poly([[w - t, 0], [w, t], [t, h], [0, h - t]]),
  ].join(' ');
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
  heptagon: (w, h) => regularPolygon(w, h, 7, -90),
  decagon: (w, h) => regularPolygon(w, h, 10, 0),
  dodecagon: (w, h) => regularPolygon(w, h, 12, 15),

  // --- Stars. Inner-radius ratios approximate PowerPoint's defaults. ---
  star4: (w, h) => star(w, h, 4, 0.38),
  star5: (w, h) => star(w, h, 5, 0.38),
  star6: (w, h) => star(w, h, 6, 0.55),
  star7: (w, h) => star(w, h, 7, 0.52),
  star8: (w, h) => star(w, h, 8, 0.62),
  star10: (w, h) => star(w, h, 10, 0.68),
  star12: (w, h) => star(w, h, 12, 0.7),
  star16: (w, h) => star(w, h, 16, 0.75),
  star24: (w, h) => star(w, h, 24, 0.83),
  star32: (w, h) => star(w, h, 32, 0.88),

  // --- Block arrows. Defaults chosen to read clearly at any aspect. ---
  upArrow: (w, h) => vArrow(w, h, true),
  downArrow: (w, h) => vArrow(w, h, false),
  leftRightArrow: (w, h) => {
    const hl = w * 0.25;
    const sh = h * 0.2;
    return poly([
      [0, h / 2], [hl, 0], [hl, h / 2 - sh], [w - hl, h / 2 - sh], [w - hl, 0],
      [w, h / 2], [w - hl, h], [w - hl, h / 2 + sh], [hl, h / 2 + sh], [hl, h],
    ]);
  },
  upDownArrow: (w, h) => {
    const hl = h * 0.25;
    const sh = w * 0.2;
    return poly([
      [w / 2, 0], [w, hl], [w / 2 + sh, hl], [w / 2 + sh, h - hl], [w, h - hl],
      [w / 2, h], [0, h - hl], [w / 2 - sh, h - hl], [w / 2 - sh, hl], [0, hl],
    ]);
  },
  quadArrow: (w, h) => {
    const s = w * 0.12; // shaft half-width
    const t = h * 0.12;
    const hl = Math.min(w, h) * 0.28; // arrowhead length
    const hw = Math.min(w, h) * 0.22; // arrowhead half-width
    const cx = w / 2;
    const cy = h / 2;
    return poly([
      [cx, 0], [cx + hw, hl], [cx + s, hl], [cx + s, cy - t], [w - hl, cy - t],
      [w - hl, cy - hw], [w, cy], [w - hl, cy + hw], [w - hl, cy + t], [cx + s, cy + t],
      [cx + s, h - hl], [cx + hw, h - hl], [cx, h], [cx - hw, h - hl], [cx - s, h - hl],
      [cx - s, cy + t], [hl, cy + t], [hl, cy + hw], [0, cy], [hl, cy - hw],
      [hl, cy - t], [cx - s, cy - t], [cx - s, hl], [cx - hw, hl],
    ]);
  },
  bentArrow: (w, h) => {
    // Shaft rising from the bottom then turning right into the arrowhead.
    const t = Math.min(w, h) * 0.22; // shaft thickness
    const hl = w * 0.28; // head length
    const hw = h * 0.4; // head half-height
    const ay = h * 0.32; // arm vertical centre
    return poly([
      [0, h], [0, h * 0.5], [t, h * 0.5], [t, ay + t / 2],
      [w - hl, ay + t / 2], [w - hl, ay + hw], [w, ay], [w - hl, ay - hw],
      [w - hl, ay - t / 2], [t + t, ay - t / 2],
    ]);
  },
  notchedRightArrow: (w, h) => {
    const hl = w * 0.4;
    const sh = h * 0.25;
    const notch = hl * 0.4;
    return poly([
      [0, h / 2 - sh], [w - hl, h / 2 - sh], [w - hl, 0], [w, h / 2],
      [w - hl, h], [w - hl, h / 2 + sh], [0, h / 2 + sh], [notch, h / 2],
    ]);
  },
  stripedRightArrow: (w, h) => {
    // Outline silhouette of the striped arrow (the stripes are decorative).
    const hl = w * 0.4;
    const sh = h * 0.25;
    return poly([
      [0, h / 2 - sh], [w - hl, h / 2 - sh], [w - hl, 0], [w, h / 2],
      [w - hl, h], [w - hl, h / 2 + sh], [0, h / 2 + sh],
    ]);
  },

  // --- Math symbols ---
  mathPlus: (w, h) => cross(w, h, 0.23),
  plus: (w, h, adj) => cross(w, h, adj['adj'] ?? 0.25),
  mathMinus: (w, h) => {
    const t = h * 0.23;
    return poly([[0, h / 2 - t], [w, h / 2 - t], [w, h / 2 + t], [0, h / 2 + t]]);
  },
  mathMultiply: (w, h) => xMark(w, h, Math.min(w, h) * 0.16),
  mathEqual: (w, h) => {
    const t = h * 0.14;
    const g = h * 0.16;
    return [
      poly([[0, h / 2 - g - t], [w, h / 2 - g - t], [w, h / 2 - g + t], [0, h / 2 - g + t]]),
      poly([[0, h / 2 + g - t], [w, h / 2 + g - t], [w, h / 2 + g + t], [0, h / 2 + g + t]]),
    ].join(' ');
  },

  // --- Basic regions ---
  donut: (w, h, adj) => {
    const r = (adj['adj'] ?? 0.25);
    const inner = Math.max(0, 0.5 - r);
    return ellipsePath(w / 2, h / 2, w / 2, h / 2) + ' ' +
      ellipsePath(w / 2, h / 2, w * inner, h * inner);
  },
  noSmoking: (w, h, adj) => {
    const r = (adj['adj'] ?? 0.18);
    const inner = Math.max(0, 0.5 - r);
    return ellipsePath(w / 2, h / 2, w / 2, h / 2) + ' ' +
      ellipsePath(w / 2, h / 2, w * inner, h * inner);
  },
  frame: (w, h, adj) => {
    const t = Math.min(w, h) * (adj['adj1'] ?? adj['adj'] ?? 0.1);
    return rect(w, h) + ' ' +
      poly([[t, t], [t, h - t], [w - t, h - t], [w - t, t]]);
  },
  halfFrame: (w, h, adj) => {
    const t = Math.min(w, h) * (adj['adj2'] ?? 0.1);
    return poly([[0, 0], [w, 0], [w - t, t], [t, t], [t, h - t], [0, h]]);
  },
  corner: (w, h, adj) => {
    const tw = w * (adj['adj1'] ?? 0.5);
    const th = h * (adj['adj2'] ?? 0.5);
    return poly([[0, 0], [tw, 0], [tw, h - th], [w, h - th], [w, h], [0, h]]);
  },
  diagStripe: (w, h, adj) => {
    const f = adj['adj'] ?? 0.5;
    return poly([[0, h * f], [w * f, 0], [w, 0], [0, h]]);
  },
  plaque: (w, h, adj) => {
    const c = Math.min(w, h) * (adj['adj'] ?? 0.16);
    return [
      `M${c},0`,
      `L${w - c},0`, `Q${w - c},${c} ${w},${c}`,
      `L${w},${h - c}`, `Q${w - c},${h - c} ${w - c},${h}`,
      `L${c},${h}`, `Q${c},${h - c} 0,${h - c}`,
      `L0,${c}`, `Q${c},${c} ${c},0`, 'Z',
    ].join(' ');
  },
  can: (w, h) => {
    const ry = Math.min(h * 0.12, w * 0.5);
    return [
      `M0,${ry}`,
      `A${w / 2},${ry} 0 0 1 ${w},${ry}`,
      `L${w},${h - ry}`,
      `A${w / 2},${ry} 0 0 1 0,${h - ry}`,
      'Z',
      // Top ellipse rim.
      `M0,${ry} A${w / 2},${ry} 0 0 0 ${w},${ry}`,
    ].join(' ');
  },
  cube: (w, h, adj) => {
    const d = Math.min(w, h) * (adj['adj'] ?? 0.25);
    return [
      // Front face + top + right side as one silhouette.
      poly([[0, d], [d, 0], [w, 0], [w, h - d], [w - d, h], [0, h]]),
      // Top edge.
      polyOpen([[0, d], [w - d, d], [w, 0]]),
      // Right edge.
      polyOpen([[w - d, d], [w - d, h]]),
    ].join(' ');
  },
  bevel: (w, h) => rect(w, h),
  teardrop: (w, h, adj) => {
    const f = adj['adj'] ?? 1; // how far the tail extends toward the top-right corner
    const cx = w / 2, cy = h / 2;
    // A circle with the top-right quadrant replaced by a point at (cx+f*cx, cy-f*cy).
    return [
      `M${cx},0`,
      `A${cx},${cy} 0 0 0 0,${cy}`,
      `A${cx},${cy} 0 0 0 ${cx},${h}`,
      `A${cx},${cy} 0 0 0 ${w},${cy}`,
      `L${cx + cx * f},${cy - cy * f}`,
      'Z',
    ].join(' ');
  },
  pie: (w, h, adj) => {
    const a1 = adjDeg(adj['adj1'], 0);
    const a2 = adjDeg(adj['adj2'], 270);
    const cx = w / 2, cy = h / 2;
    const [sx, sy] = ellPt(w, h, a1);
    return `M${cx},${cy} L${sx},${sy} ${arcCmd(cx, cy, w / 2, h / 2, a1, a2 < a1 ? a2 + 360 : a2)} Z`;
  },
  chord: (w, h, adj) => {
    const a1 = adjDeg(adj['adj1'], 45);
    const a2 = adjDeg(adj['adj2'], 270);
    const [sx, sy] = ellPt(w, h, a1);
    return `M${sx},${sy} ${arcCmd(w / 2, h / 2, w / 2, h / 2, a1, a2 < a1 ? a2 + 360 : a2)} Z`;
  },
  blockArc: (w, h, adj) => {
    const a1 = adjDeg(adj['adj1'], 180);
    const a2 = adjDeg(adj['adj2'], 0);
    const thick = adj['adj3'] ?? 0.25;
    const cx = w / 2, cy = h / 2;
    const ir = 0.5 - thick;
    const e1 = a2 < a1 ? a2 + 360 : a2;
    const [ox, oy] = ellPt(w, h, a1);
    const ix = cx + w * ir * Math.cos((e1 * Math.PI) / 180);
    const iy = cy + h * ir * Math.sin((e1 * Math.PI) / 180);
    return [
      `M${ox},${oy}`,
      arcCmd(cx, cy, w / 2, h / 2, a1, e1),
      `L${ix},${iy}`,
      arcCmd(cx, cy, w * ir, h * ir, e1, a1),
      'Z',
    ].join(' ');
  },

  // --- Decorative ---
  heart: (w, h) => {
    const x = w / 2;
    return [
      `M${x},${h * 0.28}`,
      `C${w * 0.5 - w * 0.1},${-h * 0.05} 0,${h * 0.12} 0,${h * 0.35}`,
      `C0,${h * 0.6} ${x * 0.7},${h * 0.78} ${x},${h}`,
      `C${x + x * 0.3},${h * 0.78} ${w},${h * 0.6} ${w},${h * 0.35}`,
      `C${w},${h * 0.12} ${w * 0.5 + w * 0.1},${-h * 0.05} ${x},${h * 0.28}`,
      'Z',
    ].join(' ');
  },
  lightningBolt: (w, h) =>
    poly([
      [w * 0.32, 0], [w * 0.66, 0], [w * 0.5, h * 0.38], [w * 0.78, h * 0.38],
      [w * 0.34, h], [w * 0.46, h * 0.55], [w * 0.2, h * 0.55],
    ]),
  sun: (w, h) => {
    // 8-ray sun: a star with deep notches.
    return star(w, h, 8, 0.62);
  },
  moon: (w, h, adj) => {
    const f = adj['adj'] ?? 0.5; // inner-cut radius as a fraction of width
    return [
      `M${w},0`,
      `A${w},${h / 2} 0 0 0 ${w},${h}`, // outer edge bulging left through (0, h/2)
      `A${w * f},${h / 2} 0 0 1 ${w},0`, // inner concave edge carving the crescent
      'Z',
    ].join(' ');
  },
  smileyFace: (w, h) => ellipse(w, h),
  cloud: (w, h) => {
    // Five-lobe cloud built from arcs around the box.
    const lobes: Array<[number, number, number]> = [
      [0.25, 0.55, 0.25], [0.4, 0.32, 0.22], [0.62, 0.3, 0.24],
      [0.8, 0.5, 0.22], [0.68, 0.78, 0.24], [0.38, 0.8, 0.24],
      [0.16, 0.72, 0.2],
    ];
    return lobes
      .map(([cx, cy, r]) => ellipsePath(w * cx, h * cy, w * r, h * r * (w / h)))
      .join(' ');
  },
  foldedCorner: (w, h, adj) => {
    const c = Math.min(w, h) * (adj['adj'] ?? 0.16);
    return [
      poly([[0, 0], [w, 0], [w, h - c], [w - c, h], [0, h]]),
      polyOpen([[w - c, h], [w - c, h - c], [w, h - c]]),
    ].join(' ');
  },

  // --- Banners / scrolls / waves ---
  wave: (w, h, adj) => {
    const a = h * (adj['adj1'] ?? 0.13);
    return [
      `M0,${a}`,
      `C${w * 0.33},${-a} ${w * 0.66},${3 * a} ${w},${a}`,
      `L${w},${h - a}`,
      `C${w * 0.66},${h + a} ${w * 0.33},${h - 3 * a} 0,${h - a}`,
      'Z',
    ].join(' ');
  },
  doubleWave: (w, h, adj) => {
    const a = h * (adj['adj1'] ?? 0.1);
    return [
      `M0,${a}`,
      `C${w * 0.17},${-a} ${w * 0.33},${3 * a} ${w * 0.5},${a}`,
      `C${w * 0.66},${-a} ${w * 0.83},${3 * a} ${w},${a}`,
      `L${w},${h - a}`,
      `C${w * 0.83},${h + a} ${w * 0.66},${h - 3 * a} ${w * 0.5},${h - a}`,
      `C${w * 0.33},${h + a} ${w * 0.17},${h - 3 * a} 0,${h - a}`,
      'Z',
    ].join(' ');
  },
  ribbon: (w, h) => {
    const tail = w * 0.12;
    const f = h * 0.25;
    return poly([
      [tail, 0], [w - tail, 0], [w - tail, h - f], [w, h], [w - tail * 1.6, h - f * 0.6],
      [w - tail * 1.6, h], [tail * 1.6, h], [tail * 1.6, h - f * 0.6], [0, h], [tail, h - f],
    ]);
  },
  ribbon2: (w, h) => {
    const tail = w * 0.12;
    const f = h * 0.25;
    return poly([
      [tail, f], [w - tail, f], [w - tail, 0], [w, h * 0.3], [w - tail * 1.6, h * 0.18],
      [w - tail * 1.6, h - f], [tail * 1.6, h - f], [tail * 1.6, h * 0.18], [0, h * 0.3], [tail, 0],
    ]);
  },
  verticalScroll: (w, h) => {
    const c = Math.min(w, h) * 0.16;
    return [
      cornerRect(w - c, h, c, 0, 0, c),
      ellipsePath(w - c, c, c, c),
      ellipsePath(c, h - c, c, c),
    ].join(' ');
  },
  horizontalScroll: (w, h) => {
    const c = Math.min(w, h) * 0.16;
    return [
      cornerRect(w, h - c, c, c, 0, 0),
      ellipsePath(c, h - c, c, c),
      ellipsePath(w - c, c, c, c),
    ].join(' ');
  },

  // --- Brackets & braces (open, stroked outlines) ---
  leftBracket: (w, h, adj) => {
    const c = Math.min(h / 2, w) * (adj['adj'] ?? 0.5) || h * 0.16;
    return `M${w},0 Q0,0 0,${c} L0,${h - c} Q0,${h} ${w},${h}`;
  },
  rightBracket: (w, h, adj) => {
    const c = Math.min(h / 2, w) * (adj['adj'] ?? 0.5) || h * 0.16;
    return `M0,0 Q${w},0 ${w},${c} L${w},${h - c} Q${w},${h} 0,${h}`;
  },
  bracketPair: (w, h, adj) => {
    const c = Math.min(w, h) * (adj['adj'] ?? 0.16);
    return `M${c},0 Q0,0 0,${c} L0,${h - c} Q0,${h} ${c},${h}` + ' ' +
      `M${w - c},0 Q${w},0 ${w},${c} L${w},${h - c} Q${w},${h} ${w - c},${h}`;
  },
  leftBrace: (w, h) =>
    `M${w},0 Q${w / 2},0 ${w / 2},${h * 0.25} Q${w / 2},${h / 2} 0,${h / 2} ` +
    `Q${w / 2},${h / 2} ${w / 2},${h * 0.75} Q${w / 2},${h} ${w},${h}`,
  rightBrace: (w, h) =>
    `M0,0 Q${w / 2},0 ${w / 2},${h * 0.25} Q${w / 2},${h / 2} ${w},${h / 2} ` +
    `Q${w / 2},${h / 2} ${w / 2},${h * 0.75} Q${w / 2},${h} 0,${h}`,
  bracePair: (w, h) =>
    `M${w * 0.4},0 Q0,0 0,${h * 0.25} Q0,${h / 2} 0,${h / 2} Q0,${h} ${w * 0.4},${h}` + ' ' +
    `M${w * 0.6},0 Q${w},0 ${w},${h * 0.25} Q${w},${h / 2} ${w},${h / 2} Q${w},${h} ${w * 0.6},${h}`,

  // --- Flowchart symbols ---
  flowChartDecision: (w, h) => GENERATORS.diamond!(w, h, {}),
  flowChartTerminator: (w, h) => {
    const r = Math.min(w / 2, h / 2);
    return cornerRect(w, h, r, r, r, r);
  },
  flowChartAlternateProcess: (w, h) => {
    const r = Math.min(w, h) * 0.16;
    return cornerRect(w, h, r, r, r, r);
  },
  flowChartInputOutput: (w, h) => GENERATORS.parallelogram!(w, h, { adj: 0.25 }),
  flowChartData: (w, h) => GENERATORS.parallelogram!(w, h, { adj: 0.25 }),
  flowChartPreparation: (w, h) => {
    const d = w * 0.2;
    return poly([[d, 0], [w - d, 0], [w, h / 2], [w - d, h], [d, h], [0, h / 2]]);
  },
  flowChartManualInput: (w, h) =>
    poly([[0, h * 0.25], [w, 0], [w, h], [0, h]]),
  flowChartManualOperation: (w, h) =>
    poly([[0, 0], [w, 0], [w * 0.8, h], [w * 0.2, h]]),
  flowChartExtract: (w, h) => poly([[w / 2, 0], [w, h], [0, h]]),
  flowChartMerge: (w, h) => poly([[0, 0], [w, 0], [w / 2, h]]),
  flowChartOffpageConnector: (w, h) =>
    poly([[0, 0], [w, 0], [w, h * 0.6], [w / 2, h], [0, h * 0.6]]),
  flowChartPunchedCard: (w, h) => {
    const c = Math.min(w, h) * 0.2;
    return poly([[c, 0], [w, 0], [w, h], [0, h], [0, c]]);
  },
  flowChartDocument: (w, h) => {
    const a = h * 0.12;
    return [
      `M0,0 L${w},0 L${w},${h - a}`,
      `C${w * 0.66},${h - 3 * a} ${w * 0.33},${h + a} 0,${h - a}`,
      'Z',
    ].join(' ');
  },
  flowChartMultidocument: (w, h) => GENERATORS.flowChartDocument!(w, h, {}),
  flowChartPredefinedProcess: (w, h) => rect(w, h),
  flowChartInternalStorage: (w, h) => rect(w, h),
  flowChartMagneticDisk: (w, h) => GENERATORS.can!(w, h, {}),
  flowChartMagneticDrum: (w, h) => {
    const rx = Math.min(w * 0.12, h * 0.5);
    return [
      `M${rx},0 L${w - rx},0`,
      `A${rx},${h / 2} 0 0 1 ${w - rx},${h}`,
      `L${rx},${h}`,
      `A${rx},${h / 2} 0 0 1 ${rx},0`,
      'Z',
      `M${w - rx},0 A${rx},${h / 2} 0 0 0 ${w - rx},${h}`,
    ].join(' ');
  },
  flowChartMagneticTape: (w, h) => ellipse(w, h),
  flowChartDelay: (w, h) => {
    const r = h / 2;
    return `M0,0 L${w - r},0 A${r},${r} 0 0 1 ${w - r},${h} L0,${h} Z`;
  },
  flowChartDisplay: (w, h) => {
    const r = h / 2;
    return [
      `M0,${h / 2}`, `L${w * 0.15},0`, `L${w - r},0`,
      `A${r},${r} 0 0 1 ${w - r},${h}`, `L${w * 0.15},${h}`, 'Z',
    ].join(' ');
  },
  flowChartSort: (w, h) => poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]),
  flowChartCollate: (w, h) =>
    `M0,0 L${w},0 L0,${h} L${w},${h} Z`,
  flowChartPunchedTape: (w, h) => {
    const a = h * 0.1;
    return [
      `M0,${a}`,
      `C${w * 0.25},${-a} ${w * 0.75},${3 * a} ${w},${a}`,
      `L${w},${h - a}`,
      `C${w * 0.75},${h + a} ${w * 0.25},${h - 3 * a} 0,${h - a}`,
      'Z',
    ].join(' ');
  },

  // --- Callouts (rectangular bodies with a pointer wedge) ---
  wedgeRectCallout: (w, h) => {
    const bh = h * 0.75;
    return [
      poly([[0, 0], [w, 0], [w, bh], [0, bh]]),
      poly([[w * 0.25, bh], [w * 0.45, bh], [w * 0.2, h]]),
    ].join(' ');
  },
  wedgeRoundRectCallout: (w, h) => {
    const bh = h * 0.75;
    const r = Math.min(w, bh) * 0.12;
    return [
      cornerRect(w, bh, r, r, r, r),
      poly([[w * 0.25, bh - 1], [w * 0.45, bh - 1], [w * 0.2, h]]),
    ].join(' ');
  },
  wedgeEllipseCallout: (w, h) => {
    const bh = h * 0.78;
    return [
      ellipsePath(w / 2, bh / 2, w / 2, bh / 2),
      poly([[w * 0.28, bh * 0.92], [w * 0.46, bh * 0.96], [w * 0.2, h]]),
    ].join(' ');
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
export const OPEN_PRESETS = new Set([
  'line', 'straightConnector1', 'arc',
  'leftBracket', 'rightBracket', 'bracketPair',
  'leftBrace', 'rightBrace', 'bracePair',
]);

/**
 * Presets whose path encloses a hole (outer ring + inner ring). They must fill
 * with the even-odd rule so the inner subpath carves out rather than fills.
 */
export const EVENODD_PRESETS = new Set(['donut', 'noSmoking', 'frame']);

/** True if we have an exact generator for this preset. */
export function hasPreset(preset: string): boolean {
  return preset in GENERATORS;
}

/** Generate an SVG path `d` for a preset shape in a w×h pixel box. */
export function presetPath(preset: string, w: number, h: number, adj: Adjust): string {
  const gen = GENERATORS[preset] ?? rect;
  return gen(w, h, adj);
}
