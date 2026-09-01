/**
 * Box math for direct manipulation — moving, resizing and rotating a shape.
 *
 * Kept pure (no DOM, no XML) because the hard part is geometry, not plumbing:
 * a shape's `<a:off>`/`<a:ext>` describe its *unrotated* box, while the user
 * drags it in the rotated, possibly mirrored, space it is drawn in. Resizing a
 * rotated shape therefore cannot just add the drag delta to a corner — doing so
 * makes the opposite corner wander, which is the classic bug in this feature.
 * Everything here works in the shape's local space and then puts the box back
 * so that the anchor (the corner or edge opposite the one being dragged) stays
 * exactly where it was on screen.
 *
 * All lengths are CSS px in the slide's base coordinate space, matching the
 * model; conversion to EMU happens at write-back.
 */
import type { Transform } from '../model.js';
import type { ShapeFrame } from '../render/primitives.js';

/** The eight resize handles, named by compass point as they appear on screen. */
export type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';

export interface ResizeOptions {
  /** Shift: keep the aspect ratio (corner handles only, as in PowerPoint). */
  aspect?: boolean;
  /** Alt: resize about the centre instead of the opposite corner. */
  fromCenter?: boolean;
  /** Smallest permitted width/height in px. Default 1. */
  min?: number;
}

const DEG = Math.PI / 180;

/** Move a box by a slide-space delta. Rotation is about the centre, so it does
 * not enter into this: shifting the centre shifts the offset by the same amount. */
export function moveBox(box: Transform, dx: number, dy: number): Transform {
  return { ...box, x: box.x + dx, y: box.y + dy };
}

/**
 * Resize `box` by dragging `handle` through a slide-space delta.
 *
 * The delta is taken into the shape's local space (undoing rotation, then the
 * mirroring that `flipH`/`flipV` apply), the moving edges are displaced there,
 * and the box is recentred so the anchor stays put in slide space.
 */
export function resizeBox(
  box: Transform,
  handle: Handle,
  dx: number,
  dy: number,
  options: ResizeOptions = {},
): Transform {
  const min = options.min ?? 1;
  const rot = box.rot ?? 0;

  // Into the shape's own space: undo the rotation, then the mirror. A mirrored
  // shape draws its local west edge on the screen's east side, so `flipH`/
  // `flipV` move both the grabbed handle *and* the drag direction. The two
  // cancel for the size — the box grows the way the user drags it either way —
  // but not for the anchor, which is the whole point: dragging the handle you
  // can see must pin the corner you can see opposite it.
  const [rx, ry] = rotate(dx, dy, -rot);
  const lx = box.flipH ? -rx : rx;
  const ly = box.flipV ? -ry : ry;
  const h = localHandle(handle, box);

  const west = h.includes('w');
  const east = h.includes('e');
  const north = h.startsWith('n');
  const south = h.startsWith('s');

  // Edge displacements in local space. Alt drags both edges symmetrically.
  let dl = west ? lx : options.fromCenter && east ? -lx : 0;
  let dr = east ? lx : options.fromCenter && west ? -lx : 0;
  let dt = north ? ly : options.fromCenter && south ? -ly : 0;
  let db = south ? ly : options.fromCenter && north ? -ly : 0;

  let w = box.w + dr - dl;
  let hgt = box.h + db - dt;

  // Shift on a corner keeps the ratio; the axis that grew more wins, which is
  // how PowerPoint behaves when the pointer leaves the diagonal.
  if (options.aspect && (west || east) && (north || south) && box.w > 0 && box.h > 0) {
    const s = Math.max(w / box.w, hgt / box.h);
    w = box.w * s;
    hgt = box.h * s;
  }

  w = Math.max(min, w);
  hgt = Math.max(min, hgt);

  // Where the anchor sits relative to the old centre, and where it must sit
  // relative to the new one. Dragging one edge anchors the opposite edge;
  // dragging from the centre anchors the centre itself.
  const ax = options.fromCenter ? 0 : west ? box.w / 2 : east ? -box.w / 2 : 0;
  const ay = options.fromCenter ? 0 : north ? box.h / 2 : south ? -box.h / 2 : 0;
  const ax2 = Math.sign(ax) * (w / 2);
  const ay2 = Math.sign(ay) * (hgt / 2);

  // Back out to slide space — mirror, then rotate, the inverse of the way in —
  // and turn the new centre into a top-left offset.
  const mx = box.flipH ? -(ax - ax2) : ax - ax2;
  const my = box.flipV ? -(ay - ay2) : ay - ay2;
  const [ox, oy] = rotate(mx, my, rot);
  const cx = box.x + box.w / 2 + ox;
  const cy = box.y + box.h / 2 + oy;

  return { ...box, x: cx - w / 2, y: cy - hgt / 2, w, h: hgt };
}

/**
 * Rotate `box` so its top edge faces the pointer, given the pointer's position
 * in slide space. `snap` (shift) quantises to 15°, as PowerPoint does.
 */
export function rotateBox(box: Transform, px: number, py: number, snap = false): Transform {
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  // atan2 measured from straight up, clockwise — the same sense as `rot`.
  let deg = Math.atan2(px - cx, cy - py) / DEG;
  if (snap) deg = Math.round(deg / 15) * 15;
  return { ...box, rot: normalizeAngle(deg) };
}

/** Fold an angle into [0, 360), rounded to a thousandth of a degree so that
 * repeated rotations cannot accumulate float noise in the XML. */
export function normalizeAngle(deg: number): number {
  const r = Math.round(deg * 1000) / 1000;
  return ((r % 360) + 360) % 360;
}

/** Axis-aligned bounds of a box once rotated — what the selection rect needs. */
export function boundsOf(box: Transform): { x: number; y: number; w: number; h: number } {
  const rot = box.rot ?? 0;
  if (!rot) return { x: box.x, y: box.y, w: box.w, h: box.h };
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  for (const [sx, sy] of [[-1, -1], [1, -1], [1, 1], [-1, 1]] as const) {
    const [px, py] = rotate((sx * box.w) / 2, (sy * box.h) / 2, rot);
    minX = Math.min(minX, cx + px);
    maxX = Math.max(maxX, cx + px);
    minY = Math.min(minY, cy + py);
    maxY = Math.max(maxY, cy + py);
  }
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
}

/**
 * A local box in slide space. A shape inside a group states its box in the
 * group's child space, but the selection overlay draws on the slide.
 */
export function frameRect(
  rect: { x: number; y: number; w: number; h: number },
  frame: ShapeFrame,
): { x: number; y: number; w: number; h: number } {
  return {
    x: frame.ox + frame.sx * rect.x,
    y: frame.oy + frame.sy * rect.y,
    w: rect.w * frame.sx,
    h: rect.h * frame.sy,
  };
}

/** A slide-space point in the shape's own coordinate space. */
export function unframePoint(
  frame: ShapeFrame,
  x: number,
  y: number,
): { x: number; y: number } {
  return { x: (x - frame.ox) / (frame.sx || 1), y: (y - frame.oy) / (frame.sy || 1) };
}

/** A slide-space drag delta in the shape's own coordinate space. */
export function unframeDelta(
  frame: ShapeFrame,
  dx: number,
  dy: number,
): { x: number; y: number } {
  return { x: dx / (frame.sx || 1), y: dy / (frame.sy || 1) };
}

/** Rotate a vector clockwise by `deg` (CSS/DrawingML sense: y grows downward). */
function rotate(x: number, y: number, deg: number): [number, number] {
  if (!deg) return [x, y];
  const a = deg * DEG;
  const cos = Math.cos(a);
  const sin = Math.sin(a);
  return [x * cos - y * sin, x * sin + y * cos];
}

/** The handle in the shape's own space: a mirrored shape swaps its edges. */
function localHandle(handle: Handle, box: Transform): string {
  let h: string = handle;
  if (box.flipH) h = h.replace('e', 'E').replace('w', 'e').replace('E', 'w');
  if (box.flipV) h = h.replace('n', 'N').replace('s', 'n').replace('N', 's');
  return h;
}
