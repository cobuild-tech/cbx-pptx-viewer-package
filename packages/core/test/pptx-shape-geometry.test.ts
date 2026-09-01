/**
 * Direct-manipulation box math. The invariant that matters for every resize is
 * that the corner *opposite* the dragged handle does not move on screen —
 * including when the shape is rotated or mirrored, which is where naive
 * implementations drift.
 */
import { describe, it, expect } from 'vitest';
import {
  moveBox,
  resizeBox,
  rotateBox,
  boundsOf,
  normalizeAngle,
  type Handle,
} from '../src/pptx/edit/geometry.js';
import type { Transform } from '../src/pptx/model.js';

/**
 * Where a compass corner of a box lands on screen — the frame of reference the
 * user is actually dragging in.
 *
 * Mirroring does not enter into it: a flip permutes which stored corner shows
 * up here, but the four screen positions are the same set either way.
 */
function screenCorner(b: Transform, dir: string): [number, number] {
  const rot = ((b.rot ?? 0) * Math.PI) / 180;
  const lx = ((dir.includes('w') ? -1 : 1) * b.w) / 2;
  const ly = ((dir.startsWith('n') ? -1 : 1) * b.h) / 2;
  const cx = b.x + b.w / 2;
  const cy = b.y + b.h / 2;
  return [cx + lx * Math.cos(rot) - ly * Math.sin(rot), cy + lx * Math.sin(rot) + ly * Math.cos(rot)];
}

/** The corner a handle anchors against: the diagonally opposite one. */
const OPPOSITE: Record<string, string> = { nw: 'se', ne: 'sw', se: 'nw', sw: 'ne' };

function expectAnchorFixed(box: Transform, handle: Handle, dx: number, dy: number) {
  const out = resizeBox(box, handle, dx, dy);
  const dir = OPPOSITE[handle]!;
  const before = screenCorner(box, dir);
  const after = screenCorner(out, dir);
  expect(after[0]).toBeCloseTo(before[0], 6);
  expect(after[1]).toBeCloseTo(before[1], 6);
  return out;
}

describe('moveBox', () => {
  it('offsets the box and leaves size and rotation alone', () => {
    const out = moveBox({ x: 10, y: 20, w: 100, h: 50, rot: 30 }, 5, -7);
    expect(out).toEqual({ x: 15, y: 13, w: 100, h: 50, rot: 30 });
  });
});

describe('resizeBox', () => {
  it('grows an unrotated box from the south-east handle', () => {
    const out = resizeBox({ x: 0, y: 0, w: 100, h: 50 }, 'se', 10, 5);
    expect(out).toMatchObject({ x: 0, y: 0, w: 110, h: 55 });
  });

  it('moves the offset when dragging a north-west handle', () => {
    const out = resizeBox({ x: 100, y: 100, w: 100, h: 50 }, 'nw', 10, 10);
    expect(out).toMatchObject({ x: 110, y: 110, w: 90, h: 40 });
  });

  it('changes only one axis for an edge handle', () => {
    const out = resizeBox({ x: 0, y: 0, w: 100, h: 50 }, 'e', 10, 999);
    expect(out).toMatchObject({ x: 0, y: 0, w: 110, h: 50 });
  });

  it('keeps the opposite corner fixed at every rotation', () => {
    for (const rot of [0, 15, 45, 90, 137.5, 180, 270, 359]) {
      for (const handle of ['nw', 'ne', 'se', 'sw'] as Handle[]) {
        expectAnchorFixed({ x: 40, y: 60, w: 120, h: 80, rot }, handle, 17, -9);
      }
    }
  });

  it('keeps the opposite corner fixed when the shape is mirrored', () => {
    for (const flips of [{ flipH: true }, { flipV: true }, { flipH: true, flipV: true }]) {
      for (const handle of ['nw', 'se'] as Handle[]) {
        expectAnchorFixed({ x: 10, y: 10, w: 100, h: 60, rot: 33, ...flips }, handle, 12, 8);
      }
    }
  });

  it('grows a mirrored shape in the direction it is dragged', () => {
    // Under flipH the handle drawn on the east is the box's own west edge, so
    // the pull is stored as that edge moving out — the box still grows to 110
    // and still keeps its screen-west side at x=0, which is what the user sees.
    const out = resizeBox({ x: 0, y: 0, w: 100, h: 50, flipH: true }, 'e', 10, 0);
    expect(out.w).toBeCloseTo(110);
    expect(out.x).toBeCloseTo(0);
  });

  it('preserves the aspect ratio on a corner when asked', () => {
    const out = resizeBox({ x: 0, y: 0, w: 100, h: 50 }, 'se', 40, 0, { aspect: true });
    expect(out.w / out.h).toBeCloseTo(2);
    expect(out.w).toBeCloseTo(140);
  });

  it('resizes about the centre with fromCenter', () => {
    const out = resizeBox({ x: 0, y: 0, w: 100, h: 50 }, 'e', 10, 0, { fromCenter: true });
    expect(out).toMatchObject({ x: -10, y: 0, w: 120, h: 50 });
  });

  it('clamps to the minimum size instead of inverting the box', () => {
    const out = resizeBox({ x: 0, y: 0, w: 100, h: 50 }, 'e', -500, 0, { min: 2 });
    expect(out.w).toBe(2);
    expect(out.x).toBe(0);
    expect(out.h).toBe(50);
  });
});

describe('rotateBox', () => {
  it('points the top edge at the pointer', () => {
    const box = { x: 0, y: 0, w: 100, h: 100 };
    expect(rotateBox(box, 150, 50).rot).toBeCloseTo(90); // pointer due east
    expect(rotateBox(box, 50, 150).rot).toBeCloseTo(180); // due south
    expect(rotateBox(box, -50, 50).rot).toBeCloseTo(270); // due west
  });

  it('snaps to 15 degree steps under shift', () => {
    const box = { x: 0, y: 0, w: 100, h: 100 };
    expect(rotateBox(box, 150, 47, true).rot).toBe(90);
  });

  it('normalises angles into [0, 360)', () => {
    expect(normalizeAngle(-90)).toBe(270);
    expect(normalizeAngle(360)).toBe(0);
    expect(normalizeAngle(725.0001)).toBe(5);
  });
});

describe('boundsOf', () => {
  it('returns the box itself when unrotated', () => {
    expect(boundsOf({ x: 5, y: 6, w: 10, h: 20 })).toEqual({ x: 5, y: 6, w: 10, h: 20 });
  });

  it('grows to the rotated extent', () => {
    const b = boundsOf({ x: 0, y: 0, w: 100, h: 0, rot: 90 });
    expect(b.w).toBeCloseTo(0);
    expect(b.h).toBeCloseTo(100);
  });
});
