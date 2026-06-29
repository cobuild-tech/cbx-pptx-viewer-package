import { describe, it, expect } from 'vitest';
import { presetPath, OPEN_PRESETS, EVENODD_PRESETS, hasPreset } from '../src/pptx/shapes/geometry/presets.js';

describe('preset geometry', () => {
  it('renders arc as an open (unclosed) elliptical-arc stroke', () => {
    // adj1/adj2 are angles in 60000ths of a degree: 270deg start, ~33deg end.
    const d = presetPath('arc', 100, 100, { adj1: 270 * 60000 / 100000, adj2: 33 * 60000 / 100000 });
    expect(d).toMatch(/^M[\d.,-]+ A50,50 0 [01] 1 [\d.,-]+$/);
    expect(d).not.toContain('Z'); // open path, never filled into a rectangle
    expect(OPEN_PRESETS.has('arc')).toBe(true);
  });

  it('renders homePlate as a 5-point pentagon with a pointed right edge', () => {
    const d = presetPath('homePlate', 200, 100, { adj: 0.5 });
    // tip at the vertical centre on the right edge
    expect(d).toContain('200,50');
    expect(d.match(/L/g)?.length).toBe(4); // M + 4 L = 5 points
    expect(d.endsWith('Z')).toBe(true);
  });

  it('rounds only the top-right corner for round1Rect', () => {
    expect(hasPreset('round1Rect')).toBe(true);
    // exactly one arc (the single rounded corner)
    const d = presetPath('round1Rect', 200, 100, { adj: 0.2 });
    expect(d.match(/Q/g)?.length).toBe(1);
  });

  it('renders round2SameRect with adj=0 as a plain (square) rectangle', () => {
    const d = presetPath('round2SameRect', 200, 100, { adj1: 0, adj2: 0 });
    expect(d).not.toContain('Q'); // no rounded corners at all
  });

  it('rounds all four corners for roundRect', () => {
    expect(presetPath('roundRect', 100, 100, { adj: 0.2 }).match(/Q/g)?.length).toBe(4);
  });
});

describe('extended preset geometry', () => {
  it('builds an N-pointed star with 2N vertices', () => {
    // star5 = 5 outer + 5 inner points = 10 vertices = M + 9 L.
    const d = presetPath('star5', 100, 100, {});
    expect(d.match(/L/g)?.length).toBe(9);
    expect(d.endsWith('Z')).toBe(true);
  });

  it('renders holed shapes as two subpaths filled even-odd', () => {
    for (const name of ['donut', 'frame', 'noSmoking']) {
      const d = presetPath(name, 100, 100, { adj: 0.2 });
      expect(d.match(/M/g)?.length).toBeGreaterThanOrEqual(2); // outer + inner ring
      expect(EVENODD_PRESETS.has(name)).toBe(true);
    }
  });

  it('treats brackets and braces as open (unfilled) strokes', () => {
    for (const name of ['leftBracket', 'rightBracket', 'leftBrace', 'rightBrace', 'bracketPair', 'bracePair']) {
      expect(OPEN_PRESETS.has(name)).toBe(true);
      expect(presetPath(name, 100, 100, {})).not.toContain('Z'); // never closed into a fill
    }
  });

  it('points a vertical up-arrow tip at the top centre', () => {
    const d = presetPath('upArrow', 100, 200, {});
    expect(d.startsWith('M50,0')).toBe(true); // tip at (w/2, 0)
  });

  it('draws a pie slice from the centre with an arc', () => {
    const d = presetPath('pie', 100, 100, { adj1: 0, adj2: 270 * 0.6 });
    expect(d.startsWith('M50,50')).toBe(true); // starts at centre
    expect(d).toContain('A'); // elliptical arc
  });

  it('falls back to a rectangle for an unknown preset', () => {
    expect(hasPreset('notARealShape')).toBe(false);
    expect(presetPath('notARealShape', 80, 40, {})).toBe('M0,0 L80,0 L80,40 L0,40 Z');
  });
});
