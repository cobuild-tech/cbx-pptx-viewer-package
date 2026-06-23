import { describe, it, expect } from 'vitest';
import { presetPath, OPEN_PRESETS, hasPreset } from '../src/pptx/shapes/geometry/presets.js';

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

  it('maps rounded-rect variants to a generator (not the rect fallback)', () => {
    expect(hasPreset('round1Rect')).toBe(true);
    expect(hasPreset('round2SameRect')).toBe(true);
    // a rounded rect uses arcs (Q), a plain rect does not
    expect(presetPath('round1Rect', 100, 100, {})).toContain('Q');
  });
});
