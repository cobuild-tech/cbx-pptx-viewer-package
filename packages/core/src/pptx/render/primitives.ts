/**
 * Shared DOM-rendering primitives used across every feature renderer.
 *
 * Keeps the cross-cutting pieces — the render dependency contract, positioning
 * a shape from its transform, and painting a {@link Fill} onto an element's CSS
 * background — in one place so feature slices (shapes/text/tables/pictures/…)
 * don't duplicate them.
 */
import type { Fill, Transform } from '../model.js';
import { colorToCss } from '../color.js';

export const SVG_NS = 'http://www.w3.org/2000/svg';

export interface RenderDeps {
  /** Resolve a media part path to a displayable URL. */
  imageUrl(part: string): string | undefined;
}

/** Position + rotate/flip a shape container per its transform. */
export function positioned(transform: Transform | undefined): HTMLDivElement {
  const el = document.createElement('div');
  el.style.position = 'absolute';
  if (transform) {
    el.style.left = `${transform.x}px`;
    el.style.top = `${transform.y}px`;
    el.style.width = `${transform.w}px`;
    el.style.height = `${transform.h}px`;
    const parts: string[] = [];
    if (transform.rot) parts.push(`rotate(${transform.rot}deg)`);
    if (transform.flipH || transform.flipV) {
      parts.push(`scale(${transform.flipH ? -1 : 1}, ${transform.flipV ? -1 : 1})`);
    }
    if (parts.length) {
      el.style.transform = parts.join(' ');
      el.style.transformOrigin = 'center';
    }
  }
  return el;
}

/** Apply a {@link Fill} to an element's CSS background. */
export function applyFillBackground(el: HTMLElement, fill: Fill, deps: RenderDeps): void {
  switch (fill.type) {
    case 'solid':
      el.style.background = colorToCss(fill.color);
      break;
    case 'gradient': {
      const stops = fill.stops.map((s) => `${colorToCss(s.color)} ${(s.pos * 100).toFixed(1)}%`);
      if (fill.radial) {
        const at = fill.center
          ? ` at ${(fill.center.x * 100).toFixed(1)}% ${(fill.center.y * 100).toFixed(1)}%`
          : '';
        el.style.background = `radial-gradient(farthest-corner${at}, ${stops.join(',')})`;
      } else {
        el.style.background = `linear-gradient(${fill.angle ?? 0}deg, ${stops.join(',')})`;
      }
      break;
    }
    case 'image': {
      const url = deps.imageUrl(fill.part);
      if (url) {
        el.style.backgroundImage = `url("${url}")`;
        el.style.backgroundSize = 'cover';
        el.style.backgroundPosition = 'center';
      }
      break;
    }
    case 'none':
      break;
  }
}
