/**
 * Shape renderers: preset shapes, connectors, and groups.
 *
 * Fills use native CSS via {@link applyFillBackground} with a `clip-path` for
 * non-rectangular preset geometry; outlines and custom geometry are drawn as an
 * SVG path overlay so stroke width hugs the shape edge. `renderGroup` recurses
 * through the top-level {@link renderShape} dispatcher (in render/dom.ts) so a
 * group can contain any shape kind.
 */
import type {
  PresetShape,
  ConnectorShape,
  GroupShape,
  Fill,
  Stroke,
  LineEnd,
  LineEndType,
} from '../model.js';
import { colorToCss } from '../color.js';
import { presetPath, OPEN_PRESETS, EVENODD_PRESETS } from './geometry/presets.js';
import { SVG_NS, positioned, applyFillBackground, type RenderDeps } from '../render/primitives.js';
import { renderTextBody } from '../text/render.js';
import { applyEffects } from '../effects/render.js';
import { renderShape } from '../render/dom.js';

export function renderPreset(shape: PresetShape, deps: RenderDeps): HTMLElement {
  const el = positioned(shape.transform);
  const w = shape.transform?.w ?? 0;
  const h = shape.transform?.h ?? 0;

  if (shape.geom.type === 'preset') {
    const open = OPEN_PRESETS.has(shape.geom.preset);
    const d = presetPath(shape.geom.preset, w, h, shape.geom.adjust);
    if (!open) {
      const fillLayer = document.createElement('div');
      fillLayer.style.position = 'absolute';
      fillLayer.style.inset = '0';
      applyFillBackground(fillLayer, shape.fill, deps);
      if (shape.geom.preset !== 'rect') {
        fillLayer.style.clipPath = EVENODD_PRESETS.has(shape.geom.preset)
          ? `path('evenodd', '${d}')`
          : `path('${d}')`;
      }
      el.appendChild(fillLayer);
    }
    if (shape.stroke) {
      el.appendChild(strokeOverlay(d, w, h, shape.stroke, false, EVENODD_PRESETS.has(shape.geom.preset)));
    }
  } else {
    // Custom geometry: render each subpath in a scaled SVG.
    for (const p of shape.geom.paths) {
      el.appendChild(customGeomSvg(p.d, p.w, p.h, w, h, shape.fill, shape.stroke, deps));
    }
  }

  if (shape.text) {
    if (shape.textBox && shape.transform) {
      // Place the label in its own rectangle (SmartArt txXfrm) instead of
      // filling the shape — the text body fills this wrapper minus its insets.
      const wrap = document.createElement('div');
      wrap.style.position = 'absolute';
      wrap.style.left = `${shape.textBox.x - shape.transform.x}px`;
      wrap.style.top = `${shape.textBox.y - shape.transform.y}px`;
      wrap.style.width = `${shape.textBox.w}px`;
      wrap.style.height = `${shape.textBox.h}px`;
      wrap.appendChild(renderTextBody(shape.text, deps));
      el.appendChild(wrap);
    } else {
      el.appendChild(renderTextBody(shape.text, deps));
    }
  }
  applyEffects(el, shape.effects);
  return el;
}

export function renderConnector(shape: ConnectorShape, deps: RenderDeps): HTMLElement {
  const el = positioned(shape.transform);
  const w = shape.transform?.w ?? 0;
  const h = shape.transform?.h ?? 0;
  applyEffects(el, shape.effects);
  if (!shape.stroke) return el;
  if (shape.geom.type === 'preset') {
    // Preset connector geometry is already in the shape's box (px) space.
    const d = presetPath(shape.geom.preset, w, h, shape.geom.adjust);
    el.appendChild(strokeOverlay(d, w, h, shape.stroke, false));
  } else {
    // Custom geometry lives in the path's own EMU coordinate space; it must be
    // scaled into the box via a viewBox (as for shapes), otherwise the raw
    // path coordinates draw enormous strokes across the slide.
    for (const p of shape.geom.paths) {
      el.appendChild(customGeomSvg(p.d, p.w, p.h, w, h, undefined, shape.stroke, deps));
    }
  }
  return el;
}

export function renderGroup(shape: GroupShape, deps: RenderDeps): HTMLElement {
  const el = positioned(shape.transform);
  const inner = document.createElement('div');
  inner.style.position = 'absolute';
  inner.style.left = '0';
  inner.style.top = '0';
  inner.style.transformOrigin = '0 0';

  // Map the child coordinate space (chOff/chExt) onto the group box (off/ext).
  // Guard against a missing/degenerate child extent producing an enormous scale.
  const co = shape.childOffset;
  const sx = co.w > 1 && shape.transform ? shape.transform.w / co.w : 1;
  const sy = co.h > 1 && shape.transform ? shape.transform.h / co.h : 1;
  inner.style.transform = `scale(${sx}, ${sy}) translate(${-co.x}px, ${-co.y}px)`;

  for (const childShape of shape.children) {
    const childEl = renderShape(childShape, deps);
    if (childEl) inner.appendChild(childEl);
  }
  el.appendChild(inner);
  return el;
}

function strokeOverlay(
  d: string,
  w: number,
  h: number,
  stroke: Stroke,
  fill: boolean,
  evenodd = false,
): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, 'svg');
  // A straight horizontal/vertical line has a zero-area box (e.g. 94×0). A
  // zero-dimension <svg> isn't painted by the browser even with overflow
  // visible, so the line vanishes — clamp the viewport to a minimum. The path
  // coordinates are unchanged, so the stroke stays centered on the box edge.
  svg.setAttribute('width', `${Math.max(w, 1)}`);
  svg.setAttribute('height', `${Math.max(h, 1)}`);
  svg.style.position = 'absolute';
  svg.style.inset = '0';
  svg.style.overflow = 'visible';
  const path = document.createElementNS(SVG_NS, 'path');
  path.setAttribute('d', d);
  if (evenodd) path.setAttribute('fill-rule', 'evenodd');
  path.setAttribute('fill', fill ? colorToCss(stroke.color) : 'none');
  if (!fill) {
    path.setAttribute('stroke', colorToCss(stroke.color));
    path.setAttribute('stroke-width', `${stroke.width}`);
    if (stroke.dash) path.setAttribute('stroke-dasharray', stroke.dash.join(','));
    if (stroke.cap) path.setAttribute('stroke-linecap', stroke.cap);
    applyLineEnds(svg, path, stroke);
  }
  svg.appendChild(path);
  return svg;
}

const MARKER_SIZE: Record<'sm' | 'med' | 'lg', number> = { sm: 3, med: 4.5, lg: 6 };

/** Arrowhead path in a 0..10 box, tip at (10,5) pointing along +x. */
function markerShape(type: LineEndType): string {
  switch (type) {
    case 'oval':
      return 'M0,5 a5,5 0 1,0 10,0 a5,5 0 1,0 -10,0 Z';
    case 'diamond':
      return 'M5,0 L10,5 L5,10 L0,5 Z';
    default: // triangle / arrow / stealth -> filled triangle
      return 'M0,0 L10,5 L0,10 Z';
  }
}

/** Attach SVG arrowhead markers for a stroke's head/tail ends, if any. */
function applyLineEnds(svg: SVGSVGElement, path: SVGPathElement, stroke: Stroke): void {
  if (!stroke.headEnd && !stroke.tailEnd) return;
  let defs = svg.querySelector('defs');
  if (!defs) {
    defs = document.createElementNS(SVG_NS, 'defs');
    svg.insertBefore(defs, svg.firstChild);
  }
  const color = colorToCss(stroke.color);
  if (stroke.headEnd) {
    const id = `le-${Math.random().toString(36).slice(2, 9)}`;
    defs.appendChild(makeMarker(id, stroke.headEnd, color));
    path.setAttribute('marker-start', `url(#${id})`);
  }
  if (stroke.tailEnd) {
    const id = `le-${Math.random().toString(36).slice(2, 9)}`;
    defs.appendChild(makeMarker(id, stroke.tailEnd, color));
    path.setAttribute('marker-end', `url(#${id})`);
  }
}

function makeMarker(id: string, end: LineEnd, color: string): SVGMarkerElement {
  const m = document.createElementNS(SVG_NS, 'marker');
  m.setAttribute('id', id);
  m.setAttribute('viewBox', '0 0 10 10');
  m.setAttribute('refX', '10'); // tip sits at the line's endpoint
  m.setAttribute('refY', '5');
  m.setAttribute('markerUnits', 'strokeWidth');
  m.setAttribute('markerWidth', `${MARKER_SIZE[end.len]}`);
  m.setAttribute('markerHeight', `${MARKER_SIZE[end.w]}`);
  m.setAttribute('orient', 'auto-start-reverse');
  const p = document.createElementNS(SVG_NS, 'path');
  p.setAttribute('d', markerShape(end.type));
  p.setAttribute('fill', color);
  m.appendChild(p);
  return m;
}

function setupSvgFill(svg: SVGSVGElement, fill: Fill | undefined, deps: RenderDeps): string {
  if (!fill || fill.type === 'none') return 'none';
  if (fill.type === 'solid') return colorToCss(fill.color);

  let defs = svg.querySelector('defs');
  if (!defs) {
    defs = document.createElementNS(SVG_NS, 'defs');
    svg.insertBefore(defs, svg.firstChild);
  }

  const fillId = `fill-${Math.random().toString(36).slice(2, 9)}`;

  if (fill.type === 'gradient') {
    const grad = document.createElementNS(
      SVG_NS,
      fill.radial ? 'radialGradient' : 'linearGradient',
    );
    grad.setAttribute('id', fillId);

    if (!fill.radial) {
      const angleRad = ((fill.angle ?? 0) * Math.PI) / 180;
      const dx = Math.sin(angleRad);
      const dy = -Math.cos(angleRad);
      const x1 = 0.5 - dx * 0.5;
      const y1 = 0.5 - dy * 0.5;
      const x2 = 0.5 + dx * 0.5;
      const y2 = 0.5 + dy * 0.5;
      grad.setAttribute('x1', `${(x1 * 100).toFixed(1)}%`);
      grad.setAttribute('y1', `${(y1 * 100).toFixed(1)}%`);
      grad.setAttribute('x2', `${(x2 * 100).toFixed(1)}%`);
      grad.setAttribute('y2', `${(y2 * 100).toFixed(1)}%`);
    }

    for (const stop of fill.stops) {
      const s = document.createElementNS(SVG_NS, 'stop');
      s.setAttribute('offset', `${(stop.pos * 100).toFixed(1)}%`);
      s.setAttribute('stop-color', colorToCss(stop.color));
      grad.appendChild(s);
    }
    defs.appendChild(grad);
    return `url(#${fillId})`;
  }

  if (fill.type === 'image') {
    const url = deps.imageUrl(fill.part);
    if (!url) return 'none';

    const pattern = document.createElementNS(SVG_NS, 'pattern');
    pattern.setAttribute('id', fillId);
    pattern.setAttribute('patternUnits', 'objectBoundingBox');
    pattern.setAttribute('width', '1');
    pattern.setAttribute('height', '1');

    const img = document.createElementNS(SVG_NS, 'image');
    img.setAttributeNS('http://www.w3.org/1999/xlink', 'href', url);
    img.setAttribute('width', '100%');
    img.setAttribute('height', '100%');
    img.setAttribute('preserveAspectRatio', 'none');
    pattern.appendChild(img);
    defs.appendChild(pattern);
    return `url(#${fillId})`;
  }

  return 'none';
}

function customGeomSvg(
  d: string,
  pw: number,
  ph: number,
  w: number,
  h: number,
  fill: Fill | undefined,
  stroke: Stroke | undefined,
  deps: RenderDeps,
): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, 'svg');
  // Clamp the viewport so a zero-area box (a straight custom line) still paints.
  svg.setAttribute('width', `${Math.max(w, 1)}`);
  svg.setAttribute('height', `${Math.max(h, 1)}`);
  svg.setAttribute('viewBox', `0 0 ${pw || w} ${ph || h}`);
  svg.setAttribute('preserveAspectRatio', 'none');
  svg.style.position = 'absolute';
  svg.style.inset = '0';
  // Geometry must stay within the shape box; clipping prevents an imperfect
  // path (e.g. a spiral with an off arc) from drawing lines across the slide.
  svg.style.overflow = 'hidden';

  const fillVal = setupSvgFill(svg, fill, deps);

  const path = document.createElementNS(SVG_NS, 'path');
  path.setAttribute('d', d);
  path.setAttribute('fill', fillVal);
  if (stroke) {
    path.setAttribute('stroke', colorToCss(stroke.color));
    path.setAttribute('stroke-width', `${stroke.width}`);
    path.setAttribute('vector-effect', 'non-scaling-stroke');
    if (stroke.dash) path.setAttribute('stroke-dasharray', stroke.dash.join(','));
    applyLineEnds(svg, path, stroke);
  }
  svg.appendChild(path);
  return svg;
}
