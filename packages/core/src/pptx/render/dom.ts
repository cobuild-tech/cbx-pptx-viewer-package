/**
 * DOM renderer entry point: {@link Slide} model -> HTML elements.
 *
 * Every slide renders into a div sized to the slide's base pixel space; the
 * viewer scales that div to fit the viewport. `renderShape` is the dispatcher
 * each feature slice renders through — it routes a {@link Shape} to its feature
 * renderer and is also what `renderGroup` recurses into.
 */
import type { Slide, SlideSize, Shape, FrameShape } from '../model.js';
import { applyFillBackground, positioned, type RenderDeps } from './primitives.js';
import { renderPreset, renderConnector, renderGroup } from '../shapes/render.js';
import { renderPicture } from '../pictures/render.js';
import { renderTable } from '../tables/render.js';

export type { RenderDeps } from './primitives.js';

/** Render one slide into a positioned div sized in the slide's base px space. */
export function renderSlide(slide: Slide, size: SlideSize, deps: RenderDeps): HTMLDivElement {
  const root = document.createElement('div');
  root.className = 'pptx-slide';
  root.style.position = 'relative';
  root.style.width = `${size.wPx}px`;
  root.style.height = `${size.hPx}px`;
  root.style.overflow = 'hidden';
  applyFillBackground(root, slide.background, deps);

  for (const shape of slide.shapes) {
    const el = renderShape(shape, deps);
    if (el) root.appendChild(el);
  }
  return root;
}

/** Route a shape to its feature renderer. Also used by group recursion. */
export function renderShape(
  shape: Shape,
  deps: RenderDeps,
  sx = 1,
  sy = 1,
  tx = 0,
  ty = 0,
): HTMLElement | null {
  switch (shape.kind) {
    case 'shape':
      return renderPreset(shape, deps, sx, sy, tx, ty);
    case 'picture':
      return renderPicture(shape, deps, sx, sy, tx, ty);
    case 'group':
      return renderGroup(shape, deps, sx, sy, tx, ty);
    case 'connector':
      return renderConnector(shape, deps, sx, sy, tx, ty);
    case 'frame':
      return renderFrame(shape, deps, sx, sy, tx, ty);
  }
}

function renderFrame(
  shape: FrameShape,
  deps: RenderDeps,
  sx = 1,
  sy = 1,
  tx = 0,
  ty = 0,
): HTMLElement {
  const el = positioned(shape.transform, sx, sy, tx, ty);
  if (shape.frameType === 'table' && shape.table) {
    el.appendChild(renderTable(shape.table, deps, sx, sy));
  } else if (shape.frameType === 'diagram' && shape.diagram?.length) {
    // SmartArt: the pre-laid-out shapes are positioned in the frame's own
    // coordinate space, so they drop straight into the (positioned) frame box.
    for (const s of shape.diagram) {
      const child = renderShape(s, deps, sx, sy, 0, 0);
      if (child) el.appendChild(child);
    }
  } else {
    // Charts / unresolved diagrams / unknown: a labeled placeholder so it isn't invisible.
    el.style.display = 'flex';
    el.style.alignItems = 'center';
    el.style.justifyContent = 'center';
    el.style.border = '1px dashed #bbb';
    el.style.color = '#999';
    el.style.font = '12px system-ui, sans-serif';
    el.textContent = `[${shape.frameType}]`;
  }
  return el;
}
