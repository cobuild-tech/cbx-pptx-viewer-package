/**
 * Picture renderer: {@link PictureShape} -> an `<img>` in a positioned box.
 *
 * A non-rectangular preset geometry clips the image via `clip-path`; a
 * source-rectangle crop is realised by oversizing and offsetting the image
 * inside an `overflow:hidden` box.
 */
import type { PictureShape } from '../model.js';
import { colorToCss } from '../color.js';
import { positioned, type RenderDeps } from '../render/primitives.js';
import { presetPath } from '../shapes/geometry/presets.js';
import { applyEffects } from '../effects/render.js';

export function renderPicture(shape: PictureShape, deps: RenderDeps): HTMLElement | null {
  const url = deps.imageUrl(shape.part);
  if (!url) return null;
  const el = positioned(shape.transform);
  el.style.overflow = 'hidden';

  // Clip the image to a non-rectangular preset (e.g. cropped into a circle).
  if (shape.geom?.type === 'preset' && shape.geom.preset !== 'rect') {
    const w = shape.transform?.w ?? 0;
    const h = shape.transform?.h ?? 0;
    el.style.clipPath = `path('${presetPath(shape.geom.preset, w, h, shape.geom.adjust)}')`;
  }

  const img = document.createElement('img');
  img.src = url;
  img.alt = '';
  img.draggable = false;
  img.style.position = 'absolute';
  img.style.display = 'block';

  const crop = shape.crop;
  if (crop && (crop.l || crop.t || crop.r || crop.b)) {
    const visW = Math.max(0.001, 1 - crop.l - crop.r);
    const visH = Math.max(0.001, 1 - crop.t - crop.b);
    img.style.width = `${100 / visW}%`;
    img.style.height = `${100 / visH}%`;
    img.style.left = `${(-crop.l / visW) * 100}%`;
    img.style.top = `${(-crop.t / visH) * 100}%`;
  } else {
    img.style.width = '100%';
    img.style.height = '100%';
    img.style.objectFit = 'fill';
  }
  el.appendChild(img);
  if (shape.stroke) {
    el.style.outline = `${shape.stroke.width}px solid ${colorToCss(shape.stroke.color)}`;
    el.style.outlineOffset = `-${shape.stroke.width}px`;
  }
  applyEffects(el, shape.effects);
  return el;
}
