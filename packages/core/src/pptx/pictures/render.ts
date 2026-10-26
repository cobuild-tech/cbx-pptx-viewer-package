/**
 * Picture renderer: {@link PictureShape} -> an `<img>` in a positioned box.
 *
 * The image sits on its own layer, clipped to the picture's geometry — a
 * `clip-path` for a preset, an SVG mask for custom geometry (whose path lives
 * in its own coordinate space, which `path()` cannot scale). The outline is
 * drawn over that layer along the same path, so the clip cannot cut it off and
 * it straddles the edge as PowerPoint draws it. A source-rectangle crop is
 * realised by oversizing and offsetting the image inside the clipped layer.
 */
import type { CustomPath, PictureShape } from '../model.js';
import { positioned, type RenderDeps } from '../render/primitives.js';
import { presetPath } from '../shapes/geometry/presets.js';
import { strokeOverlay, customGeomSvg } from '../shapes/render.js';
import { applyEffects } from '../effects/render.js';
import { blipFilter } from '../../oxml/drawingml/blip.js';

export function renderPicture(shape: PictureShape, deps: RenderDeps): HTMLElement | null {
  const url = deps.imageUrl(shape.part);
  if (!url) return null;
  const el = positioned(shape.transform);
  const w = shape.transform?.w ?? 0;
  const h = shape.transform?.h ?? 0;
  const geom = shape.geom;
  const custom = geom?.type === 'custom' && geom.paths.length ? geom.paths : undefined;

  const layer = document.createElement('div');
  layer.style.position = 'absolute';
  layer.style.inset = '0';
  layer.style.overflow = 'hidden';
  if (geom?.type === 'preset' && geom.preset !== 'rect') {
    layer.style.clipPath = `path('${presetPath(geom.preset, w, h, geom.adjust)}')`;
  } else if (custom) {
    const mask = geometryMask(custom, w, h);
    layer.style.maskImage = mask;
    layer.style.setProperty('-webkit-mask-image', mask);
    layer.style.maskSize = '100% 100%';
    layer.style.setProperty('-webkit-mask-size', '100% 100%');
    layer.style.maskRepeat = 'no-repeat';
    layer.style.setProperty('-webkit-mask-repeat', 'no-repeat');
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
  // Transparency and recolouring belong to the image, not its outline.
  if (shape.blip?.alpha !== undefined) layer.style.opacity = `${shape.blip.alpha}`;
  const filter = shape.blip ? blipFilter(document, shape.blip) : undefined;
  if (filter) img.style.filter = `url(#${filter})`;
  layer.appendChild(img);
  el.appendChild(layer);

  if (shape.stroke) {
    if (custom) {
      for (const p of custom) el.appendChild(customGeomSvg(p.d, p.w, p.h, w, h, undefined, shape.stroke, deps));
    } else {
      const preset = geom?.type === 'preset' ? geom : undefined;
      const d = presetPath(preset?.preset ?? 'rect', w, h, preset?.adjust ?? {});
      el.appendChild(strokeOverlay(d, w, h, shape.stroke, false));
    }
  }
  applyEffects(el, shape.effects);
  return el;
}

/** A CSS mask image that is opaque inside the custom geometry's paths. */
function geometryMask(paths: CustomPath[], w: number, h: number): string {
  const inner = paths
    .map(
      (p) =>
        `<svg viewBox="0 0 ${p.w || w} ${p.h || h}" width="100%" height="100%" preserveAspectRatio="none">` +
        `<path d="${p.d}"/></svg>`,
    )
    .join('');
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">${inner}</svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`;
}
