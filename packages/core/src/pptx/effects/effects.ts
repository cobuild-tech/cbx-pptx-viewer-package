/**
 * Effect parsing: `<a:effectLst>` -> {@link Effect}[].
 *
 * Handles the common DrawingML effects (outer/inner shadow, glow, soft edge,
 * reflection). Shadow direction is in 60000ths of a degree clockwise from 3
 * o'clock; we resolve it to px dx/dy offsets in screen space (y grows down).
 * Colors (with their alpha) resolve through the shared {@link ColorContext}.
 */
import { attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { emuToPx } from '../../oxml/units.js';
import type { Effect } from '../model.js';
import { resolveColorEl, findColorEl, type ColorContext } from '../color.js';

/** Parse an `<a:effectLst>` element into effects, in document order. */
export function parseEffectLst(effectLst: XmlNode | undefined, ctx: ColorContext): Effect[] {
  if (!effectLst) return [];
  const out: Effect[] = [];
  for (const el of effectLst.children) {
    const effect = parseEffect(el, ctx);
    if (effect) out.push(effect);
  }
  return out;
}

function parseEffect(el: XmlNode, ctx: ColorContext): Effect | null {
  switch (localName(el.name)) {
    case 'outerShdw':
    case 'innerShdw': {
      const { dx, dy } = shadowOffset(el);
      const color = resolveColorEl(findColorEl(el), ctx) ?? { hex: '000000', alpha: 0.4 };
      const blur = emuToPx(attrNum(el, 'blurRad') ?? 0);
      return localName(el.name) === 'outerShdw'
        ? { type: 'outerShadow', dx, dy, blur, color }
        : { type: 'innerShadow', dx, dy, blur, color };
    }
    case 'glow': {
      const color = resolveColorEl(findColorEl(el), ctx);
      if (!color) return null;
      return { type: 'glow', radius: emuToPx(attrNum(el, 'rad') ?? 0), color };
    }
    case 'softEdge':
      return { type: 'softEdge', radius: emuToPx(attrNum(el, 'rad') ?? 0) };
    case 'reflection':
      return {
        type: 'reflection',
        blur: emuToPx(attrNum(el, 'blurRad') ?? 0),
        dist: emuToPx(attrNum(el, 'dist') ?? 0),
        startAlpha: (attrNum(el, 'stA') ?? 100000) / 100000,
        endAlpha: (attrNum(el, 'endA') ?? 0) / 100000,
      };
    default:
      return null;
  }
}

/** Resolve a shadow's dist+dir to px dx/dy offsets (screen space, y down). */
function shadowOffset(el: XmlNode): { dx: number; dy: number } {
  const dist = emuToPx(attrNum(el, 'dist') ?? 0);
  const dirDeg = (attrNum(el, 'dir') ?? 0) / 60000;
  const rad = (dirDeg * Math.PI) / 180;
  return { dx: dist * Math.cos(rad), dy: dist * Math.sin(rad) };
}

