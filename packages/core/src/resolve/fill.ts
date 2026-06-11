/**
 * Fill and stroke parsing: solid / gradient / image / pattern fills and outlines.
 * Colors are resolved through {@link ColorContext}; image fills resolve their
 * `r:embed` to a media part path via the current scope.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../xml.js';
import { emuToPx } from '../units.js';
import type { Fill, Stroke, GradientStop } from '../model.js';
import {
  resolveContainerColor,
  resolveColorEl,
  findColorEl,
  type ColorContext,
} from './color.js';

export interface ParseScope {
  colorCtx: ColorContext;
  /** Resolve an `r:embed`/`r:link` id to a media part path. */
  resolveImage(relId: string): string | undefined;
  /** Resolve an `r:id` on a hyperlink to its (usually external) target URL. */
  resolveHyperlink?(relId: string): string | undefined;
}

const FILL_TAGS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill'];

/** First fill element directly under a container (spPr, bgPr, tcPr, etc.). */
function findFillEl(container: XmlNode | undefined): XmlNode | undefined {
  if (!container) return undefined;
  return container.children.find((c) => FILL_TAGS.includes(localName(c.name)));
}

/**
 * Parse the fill declared directly on a container. Returns `undefined` when no
 * fill is present (the caller should inherit), distinct from `{type:'none'}`.
 */
export function parseFill(container: XmlNode | undefined, scope: ParseScope): Fill | undefined {
  const el = findFillEl(container);
  if (!el) return undefined;
  return fillFromEl(el, scope);
}

export function fillFromEl(el: XmlNode, scope: ParseScope): Fill {
  switch (localName(el.name)) {
    case 'noFill':
      return { type: 'none' };
    case 'solidFill': {
      const color = resolveContainerColor(el, scope.colorCtx);
      return color ? { type: 'solid', color } : { type: 'none' };
    }
    case 'gradFill':
      return parseGradient(el, scope.colorCtx);
    case 'blipFill':
      return parseBlip(el, scope);
    case 'pattFill': {
      // Approximate a pattern with its foreground color.
      const fg = resolveContainerColor(child(el, 'fgClr'), scope.colorCtx);
      return fg ? { type: 'solid', color: fg } : { type: 'none' };
    }
    default:
      return { type: 'none' };
  }
}

function parseGradient(el: XmlNode, ctx: ColorContext): Fill {
  const stops: GradientStop[] = [];
  for (const gs of children(child(el, 'gsLst'), 'gs')) {
    const pos = (attrNum(gs, 'pos') ?? 0) / 100000;
    const color = resolveColorEl(findColorEl(gs), ctx);
    if (color) stops.push({ pos, color });
  }
  if (stops.length === 0) return { type: 'none' };

  const lin = child(el, 'lin');
  if (lin) {
    // ang is in 60000ths of a degree, measured clockwise from 3 o'clock.
    // CSS gradient angle is measured clockwise from 12 o'clock.
    const angOoxml = (attrNum(lin, 'ang') ?? 0) / 60000;
    const cssAngle = (angOoxml + 90) % 360;
    return { type: 'gradient', stops, angle: cssAngle, radial: false };
  }
  // path (radial/rectangular) gradients are approximated as radial.
  return { type: 'gradient', stops, radial: true };
}

function parseBlip(el: XmlNode, scope: ParseScope): Fill {
  const blip = child(el, 'blip');
  const rId = attr(blip, 'embed') ?? attr(blip, 'link');
  const part = rId ? scope.resolveImage(rId) : undefined;
  if (!part) return { type: 'none' };

  const fill: Fill = { type: 'image', part };
  const srcRect = child(el, 'srcRect');
  if (srcRect) {
    fill.crop = {
      l: (attrNum(srcRect, 'l') ?? 0) / 100000,
      t: (attrNum(srcRect, 't') ?? 0) / 100000,
      r: (attrNum(srcRect, 'r') ?? 0) / 100000,
      b: (attrNum(srcRect, 'b') ?? 0) / 100000,
    };
  }
  return fill;
}

const DASH_PATTERNS: Record<string, number[]> = {
  dash: [4, 3],
  dashDot: [4, 3, 1, 3],
  lgDash: [8, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [2, 2],
  sysDot: [1, 2],
  sysDashDot: [2, 2, 1, 2],
  sysDashDotDot: [2, 2, 1, 2, 1, 2],
};

/** Parse an `<a:ln>` outline into a {@link Stroke}, or undefined if no line. */
export function parseStroke(spPr: XmlNode | undefined, scope: ParseScope): Stroke | undefined {
  return strokeFromLn(child(spPr, 'ln'), scope);
}

/** Build a {@link Stroke} from an `<a:ln>`-like element (also used for table borders). */
export function strokeFromLn(ln: XmlNode | undefined, scope: ParseScope): Stroke | undefined {
  if (!ln) return undefined;
  // Explicit noFill on the line means "no outline".
  if (child(ln, 'noFill')) return undefined;

  const color = resolveContainerColor(child(ln, 'solidFill'), scope.colorCtx);
  if (!color) return undefined;

  const wEmu = attrNum(ln, 'w') ?? 9525; // default ~0.75pt
  const stroke: Stroke = { color, width: Math.max(0.5, emuToPx(wEmu)) };

  const dashVal = attr(child(ln, 'prstDash'), 'val');
  if (dashVal && DASH_PATTERNS[dashVal]) {
    stroke.dash = DASH_PATTERNS[dashVal].map((n) => n * stroke.width);
  }
  const cap = attr(ln, 'cap');
  if (cap === 'rnd') stroke.cap = 'round';
  else if (cap === 'sq') stroke.cap = 'square';
  else if (cap === 'flat') stroke.cap = 'butt';

  return stroke;
}
