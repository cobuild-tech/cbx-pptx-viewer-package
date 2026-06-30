/**
 * Fill and stroke parsing: solid / gradient / image / pattern fills and outlines.
 * Colors are resolved through {@link ColorContext}; image fills resolve their
 * `r:embed` to a media part path via the current {@link ParseScope}.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { emuToPx } from '../../oxml/units.js';
import type { Fill, Stroke, GradientStop, LineEnd, LineEndType } from '../model.js';
import { resolveContainerColor, resolveColorEl, findColorEl } from '../color.js';
import type { ParseScope } from '../scope.js';

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

function parseGradient(el: XmlNode, ctx: ParseScope['colorCtx']): Fill {
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
  // path (radial/rectangular) gradients are approximated as radial. The
  // first stop sits at the focus point given by <a:fillToRect> insets; its
  // centre is where the gradient radiates from (e.g. l=t=100% => bottom-right).
  const ftr = child(child(el, 'path'), 'fillToRect');
  const inset = (name: string) => (attrNum(ftr, name) ?? 0) / 100000;
  const center = {
    x: (inset('l') + (1 - inset('r'))) / 2,
    y: (inset('t') + (1 - inset('b'))) / 2,
  };
  return { type: 'gradient', stops, radial: true, center };
}

/**
 * Resolve a `<a:blip>` to its media relationship id. Modern Office stores
 * vector images with no primary raster embed — only an `<asvg:svgBlip>`
 * alternative in the blip's extLst — so fall back to that (browsers render SVG).
 */
export function blipEmbed(blip: XmlNode | undefined): string | undefined {
  const direct = attr(blip, 'embed') ?? attr(blip, 'link');
  if (direct) return direct;
  for (const ext of children(child(blip, 'extLst'), 'ext')) {
    const svg = child(ext, 'svgBlip');
    const rId = svg ? attr(svg, 'embed') ?? attr(svg, 'link') : undefined;
    if (rId) return rId;
  }
  return undefined;
}

function parseBlip(el: XmlNode, scope: ParseScope): Fill {
  const blip = child(el, 'blip');
  const rId = blipEmbed(blip);
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

  const headEnd = lineEnd(child(ln, 'headEnd'));
  if (headEnd) stroke.headEnd = headEnd;
  const tailEnd = lineEnd(child(ln, 'tailEnd'));
  if (tailEnd) stroke.tailEnd = tailEnd;

  return stroke;
}

const LINE_END_TYPES = new Set(['triangle', 'arrow', 'stealth', 'diamond', 'oval']);

/** Parse an `<a:headEnd>`/`<a:tailEnd>` arrowhead descriptor, if present. */
function lineEnd(el: XmlNode | undefined): LineEnd | undefined {
  const type = attr(el, 'type');
  if (!type || !LINE_END_TYPES.has(type)) return undefined;
  const size = (v: string | undefined): 'sm' | 'med' | 'lg' =>
    v === 'sm' || v === 'lg' ? v : 'med';
  return { type: type as LineEndType, w: size(attr(el, 'w')), len: size(attr(el, 'len')) };
}
