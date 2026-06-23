/**
 * Shared shape-resolution primitives.
 *
 * These helpers turn the geometry/transform/fill/stroke/text-style of a shape
 * into model values, applying the placeholder inheritance chain
 * (shape -> layout placeholder -> master placeholder). They're the reusable
 * core that every shape kind (preset, picture, connector, group) builds on, so
 * the leaf builders in this slice depend only on `props.ts` — never on each
 * other — keeping the dependency graph acyclic.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { emuToPx, angleToDeg } from '../../oxml/units.js';
import type { Transform, Geometry, Fill, Stroke, Effect } from '../model.js';
import { type ColorContext, type Theme, resolveColorEl, findColorEl } from '../color.js';
import { strokeFromLn } from './fill.js';
import { parseEffectLst } from '../effects/effects.js';
import { parseCustomGeometry } from './geometry/custom.js';
import { TextStyleChain } from '../text/textStyles.js';
import { type PhInfo, masterStyleKey, lstStyleOf } from './placeholders.js';
import type { ParseScope } from '../scope.js';

export interface SlideScopes {
  slide: ParseScope;
  layout: ParseScope;
  master: ParseScope;
}

export interface SlideBuildCtx {
  colorCtx: ColorContext;
  theme: Theme;
  layoutPhs: PhInfo[];
  masterPhs: PhInfo[];
  /** The `<p:txStyles>` element of the slide master, if present. */
  masterTxStyles?: XmlNode;
  scopes: SlideScopes;
}

export interface BuildOpts {
  /** Skip content placeholders (used when compositing layout/master shapes). */
  skipPlaceholders?: boolean;
}

/** True if a shape's `<spPr>` declares `<a:grpFill/>` (inherit the group's fill). */
export function hasGrpFill(spPr: XmlNode | undefined): boolean {
  return !!spPr && spPr.children.some((c) => localName(c.name) === 'grpFill');
}

export function parseXfrmEl(xfrm: XmlNode | undefined): Transform | undefined {
  if (!xfrm) return undefined;
  const off = child(xfrm, 'off');
  const ext = child(xfrm, 'ext');
  if (!off || !ext) return undefined;
  const t: Transform = {
    x: emuToPx(attrNum(off, 'x') ?? 0),
    y: emuToPx(attrNum(off, 'y') ?? 0),
    w: emuToPx(attrNum(ext, 'cx') ?? 0),
    h: emuToPx(attrNum(ext, 'cy') ?? 0),
  };
  const rot = attrNum(xfrm, 'rot');
  if (rot) t.rot = angleToDeg(rot);
  if (attr(xfrm, 'flipH') === '1') t.flipH = true;
  if (attr(xfrm, 'flipV') === '1') t.flipV = true;
  return t;
}

export function parseGeometry(spPr: XmlNode | undefined): Geometry {
  return resolveGeometry(spPr, undefined, undefined);
}

export function resolveTransform(
  spPr: XmlNode | undefined,
  layoutSpPr: XmlNode | undefined,
  masterSpPr: XmlNode | undefined,
): Transform | undefined {
  const xfrm = child(spPr, 'xfrm');
  const lxfrm = child(layoutSpPr, 'xfrm');
  const mxfrm = child(masterSpPr, 'xfrm');

  if (!xfrm && !lxfrm && !mxfrm) return undefined;

  const off = child(xfrm, 'off') ?? child(lxfrm, 'off') ?? child(mxfrm, 'off');
  const ext = child(xfrm, 'ext') ?? child(lxfrm, 'ext') ?? child(mxfrm, 'ext');

  if (!off || !ext) return undefined;

  const t: Transform = {
    x: emuToPx(attrNum(off, 'x') ?? 0),
    y: emuToPx(attrNum(off, 'y') ?? 0),
    w: emuToPx(attrNum(ext, 'cx') ?? 0),
    h: emuToPx(attrNum(ext, 'cy') ?? 0),
  };

  const rot = attrNum(xfrm, 'rot') ?? attrNum(lxfrm, 'rot') ?? attrNum(mxfrm, 'rot');
  if (rot) t.rot = angleToDeg(rot);

  const flipH = attr(xfrm, 'flipH') ?? attr(lxfrm, 'flipH') ?? attr(mxfrm, 'flipH');
  if (flipH === '1' || flipH === 'true') t.flipH = true;

  const flipV = attr(xfrm, 'flipV') ?? attr(lxfrm, 'flipV') ?? attr(mxfrm, 'flipV');
  if (flipV === '1' || flipV === 'true') t.flipV = true;

  return t;
}

export function resolveGeometry(
  spPr: XmlNode | undefined,
  layoutSpPr: XmlNode | undefined,
  masterSpPr: XmlNode | undefined,
): Geometry {
  const prst = child(spPr, 'prstGeom') ?? child(layoutSpPr, 'prstGeom') ?? child(masterSpPr, 'prstGeom');
  if (prst) {
    return {
      type: 'preset',
      preset: attr(prst, 'prst') ?? 'rect',
      adjust: parseAdjust(child(prst, 'avLst')),
    };
  }
  const cust = child(spPr, 'custGeom') ?? child(layoutSpPr, 'custGeom') ?? child(masterSpPr, 'custGeom');
  if (cust) return { type: 'custom', paths: parseCustomGeometry(cust) };
  return { type: 'preset', preset: 'rect', adjust: {} };
}

function parseAdjust(avLst: XmlNode | undefined): Record<string, number> {
  const adj: Record<string, number> = {};
  for (const gd of children(avLst, 'gd')) {
    const name = attr(gd, 'name');
    const fmla = attr(gd, 'fmla');
    if (name && fmla?.startsWith('val ')) {
      const n = Number(fmla.slice(4));
      if (!Number.isNaN(n)) adj[name] = n / 100000;
    }
  }
  return adj;
}

/** Resolve the solid color implied by a style reference (fillRef/lnRef/fontRef). */
export function styleRefColor(ref: XmlNode | undefined, ctx: SlideBuildCtx): Fill | undefined {
  if (!ref) return undefined;
  // idx 0 on a fillRef/lnRef means "no fill/line".
  if (attr(ref, 'idx') === '0') return undefined;
  const color = resolveColorEl(findColorEl(ref), ctx.colorCtx);
  return color ? { type: 'solid', color } : undefined;
}

export function resolveStroke(
  spPr: XmlNode | undefined,
  style: XmlNode | undefined,
  layoutSpPr: XmlNode | undefined,
  layoutStyle: XmlNode | undefined,
  masterSpPr: XmlNode | undefined,
  masterStyle: XmlNode | undefined,
  ctx: SlideBuildCtx,
): Stroke | undefined {
  const ln = child(spPr, 'ln') ?? child(layoutSpPr, 'ln') ?? child(masterSpPr, 'ln');
  const activeStyle = child(spPr, 'ln') ? style :
                       (child(layoutSpPr, 'ln') ? layoutStyle : masterStyle);
  const activeScope = child(spPr, 'ln') ? ctx.scopes.slide :
                       (child(layoutSpPr, 'ln') ? ctx.scopes.layout : ctx.scopes.master);

  if (ln) {
    if (child(ln, 'noFill')) return undefined;

    const stroke = strokeFromLn(ln, activeScope);
    if (stroke) return stroke;

    const lnRef = child(activeStyle ?? style ?? layoutStyle ?? masterStyle, 'lnRef');
    const styled = styleStroke(lnRef, ctx);
    if (styled) {
      const wEmu = attrNum(ln, 'w');
      return wEmu !== undefined ? { ...styled, width: Math.max(0.5, emuToPx(wEmu)) } : styled;
    }
    return undefined;
  }

  const lnRef = child(style, 'lnRef') ?? child(layoutStyle, 'lnRef') ?? child(masterStyle, 'lnRef');
  return styleStroke(lnRef, ctx);
}

export function styleStroke(lnRef: XmlNode | undefined, ctx: SlideBuildCtx): Stroke | undefined {
  if (!lnRef || attr(lnRef, 'idx') === '0') return undefined;
  const color = resolveColorEl(findColorEl(lnRef), ctx.colorCtx);
  return color ? { color, width: 1 } : undefined;
}

/**
 * Resolve a shape's effects: an inline `<a:effectLst>` (slide -> layout ->
 * master) wins; otherwise a `<a:effectRef idx>` selects a theme effect style,
 * with the effectRef's own color supplied as `phClr` for that style's colors.
 */
export function resolveEffects(
  spPr: XmlNode | undefined,
  style: XmlNode | undefined,
  layoutSpPr: XmlNode | undefined,
  layoutStyle: XmlNode | undefined,
  masterSpPr: XmlNode | undefined,
  masterStyle: XmlNode | undefined,
  ctx: SlideBuildCtx,
): Effect[] {
  const inline = child(spPr, 'effectLst') ?? child(layoutSpPr, 'effectLst') ?? child(masterSpPr, 'effectLst');
  if (inline) return parseEffectLst(inline, ctx.colorCtx);

  const effectRef = child(style, 'effectRef') ?? child(layoutStyle, 'effectRef') ?? child(masterStyle, 'effectRef');
  const idx = attrNum(effectRef, 'idx');
  if (!effectRef || !idx) return [];
  const effectStyle = ctx.theme.effectStyles[idx - 1];
  const styleEffectLst = child(effectStyle, 'effectLst');
  if (!styleEffectLst) return [];

  const phClr = resolveColorEl(findColorEl(effectRef), ctx.colorCtx);
  const colorCtx = phClr ? { ...ctx.colorCtx, phClr } : ctx.colorCtx;
  return parseEffectLst(styleEffectLst, colorCtx);
}

/** Assemble the text-style inheritance chain (most-specific first) for a shape. */
export function buildTextChain(
  sp: XmlNode,
  ph: { type: string; idx?: number } | undefined,
  layoutPh: PhInfo | undefined,
  masterPh: PhInfo | undefined,
  ctx: SlideBuildCtx,
): TextStyleChain {
  const chain: XmlNode[] = [];
  const own = lstStyleOf(sp);
  if (own) chain.push(own);
  const lo = lstStyleOf(layoutPh?.sp);
  if (lo) chain.push(lo);
  const mo = lstStyleOf(masterPh?.sp);
  if (mo) chain.push(mo);
  if (ph) {
    const masterStyle = child(ctx.masterTxStyles, masterStyleKey(ph.type));
    if (masterStyle) chain.push(masterStyle);
  }
  return new TextStyleChain(chain, ctx.colorCtx);
}
