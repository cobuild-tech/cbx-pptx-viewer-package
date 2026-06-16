/**
 * Shape-tree resolver: `<p:spTree>` -> {@link Shape}[].
 *
 * Walks shapes in document order (= z-order), unwrapping mc:AlternateContent.
 * For placeholders it inherits geometry and text styles from the matching
 * layout/master placeholder. Fills/strokes fall back to the shape's style
 * reference (`<p:style>`) into the theme. Groups keep their child coordinate
 * space so the renderer can apply the nested transform.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../xml.js';
import { emuToPx, angleToDeg } from '../units.js';
import type {
  Shape,
  Transform,
  Geometry,
  Fill,
  Stroke,
  PresetShape,
  PictureShape,
  GroupShape,
  ConnectorShape,
  FrameShape,
} from '../model.js';
import { type ColorContext, type Theme, resolveColorEl, findColorEl } from './color.js';
import { type ParseScope, parseFill, parseStroke, strokeFromLn } from './fill.js';
import { parseCustomGeometry } from '../geometry/custom.js';
import { parseTextBody } from '../parse/text.js';
import { TextStyleChain } from './textStyles.js';
import {
  type PhInfo,
  placeholderOf,
  matchPlaceholder,
  masterStyleKey,
  lstStyleOf,
} from './placeholders.js';
import { parseTable } from '../parse/table.js';

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

const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';
const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

export interface BuildOpts {
  /** Skip content placeholders (used when compositing layout/master shapes). */
  skipPlaceholders?: boolean;
}

/** Build the shape list for a shape tree (spTree or group), in z-order. */
export function buildShapes(
  tree: XmlNode,
  ctx: SlideBuildCtx,
  scope: ParseScope,
  opts: BuildOpts = {},
  groupFill?: Fill,
): Shape[] {
  const out: Shape[] = [];
  for (const node of expandChildren(tree)) {
    if (opts.skipPlaceholders && placeholderOf(node)) continue;
    const shape = buildShape(node, ctx, scope, groupFill);
    if (shape) out.push(shape);
  }
  return out;
}

/** Flatten mc:AlternateContent, preferring Fallback over Choice. */
function* expandChildren(tree: XmlNode): Generator<XmlNode> {
  for (const node of tree.children) {
    if (localName(node.name) === 'AlternateContent') {
      const fallback = child(node, 'Fallback') ?? child(node, 'Choice');
      if (fallback) yield* fallback.children;
    } else {
      yield node;
    }
  }
}

function buildShape(
  node: XmlNode,
  ctx: SlideBuildCtx,
  scope: ParseScope,
  groupFill?: Fill,
): Shape | null {
  switch (localName(node.name)) {
    case 'sp':
      return buildSp(node, ctx, scope, groupFill);
    case 'pic':
      return buildPic(node, ctx, scope);
    case 'grpSp':
      return buildGrp(node, ctx, scope, groupFill);
    case 'cxnSp':
      return buildCxn(node, scope);
    case 'graphicFrame':
      return buildFrame(node, ctx, scope);
    default:
      return null;
  }
}

/** True if a shape's `<spPr>` declares `<a:grpFill/>` (inherit the group's fill). */
function hasGrpFill(spPr: XmlNode | undefined): boolean {
  return !!spPr && spPr.children.some((c) => localName(c.name) === 'grpFill');
}

function parseXfrmEl(xfrm: XmlNode | undefined): Transform | undefined {
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

function parseGeometry(spPr: XmlNode | undefined): Geometry {
  return resolveGeometry(spPr, undefined, undefined);
}

function resolveTransform(
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

function resolveGeometry(
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
function styleRefColor(ref: XmlNode | undefined, ctx: SlideBuildCtx): Fill | undefined {
  if (!ref) return undefined;
  // idx 0 on a fillRef/lnRef means "no fill/line".
  if (attr(ref, 'idx') === '0') return undefined;
  const color = resolveColorEl(findColorEl(ref), ctx.colorCtx);
  return color ? { type: 'solid', color } : undefined;
}

function resolveStroke(
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
    
    let stroke = strokeFromLn(ln, activeScope);
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

function buildSp(
  sp: XmlNode,
  ctx: SlideBuildCtx,
  scope: ParseScope,
  groupFill?: Fill,
): PresetShape {
  const spPr = child(sp, 'spPr');
  const ph = placeholderOf(sp);
  const layoutPh = ph ? matchPlaceholder(ph, ctx.layoutPhs) : undefined;
  const masterPh = ph ? matchPlaceholder(ph, ctx.masterPhs) : undefined;

  const layoutSpPr = child(layoutPh?.sp, 'spPr');
  const masterSpPr = child(masterPh?.sp, 'spPr');
  const layoutStyle = child(layoutPh?.sp, 'style');
  const masterStyle = child(masterPh?.sp, 'style');

  const transform = resolveTransform(spPr, layoutSpPr, masterSpPr);

  const style = child(sp, 'style');
  
  // Resolve fill with inheritance
  let fill: Fill | undefined = parseFill(spPr, scope);
  if (fill === undefined && hasGrpFill(spPr)) {
    fill = groupFill;
  }
  if (fill === undefined && layoutSpPr) {
    fill = parseFill(layoutSpPr, ctx.scopes.layout);
    if (fill === undefined && hasGrpFill(layoutSpPr)) {
      fill = groupFill;
    }
  }
  if (fill === undefined && masterSpPr) {
    fill = parseFill(masterSpPr, ctx.scopes.master);
    if (fill === undefined && hasGrpFill(masterSpPr)) {
      fill = groupFill;
    }
  }
  if (fill === undefined) {
    const fRef = child(style, 'fillRef') ?? child(layoutStyle, 'fillRef') ?? child(masterStyle, 'fillRef');
    fill = styleRefColor(fRef, ctx);
  }
  if (fill === undefined) {
    fill = { type: 'none' };
  }

  // Resolve stroke with inheritance
  const stroke = resolveStroke(spPr, style, layoutSpPr, layoutStyle, masterSpPr, masterStyle, ctx);

  const shape: PresetShape = {
    kind: 'shape',
    geom: resolveGeometry(spPr, layoutSpPr, masterSpPr),
    fill,
  };
  if (transform) shape.transform = transform;
  if (stroke) shape.stroke = stroke;
  if (ph) shape.placeholder = ph;

  const txBody = child(sp, 'txBody');
  if (txBody && children(txBody, 'p').length > 0) {
    const chain = buildTextChain(sp, ph, layoutPh, masterPh, ctx);
    shape.text = parseTextBody(txBody, chain, ctx.colorCtx, scope);
  }
  return shape;
}

function styleStroke(lnRef: XmlNode | undefined, ctx: SlideBuildCtx): Stroke | undefined {
  if (!lnRef || attr(lnRef, 'idx') === '0') return undefined;
  const color = resolveColorEl(findColorEl(lnRef), ctx.colorCtx);
  return color ? { color, width: 1 } : undefined;
}

/** Assemble the text-style inheritance chain (most-specific first) for a shape. */
function buildTextChain(
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

function buildPic(pic: XmlNode, ctx: SlideBuildCtx, scope: ParseScope): PictureShape | null {
  const spPr = child(pic, 'spPr');
  const blipFill = child(pic, 'blipFill');
  const blip = child(blipFill, 'blip');
  const rId = attr(blip, 'embed') ?? attr(blip, 'link');
  const part = rId ? scope.resolveImage(rId) : undefined;
  if (!part) return null;

  const shape: PictureShape = { kind: 'picture', part, fill: { type: 'none' } };

  // Picture placeholders inherit geometry from the layout/master, like text ones.
  const ph = placeholderOf(pic);
  const layoutPh = ph ? matchPlaceholder(ph, ctx.layoutPhs) : undefined;
  const masterPh = ph ? matchPlaceholder(ph, ctx.masterPhs) : undefined;

  const layoutSpPr = child(layoutPh?.sp, 'spPr');
  const masterSpPr = child(masterPh?.sp, 'spPr');
  const layoutStyle = child(layoutPh?.sp, 'style');
  const masterStyle = child(masterPh?.sp, 'style');

  const transform = resolveTransform(spPr, layoutSpPr, masterSpPr);
  if (transform) shape.transform = transform;
  if (ph) shape.placeholder = ph;

  // A non-rectangular preset clips the image (e.g. picture cropped to a circle).
  if (child(spPr, 'prstGeom') || child(spPr, 'custGeom') ||
      child(layoutSpPr, 'prstGeom') || child(layoutSpPr, 'custGeom') ||
      child(masterSpPr, 'prstGeom') || child(masterSpPr, 'custGeom')) {
    shape.geom = resolveGeometry(spPr, layoutSpPr, masterSpPr);
  }

  const stroke = resolveStroke(spPr, child(pic, 'style'), layoutSpPr, layoutStyle, masterSpPr, masterStyle, ctx);
  if (stroke) shape.stroke = stroke;

  const srcRect = child(blipFill, 'srcRect');
  if (srcRect) {
    shape.crop = {
      l: (attrNum(srcRect, 'l') ?? 0) / 100000,
      t: (attrNum(srcRect, 't') ?? 0) / 100000,
      r: (attrNum(srcRect, 'r') ?? 0) / 100000,
      b: (attrNum(srcRect, 'b') ?? 0) / 100000,
    };
  }
  return shape;
}

function buildGrp(
  grp: XmlNode,
  ctx: SlideBuildCtx,
  scope: ParseScope,
  parentFill?: Fill,
): GroupShape | null {
  const grpSpPr = child(grp, 'grpSpPr');
  const xfrm = child(grpSpPr, 'xfrm');
  const transform = parseXfrmEl(xfrm);
  const chOff = child(xfrm, 'chOff');
  const chExt = child(xfrm, 'chExt');

  // The group's own fill is what `<a:grpFill/>` children resolve to; a nested
  // group with its own grpFill inherits the parent group's fill in turn.
  const groupFill = parseFill(grpSpPr, scope) ?? (hasGrpFill(grpSpPr) ? parentFill : undefined);

  const shape: GroupShape = {
    kind: 'group',
    children: buildShapes(grp, ctx, scope, {}, groupFill),
    childOffset: {
      x: emuToPx(attrNum(chOff, 'x') ?? 0),
      y: emuToPx(attrNum(chOff, 'y') ?? 0),
      w: emuToPx(attrNum(chExt, 'cx') ?? transform?.w ?? 0) || transform?.w || 1,
      h: emuToPx(attrNum(chExt, 'cy') ?? transform?.h ?? 0) || transform?.h || 1,
    },
  };
  if (transform) shape.transform = transform;
  return shape;
}

function buildCxn(cxn: XmlNode, scope: ParseScope): ConnectorShape {
  const spPr = child(cxn, 'spPr');
  const shape: ConnectorShape = { kind: 'connector', geom: parseGeometry(spPr) };
  const transform = parseXfrmEl(child(spPr, 'xfrm'));
  if (transform) shape.transform = transform;
  const stroke = parseStroke(spPr, scope);
  if (stroke) shape.stroke = stroke;
  return shape;
}

function buildFrame(frame: XmlNode, ctx: SlideBuildCtx, scope: ParseScope): FrameShape {
  const transform = parseXfrmEl(child(frame, 'xfrm'));
  const graphicData = child(child(frame, 'graphic'), 'graphicData');
  const uri = attr(graphicData, 'uri') ?? '';

  let frameType: FrameShape['frameType'] = 'unknown';
  if (uri === TABLE_URI) frameType = 'table';
  else if (uri === CHART_URI) frameType = 'chart';
  else if (uri === DIAGRAM_URI) frameType = 'diagram';

  const shape: FrameShape = { kind: 'frame', frameType };
  if (transform) shape.transform = transform;

  if (frameType === 'table') {
    const tbl = child(graphicData, 'tbl');
    if (tbl) shape.table = parseTable(tbl, ctx.colorCtx, scope);
  }
  return shape;
}
