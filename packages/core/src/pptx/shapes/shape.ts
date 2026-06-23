/**
 * Shape-tree resolver: `<p:spTree>` -> {@link Shape}[].
 *
 * Walks shapes in document order (= z-order), unwrapping mc:AlternateContent,
 * and dispatches each node to the builder for its kind. Generic shapes (sp),
 * groups (grpSp), connectors (cxnSp) and graphic frames live here; richer kinds
 * (pictures, tables, and later diagrams/charts) live in their own slices and
 * are dispatched to from here. Shared resolution lives in {@link ./props.js}.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { emuToPx } from '../../oxml/units.js';
import type {
  Shape,
  Fill,
  PresetShape,
  GroupShape,
  ConnectorShape,
  FrameShape,
} from '../model.js';
import type { ParseScope } from '../scope.js';
import {
  type SlideBuildCtx,
  type BuildOpts,
  hasGrpFill,
  parseXfrmEl,
  parseGeometry,
  resolveTransform,
  resolveGeometry,
  resolveStroke,
  resolveEffects,
  styleRefColor,
  buildTextChain,
} from './props.js';
import { parseFill, parseStroke } from './fill.js';
import { placeholderOf, matchPlaceholder } from './placeholders.js';
import { buildPic } from '../pictures/picture.js';
import { parseTable } from '../tables/table.js';
import { parseTextBody } from '../text/text.js';

export type { SlideBuildCtx, SlideScopes, BuildOpts } from './props.js';

const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';
const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

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

  // Resolve fill with inheritance.
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

  // Resolve stroke with inheritance.
  const stroke = resolveStroke(spPr, style, layoutSpPr, layoutStyle, masterSpPr, masterStyle, ctx);
  const effects = resolveEffects(spPr, style, layoutSpPr, layoutStyle, masterSpPr, masterStyle, ctx);

  const shape: PresetShape = {
    kind: 'shape',
    geom: resolveGeometry(spPr, layoutSpPr, masterSpPr),
    fill,
  };
  if (transform) shape.transform = transform;
  if (stroke) shape.stroke = stroke;
  if (effects.length) shape.effects = effects;
  if (ph) shape.placeholder = ph;

  const txBody = child(sp, 'txBody');
  if (txBody && children(txBody, 'p').length > 0) {
    const chain = buildTextChain(sp, ph, layoutPh, masterPh, ctx);
    shape.text = parseTextBody(txBody, chain, ctx.colorCtx, scope);
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
