/**
 * Shape-level structural edits — the XML half of direct manipulation.
 *
 * Where {@link ./xmlWrite.js} rewrites what a shape *says*, this rewrites where
 * it *is*: geometry write-back, deletion, and z-order. Three things make that
 * less trivial than setting a couple of attributes:
 *
 *   1. **Every shape kind keeps its transform somewhere else.** `<p:sp>`,
 *      `<p:pic>` and `<p:cxnSp>` hold an `<a:xfrm>` inside `<p:spPr>`; a group
 *      holds one inside `<p:grpSpPr>`; a `<p:graphicFrame>` holds a `<p:xfrm>`
 *      as its own direct child, in the presentation namespace rather than
 *      DrawingML.
 *   2. **The element may not exist yet.** A placeholder that inherits its
 *      position from the layout has no `<a:xfrm>` at all, so moving one means
 *      creating the element — and OOXML's schema is sequence-ordered, so it has
 *      to be created in the right place or PowerPoint rejects the file.
 *   3. **A group's `chOff`/`chExt` must survive untouched.** Children are mapped
 *      through the ext/chExt ratio, so leaving the child extent alone while
 *      changing `ext` is exactly what makes resizing a group scale everything
 *      inside it — the same trick PowerPoint uses.
 *
 * Z-order is document order in the shape tree, so reordering is a splice. As
 * with slide deletion, media a deleted shape referenced is deliberately left in
 * the package: other shapes commonly share an image part, and an unreferenced
 * one is harmless.
 */
import {
  child,
  createElement,
  insertChildAt,
  localName,
  removeAttr,
  setAttr,
  type XmlNode,
} from '../../oxml/xml.js';
import { pxToEmu, degToAngle } from '../../oxml/units.js';
import type { Transform } from '../model.js';

/** Shape-tree elements this module knows how to address. */
const SHAPE_NAMES = new Set(['sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame']);

/** Where a kind of shape keeps its transform. */
interface XfrmHost {
  /** The element that contains (or will contain) the xfrm. */
  parent: XmlNode;
  /** Index the xfrm must be inserted at if it has to be created. */
  index: number;
  /** Qualified name to create it under, e.g. `a:xfrm` or `p:xfrm`. */
  name: string;
  /** Prefix for the DrawingML children (`a:off`, `a:ext`). */
  childPrefix: string;
}

/** True if `node` is a shape-tree element with a transform we can write. */
export function isShapeNode(node: XmlNode): boolean {
  return SHAPE_NAMES.has(localName(node.name));
}

/**
 * Write `transform` into a shape's XML, creating the `<a:xfrm>` (and its
 * container) if the shape inherited its position instead of stating one.
 *
 * Returns false, having changed nothing, for a node that is not a shape.
 */
export function writeTransform(node: XmlNode, transform: Transform): boolean {
  const host = xfrmHostOf(node);
  if (!host) return false;

  const xfrm = ensureChild(host.parent, host.name, host.index);
  const p = host.childPrefix;

  // Rotation and mirroring live on the xfrm element itself. Omitting them is
  // how "not rotated" is spelled, so a zeroed value removes the attribute
  // rather than writing `rot="0"`.
  if (transform.rot) setAttr(xfrm, 'rot', String(degToAngle(transform.rot)));
  else removeAttr(xfrm, 'rot');
  setFlag(xfrm, 'flipH', transform.flipH);
  setFlag(xfrm, 'flipV', transform.flipV);

  // <a:off> then <a:ext>; a group's <a:chOff>/<a:chExt> follow and are left
  // exactly as they were, which is what scales the group's children.
  const off = ensureChild(xfrm, `${p}off`, 0);
  setAttr(off, 'x', String(pxToEmu(transform.x)));
  setAttr(off, 'y', String(pxToEmu(transform.y)));

  const ext = ensureChild(xfrm, `${p}ext`, indexAfter(xfrm, 'off'));
  // Extents are unsigned in the schema: a zero-width shape is expressible, an
  // inverted one is not.
  setAttr(ext, 'cx', String(Math.max(0, pxToEmu(transform.w))));
  setAttr(ext, 'cy', String(Math.max(0, pxToEmu(transform.h))));
  return true;
}

/**
 * Remove a shape from its parent shape tree. Returns false if the node is not
 * where `root` says it is.
 */
export function removeShape(root: XmlNode, node: XmlNode): boolean {
  const parent = parentOf(root, node);
  if (!parent) return false;
  parent.children = parent.children.filter((c) => c !== node);
  return true;
}

export type ZOrderMove = 'front' | 'back' | 'forward' | 'backward';

/**
 * Move a shape through the z-order, which is its position among its siblings.
 *
 * Only sibling *shapes* count: a shape tree opens with `<p:nvGrpSpPr>` and
 * `<p:grpSpPr>`, and stepping over those would both mis-count the move and
 * risk splicing a shape in front of them, which the schema forbids.
 *
 * Returns false if the shape is already at that end of the order.
 */
export function reorderShape(root: XmlNode, node: XmlNode, move: ZOrderMove): boolean {
  const parent = parentOf(root, node);
  if (!parent) return false;
  const shapes = parent.children.filter(isShapeNode);
  const from = shapes.indexOf(node);
  if (from === -1) return false;

  const to =
    move === 'front' ? shapes.length - 1
    : move === 'back' ? 0
    : move === 'forward' ? from + 1
    : from - 1;
  if (to === from || to < 0 || to >= shapes.length) return false;

  shapes.splice(from, 1);
  shapes.splice(to, 0, node);
  // Rebuild the child list by refilling the shape slots in their new order,
  // leaving every non-shape child (and its position) alone.
  let i = 0;
  parent.children = parent.children.map((c) => (isShapeNode(c) ? shapes[i++]! : c));
  return true;
}

/** The element whose children contain `node`, searched from `root`. */
export function parentOf(root: XmlNode, node: XmlNode): XmlNode | undefined {
  if (root.children.includes(node)) return root;
  for (const c of root.children) {
    const found = parentOf(c, node);
    if (found) return found;
  }
  return undefined;
}

function xfrmHostOf(node: XmlNode): XfrmHost | undefined {
  const kind = localName(node.name);
  const prefix = prefixOf(node.name);

  if (kind === 'graphicFrame') {
    // <p:graphicFrame> is nvGraphicFramePr, xfrm, graphic — and the xfrm is in
    // the presentation namespace here, not DrawingML.
    return {
      parent: node,
      index: indexAfterPrefix(node, 'nv'),
      name: `${prefix}xfrm`,
      childPrefix: 'a:',
    };
  }

  const container = kind === 'grpSp' ? 'grpSpPr' : 'spPr';
  if (!SHAPE_NAMES.has(kind)) return undefined;
  // The properties element is required by the schema, but a file we did not
  // write may still be missing it; create it after the non-visual properties.
  const parent = ensureChild(node, `${prefix}${container}`, indexAfterPrefix(node, 'nv'));
  return {
    parent,
    // <a:xfrm> opens CT_ShapeProperties and CT_GroupShapeProperties alike.
    index: 0,
    name: `${drawingPrefixIn(parent)}xfrm`,
    childPrefix: drawingPrefixIn(parent),
  };
}

/** Find a child by local name, or create it at `index`. */
function ensureChild(parent: XmlNode, name: string, index: number): XmlNode {
  const existing = child(parent, localName(name));
  if (existing) return existing;
  const created = createElement(name);
  insertChildAt(parent, index, created);
  return created;
}

/** Write a boolean xfrm attribute, or drop it when false — `flipH="0"` is
 * legal but PowerPoint omits it, and matching that keeps diffs small. */
function setFlag(node: XmlNode, name: string, on: boolean | undefined): void {
  if (on) setAttr(node, name, '1');
  else removeAttr(node, name);
}

/** Namespace prefix of a qualified name, including the colon ("p:sp" -> "p:"). */
function prefixOf(name: string): string {
  const i = name.indexOf(':');
  return i === -1 ? '' : name.slice(0, i + 1);
}

/**
 * The prefix DrawingML children should use inside `parent`. Real files always
 * use `a:`, but reading it off a sibling keeps us honest for a file that binds
 * the namespace to something else.
 */
function drawingPrefixIn(parent: XmlNode): string {
  for (const c of parent.children) {
    const p = prefixOf(c.name);
    if (p) return p;
  }
  return 'a:';
}

/** Index just past the child with this local name, or 0 if there is none. */
function indexAfter(parent: XmlNode, name: string): number {
  const i = parent.children.findIndex((c) => localName(c.name) === name);
  return i === -1 ? parent.children.length : i + 1;
}

/** Index just past the trailing non-visual properties element (`nv…Pr`). */
function indexAfterPrefix(parent: XmlNode, prefix: string): number {
  const i = parent.children.findIndex((c) => localName(c.name).startsWith(prefix));
  return i === -1 ? 0 : i + 1;
}
