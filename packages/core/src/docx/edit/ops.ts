/**
 * Edit operations for DOCX — the only way the document is ever mutated.
 *
 * Every op targets OOXML nodes by nodeId and is applied *surgically* to the live
 * XmlNode tree of the affected part (never regenerated from the IR), which is
 * what preserves the document's structure. applyOp() returns the inverse op,
 * capturing any prior state, so the same machinery powers undo/redo and the
 * per-version diff log.
 *
 * Ops are plain JSON-serializable data so they can be persisted in a version's
 * op-log and replayed.
 */
import type { OpcPackage } from '../../oxml/package.js';
import {
  parseXml,
  serializeNode,
  createElement,
  cloneNode,
  child,
  attr,
  localName,
  resolveIndexPath,
  type XmlNode,
} from '../../oxml/xml.js';
import { decodeNodeId, encodeNodeId, parentNodeId, indexOfNodeId } from './nodeId.js';

// ─── Patch shapes ──────────────────────────────────────────────────────────────
//
// For boolean run props: true = explicit on, false = explicit off (overrides an
// inherited "on"), null = remove the override (fall back to the style chain).

export interface RunPropPatch {
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: boolean | null;
  strike?: boolean | null;
  /** Hex 'RRGGBB' (no leading #), or null to clear. */
  color?: string | null;
  /** Font size in points, or null to clear. */
  sizePt?: number | null;
  /** Font family, or null to clear. */
  font?: string | null;
}

export type ParaAlign = 'left' | 'center' | 'right' | 'justify';

export interface ParaPropPatch {
  align?: ParaAlign | null;
  /** Paragraph style id (e.g. 'Heading1'), or null to clear. */
  styleName?: string | null;
}

// ─── Op union ──────────────────────────────────────────────────────────────────

export type EditOp =
  | { kind: 'replaceText'; target: string; text: string }
  | { kind: 'setRunProps'; target: string; props: RunPropPatch }
  | { kind: 'setParaProps'; target: string; props: ParaPropPatch }
  | { kind: 'insertNode'; parent: string; index: number; xml: string }
  | { kind: 'removeNode'; target: string }
  | { kind: 'moveNode'; target: string; toIndex: number };

// ─── Resolver ───────────────────────────────────────────────────────────────────

interface Resolved {
  part: string;
  node: XmlNode;
  parent?: XmlNode;
  index: number;
}

function resolve(pkg: OpcPackage, nodeId: string): Resolved {
  const { part, path } = decodeNodeId(nodeId);
  const root = pkg.getXml(part);
  if (!root) throw new Error(`applyOp: missing part for nodeId "${nodeId}"`);
  if (path.length === 0) return { part, node: root, index: -1 };
  const parent = resolveIndexPath(root, path.slice(0, -1));
  const index = path[path.length - 1]!;
  const node = parent?.children[index];
  if (!parent || !node) throw new Error(`applyOp: cannot resolve nodeId "${nodeId}"`);
  return { part, node, parent, index };
}

// ─── apply (returns the inverse) ─────────────────────────────────────────────────

export function applyOp(pkg: OpcPackage, op: EditOp): EditOp {
  switch (op.kind) {
    case 'replaceText': {
      const { part, node } = resolve(pkg, op.target);
      const prev = getRunText(node);
      setRunText(node, op.text);
      pkg.markDirty(part);
      return { kind: 'replaceText', target: op.target, text: prev };
    }
    case 'setRunProps': {
      const { part, node } = resolve(pkg, op.target);
      const prev = readRunProps(node, op.props);
      applyRunProps(node, op.props);
      pkg.markDirty(part);
      return { kind: 'setRunProps', target: op.target, props: prev };
    }
    case 'setParaProps': {
      const { part, node } = resolve(pkg, op.target);
      const prev = readParaProps(node, op.props);
      applyParaProps(node, op.props);
      pkg.markDirty(part);
      return { kind: 'setParaProps', target: op.target, props: prev };
    }
    case 'insertNode': {
      const { part, node: parent } = resolve(pkg, op.parent);
      const inserted = parseXml(op.xml);
      if (!inserted) throw new Error('applyOp: insertNode given invalid xml');
      const idx = Math.max(0, Math.min(op.index, parent.children.length));
      parent.children.splice(idx, 0, inserted);
      pkg.markDirty(part);
      const insertedId = encodeNodeId(part, [...decodeNodeId(op.parent).path, idx]);
      return { kind: 'removeNode', target: insertedId };
    }
    case 'removeNode': {
      const { part, parent, node, index } = resolve(pkg, op.target);
      if (!parent) throw new Error('applyOp: cannot remove the root node');
      const xml = serializeNode(node);
      parent.children.splice(index, 1);
      pkg.markDirty(part);
      return { kind: 'insertNode', parent: parentNodeId(op.target)!, index, xml };
    }
    case 'moveNode': {
      const { part, parent, node, index } = resolve(pkg, op.target);
      if (!parent) throw new Error('applyOp: cannot move the root node');
      parent.children.splice(index, 1);
      const to = Math.max(0, Math.min(op.toIndex, parent.children.length));
      parent.children.splice(to, 0, node);
      pkg.markDirty(part);
      const parentPath = decodeNodeId(op.target).path.slice(0, -1);
      const movedId = encodeNodeId(part, [...parentPath, to]);
      return { kind: 'moveNode', target: movedId, toIndex: index };
    }
  }
}

// ─── Convenience op builders (pure; for the editor / UI) ─────────────────────────

export const ops = {
  replaceText: (target: string, text: string): EditOp => ({ kind: 'replaceText', target, text }),
  setRunProps: (target: string, props: RunPropPatch): EditOp => ({ kind: 'setRunProps', target, props }),
  setParaProps: (target: string, props: ParaPropPatch): EditOp => ({ kind: 'setParaProps', target, props }),
  remove: (target: string): EditOp => ({ kind: 'removeNode', target }),
  move: (target: string, toIndex: number): EditOp => ({ kind: 'moveNode', target, toIndex }),
  insertInto: (parent: string, index: number, xml: string): EditOp => ({ kind: 'insertNode', parent, index, xml }),
  insertAfter: (siblingId: string, xml: string): EditOp => ({
    kind: 'insertNode',
    parent: parentNodeId(siblingId)!,
    index: indexOfNodeId(siblingId) + 1,
    xml,
  }),
  insertBefore: (siblingId: string, xml: string): EditOp => ({
    kind: 'insertNode',
    parent: parentNodeId(siblingId)!,
    index: indexOfNodeId(siblingId),
    xml,
  }),
};

/** XML for a blank paragraph (an empty editable run). */
export function emptyParagraphXml(): string {
  return '<w:p><w:r><w:t xml:space="preserve"></w:t></w:r></w:p>';
}

/**
 * XML for a new paragraph that clones the source paragraph's `<w:pPr>` (so its
 * style, list numbering, indentation, etc. carry over) but has empty text.
 */
export function clonedParagraphXml(pkg: OpcPackage, sourceParaId: string): string {
  const { node } = resolve(pkg, sourceParaId);
  const p = createElement('w:p');
  const pPr = child(node, 'pPr');
  if (pPr) p.children.push(cloneNode(pPr));
  p.children.push(createElement('w:r', {}, [createElement('w:t', { 'xml:space': 'preserve' })]));
  return serializeNode(p);
}

/** Serialize the current XML of any node (e.g. an `<w:tr>` to clone for a new row). */
export function nodeXml(pkg: OpcPackage, nodeId: string): string {
  return serializeNode(resolve(pkg, nodeId).node);
}

/** Clone a node (e.g. an `<w:tr>`) with all `<w:t>` text emptied — for a blank new row. */
export function clonedBlankXml(pkg: OpcPackage, nodeId: string): string {
  const clone = cloneNode(resolve(pkg, nodeId).node);
  clearText(clone);
  return serializeNode(clone);
}

function clearText(node: XmlNode): void {
  if (localName(node.name) === 't') node.text = '';
  for (const c of node.children) clearText(c);
}

/**
 * Ops that split a run's text at [start, end) into up to three runs and apply
 * `props` to the middle piece — character-level formatting that preserves the
 * run's existing formatting. Applied together they form one undo step. A
 * whole-run selection short-circuits to a plain setRunProps (no split).
 */
export function splitRunOps(
  pkg: OpcPackage,
  runId: string,
  start: number,
  end: number,
  props: RunPropPatch,
): EditOp[] {
  const { node } = resolve(pkg, runId);
  const text = getRunText(node);
  const s = Math.max(0, Math.min(start, text.length));
  const e = Math.max(s, Math.min(end, text.length));

  if (s === 0 && e === text.length) {
    return [{ kind: 'setRunProps', target: runId, props }];
  }

  const rPr = child(node, 'rPr');
  const makeRun = (slice: string, extra?: RunPropPatch): XmlNode => {
    const r = createElement('w:r');
    if (rPr) r.children.push(cloneNode(rPr));
    const t = createElement('w:t');
    t.text = slice;
    if (needsSpacePreserve(slice)) t.attrs['xml:space'] = 'preserve';
    r.children.push(t);
    if (extra) applyRunProps(r, extra);
    return r;
  };

  const pieces: XmlNode[] = [];
  if (s > 0) pieces.push(makeRun(text.slice(0, s)));
  pieces.push(makeRun(text.slice(s, e), props));
  if (e < text.length) pieces.push(makeRun(text.slice(e)));

  const parent = parentNodeId(runId)!;
  const idx = indexOfNodeId(runId);
  const out: EditOp[] = [{ kind: 'removeNode', target: runId }];
  pieces.forEach((p, k) => out.push({ kind: 'insertNode', parent, index: idx + k, xml: serializeNode(p) }));
  return out;
}

// ─── Run text ─────────────────────────────────────────────────────────────────

/** Read a run's text the same way the parser does (`<w:t>` joined, `<w:br>` → \n). */
export function getRunText(rEl: XmlNode): string {
  const parts: string[] = [];
  for (const c of rEl.children) {
    const n = localName(c.name);
    if (n === 't') parts.push(c.text ?? '');
    else if (n === 'br') {
      const t = attr(c, 'w:type') ?? attr(c, 'type');
      if (!t || t === 'textWrapping') parts.push('\n');
    }
  }
  return parts.join('');
}

/**
 * Replace a run's text, preserving its `<w:rPr>` formatting. Newlines become
 * `<w:br/>`; segments with significant whitespace get xml:space="preserve".
 */
function setRunText(rEl: XmlNode, text: string): void {
  const rPr = child(rEl, 'rPr');
  const next: XmlNode[] = rPr ? [rPr] : [];
  const segments = text.split('\n');
  segments.forEach((seg, i) => {
    if (i > 0) next.push(createElement('w:br'));
    const t = createElement('w:t');
    t.text = seg;
    if (needsSpacePreserve(seg)) t.attrs['xml:space'] = 'preserve';
    next.push(t);
  });
  rEl.children = next;
}

function needsSpacePreserve(s: string): boolean {
  return s.length > 0 && (s !== s.trim() || /\s{2,}/.test(s) || s.includes('\t'));
}

// ─── Run properties (<w:rPr>) ────────────────────────────────────────────────────

// CT_RPr child order (subset we touch + common neighbours), for schema-valid output.
const RPR_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
  'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
  'bdr', 'shd', 'vertAlign', 'rtl', 'cs', 'lang',
];

function ensureRPr(rEl: XmlNode): XmlNode {
  let rPr = child(rEl, 'rPr');
  if (!rPr) {
    rPr = createElement('w:rPr');
    rEl.children.unshift(rPr);
  }
  return rPr;
}

function reorder(parent: XmlNode, order: string[]): void {
  const rank = (c: XmlNode) => {
    const i = order.indexOf(localName(c.name));
    return i === -1 ? order.length : i;
  };
  parent.children = parent.children
    .map((c, i) => [c, i] as const)
    .sort((a, b) => rank(a[0]) - rank(b[0]) || a[1] - b[1])
    .map(([c]) => c);
}

function dropChildren(parent: XmlNode, ...tags: string[]): void {
  parent.children = parent.children.filter((c) => !tags.includes(localName(c.name)));
}

function setToggle(rPr: XmlNode, tag: string, value: boolean | null, offVal = '0'): void {
  dropChildren(rPr, tag);
  if (value === true) rPr.children.push(createElement(`w:${tag}`));
  else if (value === false) rPr.children.push(createElement(`w:${tag}`, { 'w:val': offVal }));
}

function applyRunProps(rEl: XmlNode, p: RunPropPatch): void {
  const rPr = ensureRPr(rEl);
  if (p.bold !== undefined) setToggle(rPr, 'b', p.bold);
  if (p.italic !== undefined) setToggle(rPr, 'i', p.italic);
  if (p.strike !== undefined) setToggle(rPr, 'strike', p.strike);
  if (p.underline !== undefined) {
    dropChildren(rPr, 'u');
    if (p.underline === true) rPr.children.push(createElement('w:u', { 'w:val': 'single' }));
    else if (p.underline === false) rPr.children.push(createElement('w:u', { 'w:val': 'none' }));
  }
  if (p.color !== undefined) {
    dropChildren(rPr, 'color');
    if (p.color !== null) rPr.children.push(createElement('w:color', { 'w:val': p.color }));
  }
  if (p.sizePt !== undefined) {
    dropChildren(rPr, 'sz', 'szCs');
    if (p.sizePt !== null) {
      const hp = String(Math.round(p.sizePt * 2));
      rPr.children.push(createElement('w:sz', { 'w:val': hp }));
      rPr.children.push(createElement('w:szCs', { 'w:val': hp }));
    }
  }
  if (p.font !== undefined) {
    dropChildren(rPr, 'rFonts');
    if (p.font !== null) {
      rPr.children.push(createElement('w:rFonts', { 'w:ascii': p.font, 'w:hAnsi': p.font, 'w:cs': p.font }));
    }
  }
  reorder(rPr, RPR_ORDER);
  // Drop a now-empty <w:rPr> so undo can return to the original (no rPr) state.
  if (rPr.children.length === 0) rEl.children = rEl.children.filter((c) => c !== rPr);
}

function readToggle(rPr: XmlNode | undefined, tag: string): boolean | null {
  const el = child(rPr, tag);
  if (!el) return null;
  const v = attr(el, 'w:val') ?? attr(el, 'val');
  return v === '0' || v === 'false' ? false : true;
}

function readRunProps(rEl: XmlNode, patch: RunPropPatch): RunPropPatch {
  const rPr = child(rEl, 'rPr');
  const out: RunPropPatch = {};
  if (patch.bold !== undefined) out.bold = readToggle(rPr, 'b');
  if (patch.italic !== undefined) out.italic = readToggle(rPr, 'i');
  if (patch.strike !== undefined) out.strike = readToggle(rPr, 'strike');
  if (patch.underline !== undefined) {
    const el = child(rPr, 'u');
    out.underline = !el ? null : (attr(el, 'w:val') ?? attr(el, 'val')) === 'none' ? false : true;
  }
  if (patch.color !== undefined) {
    const el = child(rPr, 'color');
    out.color = el ? (attr(el, 'w:val') ?? attr(el, 'val') ?? null) : null;
  }
  if (patch.sizePt !== undefined) {
    const el = child(rPr, 'sz');
    const v = el ? attr(el, 'w:val') ?? attr(el, 'val') : undefined;
    const n = v === undefined ? NaN : Number(v);
    out.sizePt = Number.isNaN(n) ? null : n / 2;
  }
  if (patch.font !== undefined) {
    const el = child(rPr, 'rFonts');
    out.font = el ? (attr(el, 'w:ascii') ?? attr(el, 'ascii') ?? null) : null;
  }
  return out;
}

// ─── Paragraph properties (<w:pPr>) ──────────────────────────────────────────────

const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'numPr', 'pBdr', 'shd', 'tabs',
  'spacing', 'ind', 'contextualSpacing', 'jc', 'outlineLvl', 'rPr',
];

const ALIGN_TO_JC: Record<ParaAlign, string> = {
  left: 'left',
  center: 'center',
  right: 'right',
  justify: 'both',
};
const JC_TO_ALIGN: Record<string, ParaAlign> = {
  left: 'left', start: 'left', center: 'center', right: 'right', end: 'right',
  both: 'justify', distribute: 'justify',
};

function ensurePPr(pEl: XmlNode): XmlNode {
  let pPr = child(pEl, 'pPr');
  if (!pPr) {
    pPr = createElement('w:pPr');
    pEl.children.unshift(pPr);
  }
  return pPr;
}

function applyParaProps(pEl: XmlNode, p: ParaPropPatch): void {
  const pPr = ensurePPr(pEl);
  if (p.align !== undefined) {
    dropChildren(pPr, 'jc');
    if (p.align !== null) pPr.children.push(createElement('w:jc', { 'w:val': ALIGN_TO_JC[p.align] }));
  }
  if (p.styleName !== undefined) {
    dropChildren(pPr, 'pStyle');
    if (p.styleName !== null) pPr.children.push(createElement('w:pStyle', { 'w:val': p.styleName }));
  }
  reorder(pPr, PPR_ORDER);
  if (pPr.children.length === 0) pEl.children = pEl.children.filter((c) => c !== pPr);
}

function readParaProps(pEl: XmlNode, patch: ParaPropPatch): ParaPropPatch {
  const pPr = child(pEl, 'pPr');
  const out: ParaPropPatch = {};
  if (patch.align !== undefined) {
    const jc = child(pPr, 'jc');
    const v = jc ? attr(jc, 'w:val') ?? attr(jc, 'val') : undefined;
    out.align = v ? JC_TO_ALIGN[v] ?? 'left' : null;
  }
  if (patch.styleName !== undefined) {
    const ps = child(pPr, 'pStyle');
    out.styleName = ps ? (attr(ps, 'w:val') ?? attr(ps, 'val') ?? null) : null;
  }
  return out;
}
