/**
 * Segment list -> `<w:p>` XML.
 *
 * Same rule as the PPTX writer: *reuse over recreation*. When a segment still
 * maps to its source `<w:t>` with unchanged text and no formatting override,
 * the original `<w:r>` is spliced back untouched, so an edit to one word leaves
 * every other run — and every property the parser never read — byte-identical.
 *
 * The DOCX wrinkle is that one `<w:r>` can hold several `<w:t>` children, so a
 * source run identifies a `<w:t>`, and rebuilding emits one `<w:r>` per segment
 * cloning that `<w:t>`'s owning run properties. Splitting a multi-`<w:t>` run
 * into separate runs is semantically identical in WordprocessingML.
 */
import {
  child,
  children,
  createElement,
  cloneNode,
  localName,
  setAttr,
  type XmlNode,
} from '../../oxml/xml.js';
import type { DocxParagraph, DocxRun } from '../model.js';
import type { DocxSource } from '../document/context.js';
import { applyFormat, isEmptyFormat, type DocxRunFormat } from './format.js';

/** One contiguous stretch of text that shares formatting. */
export interface DocxSegment {
  text: string;
  /** The model run this text inherits its `<w:rPr>` from, if any. */
  src?: DocxRun;
  /** Toolbar override applied on top of the source run's properties. */
  format?: DocxRunFormat;
  /** A hard line break (`<w:br/>`) before this segment's text. */
  breakBefore?: boolean;
  /** A tab stop (`<w:tab/>`) before this segment's text. */
  tabBefore?: boolean;
}

/** A paragraph after editing: the source it came from, plus its new content. */
export interface DocxParaEdit {
  /** The model paragraph this maps to; undefined for a newly created one. */
  src?: DocxParagraph;
  segments: DocxSegment[];
}

/** Resolves a model object back to the XML it was parsed from. */
export type DocxSourceLookup = (model: object) => DocxSource | undefined;

function prefixOf(name: string): string {
  const i = name.indexOf(':');
  return i === -1 ? '' : name.slice(0, i);
}

/**
 * The WordprocessingML prefix in use, learned from an existing node rather than
 * assumed — not every producer writes "w:".
 */
function wordPrefix(body: XmlNode): string {
  const stack: XmlNode[] = [body];
  while (stack.length) {
    const node = stack.shift()!;
    if (['p', 'r', 't', 'pPr', 'rPr'].includes(localName(node.name))) return prefixOf(node.name);
    stack.push(...node.children);
  }
  return prefixOf(body.name) || 'w';
}

/**
 * `<w:t>` needs `xml:space="preserve"` whenever its text has leading or
 * trailing whitespace — Word trims it otherwise. This matters far more in DOCX
 * than PPTX, because runs routinely start or end mid-sentence on a space.
 */
function makeTextNode(prefix: string, text: string): XmlNode {
  const node = createElement(prefix ? `${prefix}:t` : 't', {}, [], text);
  if (text !== text.trim()) setAttr(node, 'xml:space', 'preserve');
  return node;
}

/** The `<w:rPr>` of a source run's owning `<w:r>`, cloned. */
function cloneRPr(src: DocxSource | undefined, prefix: string): XmlNode {
  const existing = src?.owner ? child(src.owner, 'rPr') : undefined;
  return existing ? cloneNode(existing) : createElement(prefix ? `${prefix}:rPr` : 'rPr');
}

/**
 * A segment can reuse its original `<w:r>` verbatim only when it is the run's
 * sole `<w:t>`, its text is unchanged, it carries no formatting override, and
 * it needs no break/tab the original didn't already have.
 */
function canReuse(seg: DocxSegment, src: DocxSource | undefined): boolean {
  if (!src?.owner || !isEmptyFormat(seg.format)) return false;
  if (children(src.owner, 't').length !== 1) return false;
  if ((src.node.text ?? '') !== seg.text) return false;
  const hadBreak = child(src.owner, 'br') !== undefined;
  const hadTab = child(src.owner, 'tab') !== undefined;
  return !!seg.breakBefore === hadBreak && !!seg.tabBefore === hadTab;
}

function buildRunNode(seg: DocxSegment, lookup: DocxSourceLookup, prefix: string): XmlNode {
  const src = seg.src ? lookup(seg.src) : undefined;
  if (canReuse(seg, src)) return src!.owner!;

  const q = (n: string) => (prefix ? `${prefix}:${n}` : n);
  const rPr = cloneRPr(src, prefix);
  if (seg.format && !isEmptyFormat(seg.format)) applyFormat(rPr, seg.format, prefix);

  const kids: XmlNode[] = [rPr];
  // Break and tab precede the text they sit in front of, inside the same run.
  if (seg.breakBefore) kids.push(createElement(q('br')));
  if (seg.tabBefore) kids.push(createElement(q('tab')));
  kids.push(makeTextNode(prefix, seg.text));

  return createElement(q('r'), {}, kids);
}

/**
 * Build one `<w:p>`. The source paragraph's `<w:pPr>` is carried over (cloned
 * when the paragraph is new, e.g. after a split) so the new paragraph keeps its
 * style, indent, numbering and alignment.
 */
function buildParaNode(pe: DocxParaEdit, lookup: DocxSourceLookup, prefix: string): XmlNode {
  const srcNode = pe.src ? lookup(pe.src)?.node : undefined;
  const pNode = createElement(prefix ? `${prefix}:p` : 'p');
  if (srcNode) {
    pNode.attrs = { ...srcNode.attrs };
    // w14:paraId must be unique per paragraph; a split would duplicate it.
    delete pNode.attrs['w14:paraId'];
    delete pNode.attrs['w14:textId'];
  }

  const pPr = srcNode ? child(srcNode, 'pPr') : undefined;
  if (pPr) pNode.children.push(cloneNode(pPr));

  for (const seg of pe.segments) {
    if (!seg.text && !seg.breakBefore && !seg.tabBefore) continue;
    pNode.children.push(buildRunNode(seg, lookup, prefix));
  }

  return pNode;
}

/**
 * Replace a paragraph's node in the body with the (one or more) paragraphs it
 * became. Returns the nodes that were written, in document order.
 *
 * Editing is addressed per source paragraph rather than per body, because a
 * DOCX body is one long flow — there is no `<a:txBody>`-sized unit to rewrite.
 */
export function writeParagraphs(
  body: XmlNode,
  original: XmlNode,
  edits: DocxParaEdit[],
  lookup: DocxSourceLookup,
): XmlNode[] {
  const prefix = wordPrefix(body);
  const built = edits.map((pe) => buildParaNode(pe, lookup, prefix));

  // A paragraph that lost all its content still exists as an empty paragraph.
  if (built.length === 0) built.push(createElement(prefix ? `${prefix}:p` : 'p'));

  const parent = findParent(body, original) ?? body;
  const at = parent.children.indexOf(original);
  if (at === -1) parent.children.push(...built);
  else parent.children.splice(at, 1, ...built);

  return built;
}

/**
 * The element that directly contains `target`. Needed because a paragraph may
 * sit inside a table cell or a content control rather than the body itself.
 */
export function findParent(root: XmlNode, target: XmlNode): XmlNode | undefined {
  if (root.children.includes(target)) return root;
  for (const c of root.children) {
    const found = findParent(c, target);
    if (found) return found;
  }
  return undefined;
}
