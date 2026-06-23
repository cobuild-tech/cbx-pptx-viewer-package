/**
 * Placeholder indexing and matching.
 *
 * A placeholder on a slide inherits geometry and text styles from the matching
 * placeholder on its layout, which inherits from the master. Matching is by
 * `idx` when present, otherwise by `type` (with title-like types grouped).
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';

export interface PhInfo {
  type: string;
  idx?: number;
  sp: XmlNode;
}

const TITLE_TYPES = new Set(['title', 'ctrTitle']);

function phGroup(type: string): string {
  return TITLE_TYPES.has(type) ? 'title' : type;
}

/** The placeholder descriptor for a shape, or undefined if it isn't one. */
export function placeholderOf(sp: XmlNode): { type: string; idx?: number } | undefined {
  const ph = findPh(sp);
  if (!ph) return undefined;
  const type = attr(ph, 'type') ?? 'body';
  const idx = attrNum(ph, 'idx');
  return idx === undefined ? { type } : { type, idx };
}

function findPh(sp: XmlNode): XmlNode | undefined {
  // <p:nvSpPr><p:nvPr><p:ph .../></p:nvPr></p:nvSpPr> (or nvPicPr/nvGrpSpPr...)
  const nv = sp.children.find((c) => localName(c.name).startsWith('nv'));
  return child(child(nv, 'nvPr'), 'ph');
}

/** Index all placeholder shapes directly under a shape tree. */
export function indexPlaceholders(spTree: XmlNode | undefined): PhInfo[] {
  if (!spTree) return [];
  const out: PhInfo[] = [];
  for (const sp of spTree.children) {
    const ph = placeholderOf(sp);
    if (ph) out.push({ ...ph, sp });
  }
  return out;
}

/** Find the placeholder in `candidates` that a slide placeholder inherits from. */
export function matchPlaceholder(
  target: { type: string; idx?: number },
  candidates: PhInfo[],
): PhInfo | undefined {
  if (target.idx !== undefined) {
    const byIdx = candidates.find((c) => c.idx === target.idx);
    if (byIdx) return byIdx;
  }
  const group = phGroup(target.type);
  const byType = candidates.find((c) => phGroup(c.type) === group);
  if (byType) return byType;
  // Title placeholders also match a generic body if no title exists, and vice versa.
  return candidates.find((c) => c.idx === target.idx);
}

/** The master text-style key (in `<p:txStyles>`) for a placeholder type. */
export function masterStyleKey(type: string | undefined): string {
  if (type && TITLE_TYPES.has(type)) return 'titleStyle';
  if (type === 'body' || type === 'subTitle' || type === undefined) return 'bodyStyle';
  return 'otherStyle';
}

/** Helper used by the resolver to read a child shape's txBody/lstStyle, if any. */
export function lstStyleOf(sp: XmlNode | undefined): XmlNode | undefined {
  return child(child(sp, 'txBody'), 'lstStyle');
}
