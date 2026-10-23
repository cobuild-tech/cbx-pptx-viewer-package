/**
 * Content-control flattening.
 *
 * WordprocessingML wraps real content in structured document tags (<w:sdt>,
 * used for form fields / content controls) and legacy <w:smartTag> elements.
 * These are transparent to layout: the actual paragraphs / rows / cells / runs
 * live inside <w:sdtContent> (or directly inside a smartTag). Parsers must see
 * through them, or all content-control content (which in real-world docs holds
 * table values, dates, names, …) is silently dropped.
 *
 * {@link logicalChildren} returns a node's effective children with sdt/smartTag
 * wrappers removed, recursively, at every nesting level.
 */
import { child, localName, type XmlNode } from '../oxml/xml.js';

/**
 * Wrappers whose content is simply part of the flow: content controls, smart
 * tags, custom-XML markup, bidi overrides, and tracked *insertions* (Word shows
 * an insertion's text; a tracked deletion — <w:del>/<w:moveFrom> — is not
 * unwrapped, so its runs never reach a parser).
 */
const TRANSPARENT = new Set(['smartTag', 'customXml', 'ins', 'moveTo', 'dir', 'bdo']);

/** Tracked-change wrappers: their content is inserted, deleted or moved text. */
const REVISIONS = new Set(['ins', 'del', 'moveTo', 'moveFrom']);

export interface LogicalOptions {
  /**
   * Keep tracked-change wrappers (<w:ins>, <w:del>, moves) as nodes, so a
   * paragraph parser can attribute their runs. Otherwise insertions unwrap and
   * deletions drop out.
   */
  keepRevisions?: boolean;
  /**
   * Keep <w:fldSimple> as a node instead of unwrapping it, so a paragraph
   * parser can read its field instruction (PAGE, NUMPAGES…).
   */
  keepFields?: boolean;
}

/** A node's children with sdt/smartTag/insertion wrappers flattened away. */
export function logicalChildren(node: XmlNode | undefined, opts: LogicalOptions = {}): XmlNode[] {
  if (!node) return [];
  const out: XmlNode[] = [];
  for (const c of node.children) {
    const name = localName(c.name);
    if (name === 'sdt') {
      out.push(...logicalChildren(child(c, 'sdtContent'), opts));
    } else if (opts.keepRevisions && REVISIONS.has(name)) {
      out.push(c);
    } else if (TRANSPARENT.has(name) || (name === 'fldSimple' && !opts.keepFields)) {
      out.push(...logicalChildren(c, opts));
    } else {
      out.push(c);
    }
  }
  return out;
}

/** {@link logicalChildren} filtered to a given local name. */
export function logicalChildrenNamed(node: XmlNode | undefined, name: string): XmlNode[] {
  return logicalChildren(node).filter((c) => localName(c.name) === name);
}
