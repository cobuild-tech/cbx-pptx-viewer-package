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

/** A node's children with <w:sdt>/<w:smartTag> wrappers flattened away. */
export function logicalChildren(node: XmlNode | undefined): XmlNode[] {
  if (!node) return [];
  const out: XmlNode[] = [];
  for (const c of node.children) {
    const name = localName(c.name);
    if (name === 'sdt') {
      out.push(...logicalChildren(child(c, 'sdtContent')));
    } else if (name === 'smartTag' || name === 'fldSimple') {
      // smartTag / simple field: transparent wrappers around runs.
      out.push(...logicalChildren(c));
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
