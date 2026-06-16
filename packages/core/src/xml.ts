/**
 * Thin XML layer over fast-xml-parser.
 *
 * OOXML is order-sensitive: the order of children in a slide's shape tree is
 * the z-order, and runs / line-breaks / fields are interleaved inside a
 * paragraph. So we parse with `preserveOrder` and normalize the somewhat awkward
 * output into a simple, namespace-aware {@link XmlNode} tree with helpers.
 *
 * Namespace prefixes in real .pptx files are stable (Microsoft always emits
 * `p:`, `a:`, `r:`), but our lookups match by *local name* when the query has no
 * prefix, so we are robust to prefix variation.
 */
import { XMLParser } from 'fast-xml-parser';

export interface XmlNode {
  /** Qualified tag name including prefix, e.g. "p:sp", "a:off". */
  name: string;
  /** Attributes, keyed by their qualified name (e.g. "r:embed", "cx"). */
  attrs: Record<string, string>;
  /** Child elements, in document order. */
  children: XmlNode[];
  /** Concatenated direct text content of this element. */
  text: string;
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  preserveOrder: true,
  trimValues: false, // significant whitespace in <a:t> runs must survive
  parseTagValue: false,
  parseAttributeValue: false,
  processEntities: true,
  ignoreDeclaration: true,
  ignorePiTags: true,
});

/** Strip the namespace prefix from a qualified name ("a:off" -> "off"). */
export function localName(name: string): string {
  const i = name.indexOf(':');
  return i === -1 ? name : name.slice(i + 1);
}

interface RawEntry {
  [key: string]: unknown;
  ':@'?: Record<string, string>;
}

function buildNode(entry: RawEntry): XmlNode | null {
  const tag = Object.keys(entry).find((k) => k !== ':@');
  if (tag === undefined) return null;

  const node: XmlNode = {
    name: tag,
    attrs: entry[':@'] ? { ...entry[':@'] } : {},
    children: [],
    text: '',
  };

  const rawChildren = entry[tag];
  if (Array.isArray(rawChildren)) {
    for (const c of rawChildren as RawEntry[]) {
      if ('#text' in c) {
        node.text += String(c['#text']);
      } else {
        const child = buildNode(c);
        if (child) node.children.push(child);
      }
    }
  }
  return node;
}

/** Parse an XML string into the root {@link XmlNode}, or null if empty. */
export function parseXml(xml: string): XmlNode | null {
  const arr = parser.parse(xml) as RawEntry[];
  for (const entry of arr) {
    const node = buildNode(entry);
    if (node) return node;
  }
  return null;
}

function matches(node: XmlNode, name: string): boolean {
  return name.includes(':') ? node.name === name : localName(node.name) === name;
}

/** First child matching `name` (local name if `name` has no prefix). */
export function child(node: XmlNode | undefined, name: string): XmlNode | undefined {
  if (!node) return undefined;
  return node.children.find((c) => matches(c, name));
}

/** All children matching `name`, in document order. */
export function children(node: XmlNode | undefined, name: string): XmlNode[] {
  if (!node) return [];
  return node.children.filter((c) => matches(c, name));
}

/**
 * Resolve a descendant by a slash path of local names, e.g.
 * `path(sp, "spPr/xfrm/off")`. Returns the first match at each step.
 */
export function path(node: XmlNode | undefined, p: string): XmlNode | undefined {
  let cur: XmlNode | undefined = node;
  for (const seg of p.split('/')) {
    cur = child(cur, seg);
    if (!cur) return undefined;
  }
  return cur;
}

/** String attribute by qualified or local name, namespace-prefix-robust. */
export function attr(node: XmlNode | undefined, name: string): string | undefined {
  if (!node) return undefined;
  if (name in node.attrs) return node.attrs[name];
  const queryLocal = localName(name);
  const queryHasPrefix = name.includes(':');
  for (const [k, v] of Object.entries(node.attrs)) {
    const attrLocal = localName(k);
    const attrHasPrefix = k.includes(':');
    if (attrLocal === queryLocal) {
      if (queryHasPrefix) {
        if (attrHasPrefix) return v;
      } else {
        return v;
      }
    }
  }
  return undefined;
}

/** Numeric attribute, or `undefined` if absent/unparseable. */
export function attrNum(node: XmlNode | undefined, name: string): number | undefined {
  const v = attr(node, name);
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isNaN(n) ? undefined : n;
}

/** Boolean attribute. OOXML booleans are "1"/"0"/"true"/"false". */
export function attrBool(
  node: XmlNode | undefined,
  name: string,
  fallback = false,
): boolean {
  const v = attr(node, name);
  if (v === undefined) return fallback;
  return v === '1' || v === 'true';
}
