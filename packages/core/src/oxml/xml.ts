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

// ─── Serialization (write-back) ───────────────────────────────────────────────
//
// The reverse of parseXml. We hand-roll this rather than use fast-xml-parser's
// XMLBuilder because the contract we need is *structural* fidelity, not byte
// fidelity: on export, untouched parts are re-zipped from their original bytes
// (see OpcPackage.toBytes), and only edited parts are re-serialized — Word
// re-parses them, so exact byte layout is irrelevant as long as the XML is
// valid and structurally identical when re-parsed.
//
// Faithfulness guarantee: parseXml(serializeXml(node)) deep-equals the original
// node for any WordprocessingML tree. Children are emitted in order, attributes
// in insertion order, and the concatenated `.text` is emitted after children
// (OOXML structural elements have no mixed content, so position is immaterial;
// re-parsing concatenates direct text back into `.text` identically).

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

function escapeText(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttr(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Serialize a single node and its subtree to an XML fragment (no declaration). */
export function serializeNode(node: XmlNode): string {
  let attrs = '';
  for (const [k, v] of Object.entries(node.attrs)) {
    attrs += ` ${k}="${escapeAttr(v)}"`;
  }
  const hasChildren = node.children.length > 0;
  const hasText = node.text.length > 0;
  if (!hasChildren && !hasText) return `<${node.name}${attrs}/>`;

  let inner = '';
  for (const c of node.children) inner += serializeNode(c);
  if (hasText) inner += escapeText(node.text);
  return `<${node.name}${attrs}>${inner}</${node.name}>`;
}

/** Serialize a part root to a complete XML document string (with declaration). */
export function serializeXml(node: XmlNode): string {
  return XML_DECL + serializeNode(node);
}

// ─── Mutation helpers ─────────────────────────────────────────────────────────
//
// XmlNode is a plain mutable object, so these are thin, but centralizing them
// keeps the edit layer readable and the invariants in one place.

/** Construct a new element node. */
export function createElement(
  name: string,
  attrs: Record<string, string> = {},
  children: XmlNode[] = [],
  text = '',
): XmlNode {
  return { name, attrs: { ...attrs }, children, text };
}

/** Deep-clone a node and its subtree (no shared references). */
export function cloneNode(node: XmlNode): XmlNode {
  return {
    name: node.name,
    attrs: { ...node.attrs },
    children: node.children.map(cloneNode),
    text: node.text,
  };
}

/** Replace a node's direct text content. */
export function setText(node: XmlNode, text: string): void {
  node.text = text;
}

/** Set (or add) an attribute by its qualified name. */
export function setAttr(node: XmlNode, name: string, value: string): void {
  node.attrs[name] = value;
}

/** Remove an attribute, matching by qualified or local name (prefix-robust). */
export function removeAttr(node: XmlNode, name: string): void {
  if (name in node.attrs) {
    delete node.attrs[name];
    return;
  }
  const queryLocal = localName(name);
  for (const k of Object.keys(node.attrs)) {
    if (localName(k) === queryLocal) delete node.attrs[k];
  }
}

export function insertChildAt(parent: XmlNode, index: number, child: XmlNode): void {
  parent.children.splice(Math.max(0, Math.min(index, parent.children.length)), 0, child);
}

/**
 * Insert `node` into `parent` at the position `order` dictates, i.e. before the
 * first existing child that must follow it.
 *
 * OOXML content models are sequence-ordered: an element in the wrong slot makes
 * Office call the part corrupt. `order` lists the local names in schema order;
 * a name it does not mention is appended, since we have nothing to place it by.
 */
export function insertInOrder(parent: XmlNode, node: XmlNode, order: string[]): void {
  const rank = order.indexOf(localName(node.name));
  if (rank === -1) {
    parent.children.push(node);
    return;
  }
  const at = parent.children.findIndex((c) => {
    const r = order.indexOf(localName(c.name));
    return r !== -1 && r > rank;
  });
  if (at === -1) parent.children.push(node);
  else parent.children.splice(at, 0, node);
}

export function removeChildAt(parent: XmlNode, index: number): XmlNode | undefined {
  if (index < 0 || index >= parent.children.length) return undefined;
  return parent.children.splice(index, 1)[0];
}

export function replaceChildAt(parent: XmlNode, index: number, child: XmlNode): void {
  if (index < 0 || index >= parent.children.length) return;
  parent.children[index] = child;
}

/**
 * Resolve a chain of child indices from `root` to a descendant node.
 * e.g. resolveIndexPath(documentRoot, [1, 4, 0]) === body.children[4].children[0]
 * Returns undefined if any index is out of range.
 */
export function resolveIndexPath(root: XmlNode, indexPath: number[]): XmlNode | undefined {
  let cur: XmlNode | undefined = root;
  for (const i of indexPath) {
    if (!cur || i < 0 || i >= cur.children.length) return undefined;
    cur = cur.children[i];
  }
  return cur;
}
