/**
 * Stable node identity for editable DOCX documents.
 *
 * A nodeId addresses an XML element as `"{partPath}#{indexPath}"`, where
 * indexPath is the chain of child indices from the part's root element
 * (e.g. `word/document.xml#1.4.0` = document.children[1].children[4].children[0]).
 *
 * Why index paths (not intrinsic ids): they are a pure function of tree
 * position, survive clone/serialize with no node pollution, and compose
 * correctly with undo/redo and version replay. Edits re-derive the IR after
 * every change, so the paths stamped on the DOM always reflect the current tree.
 */
import { resolveIndexPath, type XmlNode } from '../../oxml/xml.js';

export interface NodeRef {
  /** OPC part path, e.g. 'word/document.xml' or 'word/header1.xml'. */
  part: string;
  /** Child-index chain from the part's root element. */
  path: number[];
}

export function encodeNodeId(part: string, path: number[]): string {
  return `${part}#${path.join('.')}`;
}

export function decodeNodeId(id: string): NodeRef {
  const hash = id.lastIndexOf('#');
  const part = hash === -1 ? id : id.slice(0, hash);
  const pathStr = hash === -1 ? '' : id.slice(hash + 1);
  const path = pathStr === '' ? [] : pathStr.split('.').map((s) => Number(s));
  return { part, path };
}

/** The nodeId of the parent of the given node, or undefined if it is a root. */
export function parentNodeId(id: string): string | undefined {
  const { part, path } = decodeNodeId(id);
  if (path.length === 0) return undefined;
  return encodeNodeId(part, path.slice(0, -1));
}

/** The child index this nodeId occupies within its parent (or -1 for a root). */
export function indexOfNodeId(id: string): number {
  const { path } = decodeNodeId(id);
  return path.length === 0 ? -1 : path[path.length - 1]!;
}

/** Resolve a nodeId against an already-loaded part root. */
export function resolveNode(root: XmlNode, ref: NodeRef): XmlNode | undefined {
  return resolveIndexPath(root, ref.path);
}
