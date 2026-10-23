/**
 * The editor's view of a paragraph: one line of *atoms*, each a character of
 * text, a tab, a line break, or an object (a picture, an equation, a note
 * reference) — so that "offset 12 in this paragraph" means one thing to the
 * caret, to the DOM and to the XML alike.
 *
 * The atoms are derived from the parsed model, not by re-walking the XML: the
 * parser already decided what is content (field instructions are not, an
 * anchored drawing is not, a deleted run is), and it recorded the exact XML
 * node behind every run (`DocxDocument.sourceOf`). Deriving from that keeps the
 * editor and the renderer in agreement by construction.
 */
import { localName, type XmlNode } from '../../oxml/xml.js';
import type { DocxDocument } from '../document/document.js';
import type { DocxBlock, DocxDrawing, DocxParagraph, DocxRun } from '../model.js';

export type SegKind = 'text' | 'tab' | 'br' | 'obj';

/** A stretch of atoms backed by one XML node. */
export interface Seg {
  /** Offset of the first atom in the paragraph. */
  start: number;
  len: number;
  kind: SegKind;
  /** `<w:t>` for text, `<w:tab>`/`<w:br>`, or the object's own element. */
  node: XmlNode;
  run: DocxRun;
}

/** Where one model run's atoms sit in its paragraph. */
export interface RunSpan {
  /** Offset of the run's line break, if it has one of its own. */
  brAt?: number;
  /** Offset of the run's tab, if it has one of its own. */
  tabAt?: number;
  /** Offset of the run's content (text or object). */
  at: number;
  /** Atoms of content: its text length, 1 for an object, 0 for none. */
  len: number;
  kind: 'text' | 'obj' | 'none';
}

export interface ParaMap {
  para: DocxParagraph;
  /** The `<w:p>`. */
  node: XmlNode;
  segs: Seg[];
  length: number;
  /** One character per atom: text as is, `\t`, `\n`, U+FFFC for an object. */
  text: string;
  runs: Map<DocxRun, RunSpan>;
  /** The same spans by each run's source element (what a drawn run's key names). */
  bySource: Map<XmlNode, RunSpan>;
}

/** The character standing for an object atom in {@link ParaMap.text}. */
export const OBJECT_CHAR = '￼';

/**
 * A comment's reference mark is not an atom: it is invisible, and a Backspace
 * that silently removed it would look like a key that did nothing.
 */
function isObjectRun(run: DocxRun): boolean {
  return !!(run.drawing || run.math || run.note || run.separator);
}

/** Build the atom map of a parsed paragraph (an original, never a layout copy). */
export function mapParagraph(doc: DocxDocument, para: DocxParagraph): ParaMap | undefined {
  const pSrc = doc.sourceOf(para);
  if (!pSrc) return undefined;
  const segs: Seg[] = [];
  const runs = new Map<DocxRun, RunSpan>();
  const bySource = new Map<XmlNode, RunSpan>();
  let off = 0;
  let text = '';
  const push = (kind: SegKind, node: XmlNode, run: DocxRun, chars: string) => {
    segs.push({ start: off, len: chars.length, kind, node, run });
    off += chars.length;
    text += chars;
  };
  for (const run of para.runs) {
    const src = doc.sourceOf(run);
    const span: RunSpan = { at: 0, len: 0, kind: 'none' };
    if (run.breakBefore && src?.br) {
      span.brAt = off;
      push('br', src.br, run, '\n');
    }
    if (run.tabBefore && src?.tab) {
      span.tabAt = off;
      push('tab', src.tab, run, '\t');
    }
    span.at = off;
    const local = src ? localName(src.node.name) : '';
    if (src && (local === 't' || local === 'delText')) {
      span.kind = 'text';
      span.len = run.text.length;
      if (run.text) push('text', src.node, run, run.text);
    } else if (src && local !== 'r' && (isObjectRun(run) || run.text)) {
      // A symbol, a non-breaking hyphen, a note reference: one element, one atom.
      span.kind = 'obj';
      span.len = 1;
      push('obj', src.node, run, OBJECT_CHAR);
    }
    runs.set(run, span);
    if (src) bySource.set(src.node, span);
  }
  return { para, node: pSrc.node, segs, length: off, text, runs, bySource };
}

/** Text a run holds that the user may not type into: field results and deleted text. */
export function isLockedRun(run: DocxRun): boolean {
  return run.fieldCode !== undefined || run.revision?.kind === 'del';
}

/**
 * Every paragraph of the main story in reading order — body paragraphs and
 * those in table cells — and, separately, those inside text boxes, which are
 * stories of their own.
 */
export function storyParagraphs(blocks: DocxBlock[]): { body: DocxParagraph[]; boxes: DocxParagraph[] } {
  const body: DocxParagraph[] = [];
  const boxes: DocxParagraph[] = [];
  const visitDrawing = (d: DocxDrawing | undefined) => {
    if (!d) return;
    if (d.kind === 'shape') walk(d.content, boxes);
    else if (d.kind === 'group') for (const item of d.items) visitDrawing(item.drawing);
  };
  const walk = (list: DocxBlock[], into: DocxParagraph[]) => {
    for (const b of list) {
      if (b.kind === 'paragraph') {
        into.push(b);
        for (const r of b.runs) visitDrawing(r.drawing);
        for (const a of b.anchors ?? []) visitDrawing(a.drawing);
      } else if (b.kind === 'table') {
        for (const row of b.rows) for (const cell of row) if (cell) walk(cell.content, into);
      }
    }
  };
  walk(blocks, body);
  return { body, boxes };
}

const WORD = /[\p{L}\p{N}_]/u;

/** Start of the word before `off` (skipping spaces first), as Ctrl+Backspace deletes. */
export function wordStartBefore(text: string, off: number): number {
  let i = off;
  while (i > 0 && /\s/.test(text[i - 1]!)) i--;
  if (i > 0 && !WORD.test(text[i - 1]!)) return i - 1;
  while (i > 0 && WORD.test(text[i - 1]!)) i--;
  return i;
}

/** End of the word after `off` (then its trailing spaces), as Ctrl+Delete deletes. */
export function wordEndAfter(text: string, off: number): number {
  let i = off;
  if (i < text.length && !WORD.test(text[i]!) && !/\s/.test(text[i]!)) return i + 1;
  while (i < text.length && WORD.test(text[i]!)) i++;
  while (i < text.length && /[ \t]/.test(text[i]!)) i++;
  return i;
}

/** The word around `off`, if the caret sits inside one (not at its edges). */
export function wordAround(text: string, off: number): [number, number] | undefined {
  if (off <= 0 || off >= text.length) return undefined;
  if (!WORD.test(text[off - 1]!) || !WORD.test(text[off]!)) return undefined;
  let s = off;
  while (s > 0 && WORD.test(text[s - 1]!)) s--;
  let e = off;
  while (e < text.length && WORD.test(text[e]!)) e++;
  return [s, e];
}

/** The previous grapheme-ish boundary: one code point back. */
export function prevCharBoundary(text: string, off: number): number {
  if (off <= 0) return 0;
  const c = text.charCodeAt(off - 1);
  return c >= 0xdc00 && c <= 0xdfff && off >= 2 ? off - 2 : off - 1;
}

export function nextCharBoundary(text: string, off: number): number {
  if (off >= text.length) return text.length;
  const c = text.charCodeAt(off);
  return c >= 0xd800 && c <= 0xdbff && off + 2 <= text.length ? off + 2 : off + 1;
}
