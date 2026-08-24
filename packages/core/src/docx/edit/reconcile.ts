/**
 * DOM -> segment list, for DOCX.
 *
 * Same strategy as the PPTX reconciler: let contentEditable do whatever it
 * likes and read the result back, so typing, Enter, Backspace, paste and
 * toolbar formatting all arrive through one code path.
 *
 * The DOCX difference is the unit. A PPTX text body is a self-contained box, so
 * one editable region holds many paragraphs. A DOCX body is one long flow, so
 * each paragraph is its own editable region — and pressing Enter has to be
 * detected as "this region now contains more than one block".
 */
import type { DocxParagraph, DocxRun } from '../model.js';
import { EDIT_ATTR } from '../../oxml/edit/attrs.js';
import { mergeFormat, type RunFormat } from '../../oxml/edit/format.js';
import type { Resolver } from '../../oxml/edit/selection.js';
import type { DocxParaEdit, DocxSegment } from './xmlWrite.js';

/** The zero-width space the renderer uses to give empty paragraphs height. */
const ZWSP = '​';

function isElement(n: Node): n is Element {
  return n.nodeType === 1;
}

function isText(n: Node): n is Text {
  return n.nodeType === 3;
}

function readFmtAttr(el: Element): RunFormat | undefined {
  const raw = el.getAttribute(EDIT_ATTR.fmt);
  if (!raw) return undefined;
  try {
    return JSON.parse(raw) as RunFormat;
  } catch {
    return undefined;
  }
}

interface WalkState {
  src?: DocxRun;
  format?: RunFormat;
}

/** Pending break/tab, carried onto the next segment that has text. */
interface Pending {
  breakBefore?: boolean;
  tabBefore?: boolean;
}

function walk(
  nodes: Iterable<Node>,
  state: WalkState,
  out: DocxSegment[],
  pending: Pending,
  resolve: Resolver,
): void {
  for (const node of nodes) {
    if (isText(node)) {
      const text = node.data;
      if (!text) continue;
      const seg: DocxSegment = { text };
      if (state.src) seg.src = state.src;
      if (state.format) seg.format = state.format;
      if (pending.breakBefore) {
        seg.breakBefore = true;
        pending.breakBefore = false;
      }
      if (pending.tabBefore) {
        seg.tabBefore = true;
        pending.tabBefore = false;
      }
      out.push(seg);
      continue;
    }
    if (!isElement(node)) continue;

    // List markers and other generated decoration are marked uneditable and
    // carry no run key — they are not content and must not become runs.
    const runKey = node.getAttribute(EDIT_ATTR.run);
    if (!runKey && (node as HTMLElement).contentEditable === 'false') continue;

    if (node.tagName === 'BR') {
      pending.breakBefore = true;
      continue;
    }

    const resolved = runKey ? (resolve(runKey) as DocxRun | undefined) : undefined;
    const fmt = readFmtAttr(node);
    const next: WalkState = {};
    const src = resolved ?? state.src;
    if (src) next.src = src;
    if (fmt || state.format) next.format = mergeFormat(state.format, fmt);
    walk(Array.from(node.childNodes), next, out, pending, resolve);
  }
}

/** Two segments can be one run when they share a source and formatting. */
function sameRun(a: DocxSegment, b: DocxSegment): boolean {
  if (b.breakBefore || b.tabBefore) return false;
  if (a.src !== b.src) return false;
  return JSON.stringify(a.format ?? {}) === JSON.stringify(b.format ?? {});
}

/**
 * Give unsourced text a source to inherit properties from — the run to its
 * left, falling back to the one on its right. This is what makes text typed at
 * the very start of a paragraph pick up that paragraph's formatting rather than
 * landing as a bare run with no rPr.
 */
function inheritSources(segments: DocxSegment[]): DocxSegment[] {
  let last: DocxRun | undefined;
  const forward = segments.map((seg) => {
    if (seg.src) last = seg.src;
    else if (last) return { ...seg, src: last };
    return seg;
  });
  let nextSrc: DocxRun | undefined;
  for (let i = forward.length - 1; i >= 0; i--) {
    const seg = forward[i]!;
    if (seg.src) nextSrc = seg.src;
    else if (nextSrc) forward[i] = { ...seg, src: nextSrc };
  }
  return forward;
}

function tidy(segments: DocxSegment[]): DocxSegment[] {
  const cleaned = segments
    // The zero-width space only existed to give an empty line height.
    .map((s) => ({ ...s, text: s.text.split(ZWSP).join('') }))
    .filter((s) => s.text.length > 0 || s.breakBefore || s.tabBefore);

  const sourced = inheritSources(cleaned);

  // Coalesce so a retyped paragraph doesn't explode into one run per
  // keystroke-boundary the browser happened to create.
  const merged: DocxSegment[] = [];
  for (const seg of sourced) {
    const prev = merged[merged.length - 1];
    if (prev && sameRun(prev, seg)) prev.text += seg.text;
    else merged.push({ ...seg });
  }
  return merged;
}

/**
 * A trailing `<br>` is the filler browsers put in an empty editable block, not
 * content. Our renderer only emits `<br>` for a real `<w:br/>`, which is always
 * followed by its text, so a trailing one is never ours.
 */
function withoutTrailingFiller(el: Element): Node[] {
  const nodes = Array.from(el.childNodes);
  while (nodes.length) {
    const last = nodes[nodes.length - 1]!;
    if (isElement(last) && last.tagName === 'BR') nodes.pop();
    else break;
  }
  return nodes;
}

function segmentsOf(nodes: Iterable<Node>, resolve: Resolver): DocxSegment[] {
  const raw: DocxSegment[] = [];
  walk(nodes, {}, raw, {}, resolve);
  return tidy(raw);
}

/**
 * Read one edited paragraph element back into the paragraph list the XML writer
 * consumes.
 *
 * Pressing Enter makes the browser split the editable div's content into two
 * blocks (or clone the div outright). Any block-level child therefore means the
 * paragraph became several, and each becomes a ParaEdit sharing the original's
 * source — which is exactly how the writer clones `<w:pPr>` onto the new
 * paragraphs.
 */
export function reconcileParagraph(el: Element, resolve: Resolver): DocxParaEdit[] {
  const src = resolve(el.getAttribute(EDIT_ATTR.para)) as DocxParagraph | undefined;
  const base = src ? { src } : {};

  const nodes = withoutTrailingFiller(el);
  const blocks = nodes.filter(
    (n) => isElement(n) && (n.tagName === 'DIV' || n.tagName === 'P'),
  ) as Element[];

  if (blocks.length === 0) {
    return [{ ...base, segments: segmentsOf(nodes, resolve) }];
  }

  // Mixed inline content and blocks: the loose inline content is the first
  // paragraph, each block is another.
  const out: DocxParaEdit[] = [];
  const loose = nodes.filter((n) => !blocks.includes(n as Element));
  if (loose.length) {
    const segments = segmentsOf(loose, resolve);
    if (segments.length) out.push({ ...base, segments });
  }
  for (const block of blocks) {
    out.push({ ...base, segments: segmentsOf(withoutTrailingFiller(block), resolve) });
  }
  return out.length ? out : [{ ...base, segments: [] }];
}
