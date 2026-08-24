/**
 * DOM -> segment list.
 *
 * Rather than intercepting keystrokes, the editor lets contentEditable do
 * whatever it likes and then reads the result back. Typing, Enter, Backspace,
 * paste and toolbar formatting all arrive here as "the DOM now looks like
 * this", and all leave as the same ParaEdit list — so there is one code path to
 * get right instead of one per interaction.
 *
 * The walk carries the source run and any formatting override *down* the tree,
 * so text the browser re-parented (or that the user typed at a run boundary)
 * still inherits the properties of the run it visually belongs to.
 */
import type { Paragraph, TextRun } from '../model.js';
import { EDIT_ATTR } from '../text/render.js';
import { mergeFormat, type RunFormat } from './format.js';
import type { ParaEdit, Segment } from './xmlWrite.js';

/** The zero-width space the renderer uses to give empty paragraphs height. */
const ZWSP = '​';

/** Resolves a key stamped on the DOM back to the model object it names. */
export type Resolver = (key: string | null | undefined) => object | undefined;

function isElement(n: Node): n is Element {
  return n.nodeType === 1;
}

function isText(n: Node): n is Text {
  return n.nodeType === 3;
}

/** Parse the JSON format payload the toolbar stamps onto a wrapper span. */
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
  src?: TextRun;
  format?: RunFormat;
}

function segment(text: string, state: WalkState, isBreak = false): Segment {
  const seg: Segment = { text };
  if (isBreak) seg.isBreak = true;
  if (state.src) seg.src = state.src;
  if (state.format) seg.format = state.format;
  return seg;
}

/** Walk a list of sibling nodes, carrying source/format inheritance downward. */
function walk(nodes: Iterable<Node>, state: WalkState, out: Segment[], resolve: Resolver): void {
  for (const node of nodes) {
    if (isText(node)) {
      if (node.data) out.push(segment(node.data, state));
      continue;
    }
    if (!isElement(node)) continue;

    // Bullets and other generated decoration are marked uneditable and carry no
    // run key — they are not content and must not become runs.
    const runKey = node.getAttribute(EDIT_ATTR.run);
    if (!runKey && (node as HTMLElement).contentEditable === 'false') continue;

    if (node.tagName === 'BR') {
      out.push(segment('\n', state, true));
      continue;
    }

    const resolved = runKey ? (resolve(runKey) as TextRun | undefined) : undefined;
    const fmt = readFmtAttr(node);
    const next: WalkState = {};
    const src = resolved ?? state.src;
    if (src) next.src = src;
    if (fmt || state.format) next.format = mergeFormat(state.format, fmt);
    walk(Array.from(node.childNodes), next, out, resolve);
  }
}

/** Two segments can be one run when they share a source and formatting. */
function sameRun(a: Segment, b: Segment): boolean {
  if (a.isBreak || b.isBreak) return false;
  if (a.src !== b.src) return false;
  return JSON.stringify(a.format ?? {}) === JSON.stringify(b.format ?? {});
}

/**
 * Split text segments on newlines into explicit break segments. Paragraphs
 * render with `white-space: pre-wrap`, so a model `<a:br>` reaches the DOM as a
 * literal newline; the browser may also leave a real `<br>`. Both normalize
 * here so the writer only ever sees explicit breaks.
 */
function expandBreaks(segments: Segment[]): Segment[] {
  const out: Segment[] = [];
  for (const seg of segments) {
    if (seg.isBreak || !seg.text.includes('\n')) {
      out.push(seg);
      continue;
    }
    const parts = seg.text.split('\n');
    parts.forEach((part, i) => {
      if (i > 0) out.push({ text: '\n', isBreak: true, ...(seg.src ? { src: seg.src } : {}) });
      if (part) out.push({ ...seg, text: part });
    });
  }
  return out;
}

/**
 * Give unsourced text a source to inherit properties from — the run to its
 * left, falling back to the one on its right. This is what makes text typed at
 * the very start of a paragraph pick up that paragraph's formatting instead of
 * landing as a bare run with no rPr.
 */
function inheritSources(segments: Segment[]): Segment[] {
  let last: TextRun | undefined;
  const forward = segments.map((seg) => {
    if (seg.src) last = seg.src;
    else if (last) return { ...seg, src: last };
    return seg;
  });
  let nextSrc: TextRun | undefined;
  for (let i = forward.length - 1; i >= 0; i--) {
    const seg = forward[i]!;
    if (seg.src) nextSrc = seg.src;
    else if (nextSrc) forward[i] = { ...seg, src: nextSrc };
  }
  return forward;
}

function tidy(segments: Segment[]): Segment[] {
  const expanded = expandBreaks(segments)
    // The zero-width space only existed to give an empty line height.
    .map((s) => (s.isBreak ? s : { ...s, text: s.text.split(ZWSP).join('') }))
    .filter((s) => s.isBreak || s.text.length > 0);

  const sourced = inheritSources(expanded);

  // Coalesce so a paragraph the user retyped doesn't explode into one run per
  // keystroke-boundary the browser happened to create.
  const merged: Segment[] = [];
  for (const seg of sourced) {
    const prev = merged[merged.length - 1];
    if (prev && sameRun(prev, seg)) prev.text += seg.text;
    else merged.push({ ...seg });
  }
  return merged;
}

/**
 * A trailing `<br>` is the filler Chrome puts in an empty editable block, not
 * content. Our renderer never emits `<br>`, so any we see is browser-generated.
 */
function dropTrailingFiller(el: Element): Node[] {
  const nodes = Array.from(el.childNodes);
  while (nodes.length) {
    const last = nodes[nodes.length - 1]!;
    if (isElement(last) && last.tagName === 'BR') nodes.pop();
    else break;
  }
  return nodes;
}

function paragraphOf(el: Element, resolve: Resolver): ParaEdit {
  const src = resolve(el.getAttribute(EDIT_ATTR.para)) as Paragraph | undefined;
  const raw: Segment[] = [];
  // Walk the retained children rather than the element, so trailing filler is
  // excluded without mutating the live DOM.
  walk(dropTrailingFiller(el), {}, raw, resolve);
  return { ...(src ? { src } : {}), segments: tidy(raw) };
}

/**
 * Read an edited text-body element back into the paragraph list the XML writer
 * consumes. Direct element children are paragraphs; a duplicated paragraph key
 * (what pressing Enter produces) simply yields two ParaEdits sharing a source,
 * which is exactly how the writer clones `<a:pPr>` onto the new paragraph.
 */
export function reconcileTextBody(boxEl: Element, resolve: Resolver): ParaEdit[] {
  const paras: ParaEdit[] = [];
  let loose: Segment[] = [];

  const flushLoose = () => {
    if (loose.length) {
      paras.push({ segments: tidy(loose) });
      loose = [];
    }
  };

  for (const node of Array.from(boxEl.childNodes)) {
    if (isElement(node) && node.hasAttribute(EDIT_ATTR.para)) {
      flushLoose();
      paras.push(paragraphOf(node, resolve));
      continue;
    }
    if (isElement(node) && (node.tagName === 'DIV' || node.tagName === 'P')) {
      // A block the browser created without copying our marker.
      flushLoose();
      paras.push(paragraphOf(node, resolve));
      continue;
    }
    // Bare text or inline markup promoted to the box — collect into a paragraph.
    walk([node], {}, loose, resolve);
  }
  flushLoose();

  return paras.length ? paras : [{ segments: [] }];
}
