/**
 * Selection formatting — format-agnostic.
 *
 *
 * `document.execCommand` is deprecated and each browser invents its own markup
 * for it (`<b>`, `<font>`, inline styles), which reconciliation would then have
 * to reverse-engineer. Instead the toolbar wraps the selection in one span that
 * carries the intended {@link RunFormat} as data, and paints the equivalent CSS
 * so the change is visible immediately. Reconciliation reads the data back —
 * the CSS is only ever presentation.
 */
import { ptToPx } from '../units.js';
import { EDIT_ATTR } from './attrs.js';
import { mergeFormat, type ParaFormat, type RunFormat } from './format.js';

/** Resolves a key stamped on the DOM back to the model object it names. */
export type Resolver = (key: string | null | undefined) => object | undefined;

/** Reads the resolved formatting off a format's own run model. */
export type ReadRunFormat = (run: object) => RunFormat;

/** Reads the resolved paragraph formatting off a format's own paragraph model. */
export type ReadParaFormat = (para: object) => ParaFormat;

/** Paint a format onto an element so the user sees it before the commit. */
function paint(el: HTMLElement, f: RunFormat): void {
  if (f.bold !== undefined) el.style.fontWeight = f.bold ? '700' : '400';
  if (f.italic !== undefined) el.style.fontStyle = f.italic ? 'italic' : 'normal';
  if (f.sizePt !== undefined) el.style.fontSize = `${ptToPx(f.sizePt)}px`;
  if (f.colorHex !== undefined) el.style.color = `#${f.colorHex.replace(/^#/, '')}`;
  if (f.font !== undefined) el.style.fontFamily = `"${f.font}", Arial, Helvetica, sans-serif`;
  const decorations: string[] = [];
  if (f.underline) decorations.push('underline');
  if (f.strike) decorations.push('line-through');
  el.style.textDecoration = decorations.length ? decorations.join(' ') : '';
}

/** The editable text-body element a node sits inside, if any. */
export function bodyElementOf(node: Node | null, root: Element): HTMLElement | null {
  let el: Node | null = node;
  while (el && el !== root) {
    if (el.nodeType === 1 && (el as Element).hasAttribute(EDIT_ATTR.body)) return el as HTMLElement;
    el = el.parentNode;
  }
  return null;
}

/**
 * Wrap the current selection in a format marker. Returns the text-body element
 * that changed so the caller knows what to commit, or null if the selection is
 * empty or outside an editable body.
 *
 * **One marker per paragraph.** `extractContents` *moves* nodes, so a single
 * span spanning a paragraph boundary would pull the second paragraph's content
 * into the first — merging paragraphs, and losing text where the boundary fell.
 * Clamping the range to each paragraph in turn leaves the structure alone and
 * gives reconciliation exactly the paragraphs it started with.
 */
export function applyFormatToSelection(root: Element, format: RunFormat): HTMLElement | null {
  const sel = root.ownerDocument?.defaultView?.getSelection();
  if (!sel || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (range.collapsed) return null;

  // PPTX's editable unit is a whole text body; DOCX's is a single paragraph,
  // which carries no body marker — so fall back to the paragraph. The element
  // returned is what the caller commits, and both viewers accept either.
  const body =
    bodyElementOf(range.commonAncestorContainer, root) ??
    paraElementOf(range.commonAncestorContainer, root) ??
    paraElementOf(range.startContainer, root);
  if (!body) return null;

  const doc = root.ownerDocument!;
  const paras = paraElements(body).filter((p) => range.intersectsNode(p));
  const pieces = paras.length > 1 ? paras.map((p) => clampToParagraph(doc, range, p)) : [range];

  const wrappers: HTMLElement[] = [];
  for (const piece of pieces) {
    // A selection that merely touches a paragraph's edge covers none of its
    // text; wrapping that would leave an empty marker behind.
    if (piece.collapsed) continue;
    const wrapper = doc.createElement('span');
    wrapper.setAttribute(EDIT_ATTR.fmt, JSON.stringify(format));
    paint(wrapper, format);
    // extractContents may split spans mid-run; each half keeps its run key, so
    // both still resolve to the same source run and inherit the wrapper's format.
    wrapper.appendChild(piece.extractContents());
    piece.insertNode(wrapper);
    wrappers.push(wrapper);
  }
  if (wrappers.length === 0) return null;

  // Keep the same text selected so repeated toolbar clicks compound.
  const first = wrappers[0]!;
  const last = wrappers[wrappers.length - 1]!;
  const after = doc.createRange();
  after.setStartBefore(first.firstChild ?? first);
  after.setEndAfter(last.lastChild ?? last);
  sel.removeAllRanges();
  sel.addRange(after);

  return body;
}

/** `range` clipped to the part of it that lies inside one paragraph. */
function clampToParagraph(doc: Document, range: Range, para: Element): Range {
  const piece = doc.createRange();
  piece.selectNodeContents(para);
  // Whichever end of the selection is inside this paragraph wins; the other
  // stays at the paragraph's own edge.
  if (para.contains(range.startContainer) || para === range.startContainer) {
    piece.setStart(range.startContainer, range.startOffset);
  }
  if (para.contains(range.endContainer) || para === range.endContainer) {
    piece.setEnd(range.endContainer, range.endOffset);
  }
  return piece;
}

/** The editable paragraph element a node sits inside, if any. */
export function paraElementOf(node: Node | null, root: Element): HTMLElement | null {
  let el: Node | null = node;
  while (el && el !== root) {
    if (el.nodeType === 1 && (el as Element).hasAttribute(EDIT_ATTR.para)) return el as HTMLElement;
    el = el.parentNode;
  }
  return null;
}

/**
 * The paragraphs the selection covers, and the body they belong to.
 *
 * Unlike run formatting, a paragraph command works from a bare caret — there is
 * nothing to select, the paragraph the caret is in *is* the target — so this
 * deliberately does not require a range with content. A selection spanning two
 * text bodies is refused rather than partly applied: each body commits on its
 * own, and PowerPoint does not offer the gesture either.
 */
export function paragraphsInSelection(
  root: Element,
): { body: HTMLElement; paras: HTMLElement[] } | null {
  const sel = root.ownerDocument?.defaultView?.getSelection();
  if (!sel || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);

  const body = bodyElementOf(range.commonAncestorContainer, root);
  // Only the box the user has actually entered may be formatted; every other
  // body is addressable but read-only until entered.
  if (!body || body.getAttribute('contenteditable') !== 'true') return null;

  const all = Array.from(body.children).filter((el): el is HTMLElement =>
    el.hasAttribute(EDIT_ATTR.para),
  );
  const covered = all.filter((el) => range.intersectsNode(el));
  if (covered.length) return { body, paras: covered };

  // A collapsed caret can sit *between* paragraphs rather than inside one —
  // entering a box leaves it after the last child, which intersects nothing —
  // so fall back to the paragraph the position belongs to.
  const single =
    paraElementOf(range.startContainer, root) ??
    (range.startContainer === body ? paraBeforeOffset(body, range.startOffset, all) : null);
  return single ? { body, paras: [single] } : null;
}

/** The paragraph a child-offset into `body` falls in or just after. */
function paraBeforeOffset(
  body: HTMLElement,
  offset: number,
  paras: HTMLElement[],
): HTMLElement | null {
  const kids = Array.from(body.childNodes);
  const before = paras.filter((p) => kids.indexOf(p) < offset);
  return before[before.length - 1] ?? paras[0] ?? null;
}

/** Mark a paragraph with the formatting a commit should write into its `pPr`. */
export function stampParaFormat(el: HTMLElement, format: ParaFormat): void {
  el.setAttribute(EDIT_ATTR.paraFmt, JSON.stringify(format));
}

/**
 * The paragraph formatting in effect across the selection, for the toolbar.
 *
 * Where the selected paragraphs disagree the property is reported as
 * `undefined` — "mixed" — which is what stops a toolbar from claiming a state
 * that only some of the selection is in.
 */
export function paraFormatAtSelection(
  root: Element,
  resolve: Resolver,
  readParaFormat: ReadParaFormat,
): ParaFormat {
  const found = paragraphsInSelection(root);
  if (!found) return {};

  const formats = found.paras.map((el) => {
    const model = resolve(el.getAttribute(EDIT_ATTR.para));
    const resolved = model ? readParaFormat(model) : {};
    // A stamp that has not been committed yet is nearer than the model.
    const raw = el.getAttribute(EDIT_ATTR.paraFmt);
    if (!raw) return resolved;
    try {
      return mergeFormat(resolved, JSON.parse(raw) as ParaFormat);
    } catch {
      return resolved;
    }
  });

  const first = formats[0];
  if (!first) return {};
  const out: ParaFormat = { ...first };
  for (const other of formats.slice(1)) {
    for (const key of Object.keys(out) as Array<keyof ParaFormat>) {
      if (other[key] !== out[key]) delete out[key];
    }
  }
  return out;
}

/**
 * Where a caret or selection sits inside a text body, in terms the DOM cannot
 * lose: a paragraph index and character offsets within it.
 *
 * A commit re-renders the body from the committed XML, so every node the
 * selection pointed at is gone. Element keys are no help either — they are
 * reissued per render. Offsets survive because the *text* survives.
 */
export interface CaretAddress {
  /** Paragraph the selection starts in, by index within the body. */
  paraIndex: number;
  start: number;
  /** Paragraph it ends in — a selection may span several. */
  endParaIndex: number;
  end: number;
}

/** The paragraph elements of a text body, in order. */
function paraElements(bodyEl: Element): HTMLElement[] {
  return Array.from(bodyEl.children).filter((el): el is HTMLElement =>
    el.hasAttribute(EDIT_ATTR.para),
  );
}

/**
 * The editable text nodes of a paragraph, in order.
 *
 * Generated decoration — a bullet glyph, a list number — is marked
 * `contenteditable="false"` and is not content, so it must not count towards an
 * offset or the caret would drift by the width of the marker on every commit.
 */
function textNodesIn(root: Node): Text[] {
  const out: Text[] = [];
  const walk = (node: Node): void => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === 3) out.push(child as Text);
      else if (child.nodeType === 1) {
        if ((child as Element).getAttribute('contenteditable') === 'false') continue;
        walk(child);
      }
    }
  };
  walk(root);
  return out;
}

/**
 * Character offset of a DOM position within `para`, counting editable text only.
 *
 * Measured by cloning the range from the paragraph's start to the position: the
 * container may be a text node, but it may equally be an element — the toolbar
 * leaves the selection wrapping a whole span — and only the clone handles both
 * without a second traversal.
 */
function offsetIn(para: Element, container: Node, offset: number): number {
  const doc = para.ownerDocument;
  if (!doc) return 0;
  const range = doc.createRange();
  range.setStart(para, 0);
  try {
    range.setEnd(container, offset);
  } catch {
    // A position outside the paragraph: the caller's clamping takes over.
    return 0;
  }
  return textNodesIn(range.cloneContents()).reduce((n, t) => n + t.data.length, 0);
}

/** The DOM position a character offset names, clamped into the paragraph. */
/**
 * The DOM position a character offset names, clamped into the paragraph.
 *
 * A position that lands exactly on a run boundary belongs to the *start of the
 * next* run, not the end of the previous one. Both look identical on screen,
 * but the toolbar reads its state from the node the selection starts in — so
 * getting this wrong makes a formatting command report the formatting of the
 * text beside it, which looks like the command did nothing.
 */
function positionAt(para: Element, offset: number): { node: Node; offset: number } {
  const nodes = textNodesIn(para);
  let remaining = offset;
  for (let i = 0; i < nodes.length; i++) {
    const node = nodes[i]!;
    if (remaining < node.data.length || i === nodes.length - 1) {
      return { node, offset: Math.min(Math.max(remaining, 0), node.data.length) };
    }
    remaining -= node.data.length;
  }
  return { node: para, offset: 0 };
}

/** Read the caret out of a text body so it can be put back after a commit. */
export function readCaretAddress(bodyEl: HTMLElement): CaretAddress | null {
  const sel = bodyEl.ownerDocument.defaultView?.getSelection();
  if (!sel || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (!bodyEl.contains(range.startContainer)) return null;

  const paras = paraElements(bodyEl);
  const holds = (p: Element, n: Node) => p === n || p.contains(n);
  const paraIndex = paras.findIndex((p) => holds(p, range.startContainer));
  const para = paras[paraIndex];
  if (!para) return null;

  const endParaIndex = paras.findIndex((p) => holds(p, range.endContainer));
  const endPara = paras[endParaIndex];
  const start = offsetIn(para, range.startContainer, range.startOffset);
  return endPara
    ? {
        paraIndex,
        start,
        endParaIndex,
        end: offsetIn(endPara, range.endContainer, range.endOffset),
      }
    : { paraIndex, start, endParaIndex: paraIndex, end: start };
}

/**
 * Put a caret back after a re-render. Clamps rather than failing: landing at
 * the end of a paragraph is fine, having no selection at all is not — the
 * toolbar's state is driven by `selectionchange`.
 */
export function restoreCaretAddress(bodyEl: HTMLElement, at: CaretAddress): void {
  const doc = bodyEl.ownerDocument;
  const sel = doc.defaultView?.getSelection();
  if (!sel) return;
  const paras = paraElements(bodyEl);
  const para = paras[Math.min(at.paraIndex, paras.length - 1)];
  const endPara = paras[Math.min(at.endParaIndex, paras.length - 1)];
  if (!para || !endPara) return;

  const from = positionAt(para, at.start);
  const to = positionAt(endPara, para === endPara ? Math.max(at.start, at.end) : at.end);
  const range = doc.createRange();
  range.setStart(from.node, from.offset);
  try {
    range.setEnd(to.node, to.offset);
  } catch {
    // Paragraphs can vanish under an edit (a merge); a caret is better than
    // nothing, since the toolbar's state is driven by having a selection.
    range.setEnd(from.node, from.offset);
  }
  sel.removeAllRanges();
  sel.addRange(range);
}

/**
 * The formatting in effect at the caret, for the toolbar's active state.
 * `readRunFormat` reads the *resolved* model run rather than its raw XML
 * properties, so everything the run inherits is reflected too.
 */
export function formatAtSelection(
  root: Element,
  resolve: Resolver,
  readRunFormat: ReadRunFormat,
): RunFormat {
  const sel = root.ownerDocument?.defaultView?.getSelection();
  if (!sel || sel.rangeCount === 0) return {};

  let node: Node | null = sel.getRangeAt(0).startContainer;
  let pending: RunFormat = {};
  const overrides: RunFormat[] = [];

  while (node && node !== root) {
    if (node.nodeType === 1) {
      const el = node as Element;
      const raw = el.getAttribute(EDIT_ATTR.fmt);
      if (raw) {
        try {
          overrides.unshift(JSON.parse(raw) as RunFormat);
        } catch {
          /* a malformed marker just contributes nothing */
        }
      }
      const runKey = el.getAttribute(EDIT_ATTR.run);
      if (runKey) {
        const run = resolve(runKey);
        if (run) pending = readRunFormat(run);
        break;
      }
    }
    node = node.parentNode;
  }

  // Nearer markers win over the run's own resolved properties.
  return overrides.reduce<RunFormat>((acc, o) => mergeFormat(acc, o), pending);
}
