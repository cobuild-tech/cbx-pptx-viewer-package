/**
 * Word-style editing on the laid-out pages.
 *
 * The whole page stack is one editable surface, so the caret moves and a
 * selection extends across paragraphs, tables and page breaks exactly as in
 * Word's Print Layout. The browser is never allowed to change the DOM itself:
 * every `beforeinput` is cancelled and turned into a command on the
 * {@link DocxEditSession}, which edits the XML; the pages are then re-laid out
 * and the caret put back where the command says it belongs. What is on screen
 * is therefore always the document, never the browser's improvisation.
 *
 * DOM ↔ position mapping goes through the run stamps the renderer leaves
 * (`data-cbx-run` + `data-cbx-part`) and the layout-copy origins the paginator
 * records, so a caret in the second half of a paragraph split across two
 * pages still addresses the right characters.
 */
import { EDIT_ATTR } from '../../oxml/edit/attrs.js';
import type { RunFormat } from '../../oxml/edit/format.js';
import { mergeFormat } from '../../oxml/edit/format.js';
import type { XmlNode } from '../../oxml/xml.js';
import type { DocxEditContext } from './context.js';
import { readRunFormat } from './format.js';
import { caret, isCollapsed, type DocPos, type DocSel, type DocxEditSession, type DocxParaFormat } from './session.js';
import { wordAround, type ParaMap } from './story.js';

export interface DocxEditControllerOptions {
  /** The editable page stack. */
  host: HTMLElement;
  /** The scrolling viewport around it. */
  scroller: HTMLElement;
  session: DocxEditSession;
  ctx: DocxEditContext;
  /** Re-lay out and redraw the pages from the current document. */
  render(): void;
  /** The caret's character and paragraph formatting, for a toolbar. */
  onSelectionChange?(format: RunFormat, para: DocxParaFormat): void;
}

interface Piece {
  el: Element;
  start: number;
  end: number;
  text: boolean;
}

export class DocxEditController {
  private readonly o: DocxEditControllerOptions;
  private readonly listeners: Array<[EventTarget, string, EventListener, boolean?]> = [];
  /** `<w:p>` -> its rendered fragments (one per page it spans), rebuilt lazily after each render. */
  private fragments: Map<XmlNode, HTMLElement[]> | null = null;
  /** Formatting chosen with nothing selected, applied to what is typed next at this spot. */
  private pending: { at: DocPos; format: RunFormat } | undefined;
  private composing: DocSel | undefined;
  /** Set while the controller itself moves the selection. */
  private settingSelection = false;

  constructor(options: DocxEditControllerOptions) {
    this.o = options;
    const host = options.host;
    host.setAttribute('contenteditable', 'true');
    host.spellcheck = false;
    host.style.outline = 'none';
    host.style.caretColor = '#000';
    host.style.cursor = 'text';
    const doc = host.ownerDocument;
    try {
      // Firefox's table and image grips would edit the DOM behind our back.
      doc.execCommand?.('enableObjectResizing', false, 'false');
      doc.execCommand?.('enableInlineTableEditing', false, 'false');
    } catch {
      /* not supported */
    }
    this.on(host, 'beforeinput', (e) => this.onBeforeInput(e as InputEvent));
    this.on(host, 'keydown', (e) => this.onKeyDown(e as KeyboardEvent));
    this.on(host, 'paste', (e) => this.onPaste(e as ClipboardEvent));
    this.on(host, 'cut', (e) => this.onCut(e as ClipboardEvent));
    this.on(host, 'dragstart', (e) => e.preventDefault());
    this.on(host, 'drop', (e) => e.preventDefault());
    this.on(host, 'compositionstart', () => {
      this.composing = this.readSelection();
    });
    this.on(host, 'compositionend', (e) => this.onCompositionEnd(e as CompositionEvent));
    // Anything that still slipped through (autocorrect, an extension) is
    // discarded by redrawing the document as it really is.
    this.on(host, 'input', () => {
      if (!this.composing) this.redraw(this.readSelection());
    });
    this.on(host, 'click', (e) => this.onClick(e as MouseEvent));
    this.on(doc, 'selectionchange', () => this.onSelectionChange());
  }

  private on(target: EventTarget, type: string, fn: (e: Event) => void, capture = false): void {
    target.addEventListener(type, fn, capture);
    this.listeners.push([target, type, fn, capture]);
  }

  destroy(): void {
    for (const [t, type, fn, capture] of this.listeners) t.removeEventListener(type, fn, capture);
    this.o.host.removeAttribute('contenteditable');
    this.o.host.style.caretColor = '';
    this.o.host.style.cursor = '';
  }

  /** The pages were redrawn: forget the DOM index. */
  invalidate(): void {
    this.fragments = null;
  }

  // ─── Commands (toolbar entry points) ───────────────────────────────────────

  /** Apply character formatting: to the selection, the word at the caret, or what is typed next. */
  format(format: RunFormat): void {
    const sel = this.readSelection() ?? this.lastSel;
    if (!sel) return;
    if (isCollapsed(sel)) {
      const pm = this.o.session.map(sel.focus.p);
      const word = pm ? wordAround(pm.text, sel.focus.off) : undefined;
      if (word) {
        const range = { anchor: { p: sel.focus.p, off: word[0] }, focus: { p: sel.focus.p, off: word[1] } };
        const done = this.o.session.formatSelection(range, format);
        if (done) this.redraw(sel);
        return;
      }
      // Word's behaviour with nothing selected: the format waits for typing.
      this.pending = { at: sel.focus, format: mergeFormat(this.pending && samePos(this.pending.at, sel.focus) ? this.pending.format : undefined, format) };
      this.report(sel);
      this.o.host.focus();
      this.setSelection(sel);
      return;
    }
    this.apply(this.o.session.formatSelection(sel, format), sel);
  }

  paragraphFormat(format: DocxParaFormat): void {
    const sel = this.readSelection() ?? this.lastSel;
    if (sel) this.apply(this.o.session.setParagraphFormat(sel, format), sel);
  }

  toggleList(kind: 'bullet' | 'number'): void {
    const sel = this.readSelection() ?? this.lastSel;
    if (sel) this.apply(this.o.session.toggleList(sel, kind), sel);
  }

  indent(delta: 1 | -1): void {
    const sel = this.readSelection() ?? this.lastSel;
    if (sel) this.apply(this.o.session.indent(sel, delta), sel);
  }

  undo(): void {
    // A restore replaces the whole XML tree: every stamp on screen is stale.
    if (this.o.session.undo()) {
      this.o.ctx.forgetNodes();
      this.redraw(this.o.session.lastSelection);
    }
  }

  redo(): void {
    if (this.o.session.redo()) {
      this.o.ctx.forgetNodes();
      this.redraw(this.o.session.lastSelection);
    }
  }

  /** Put the caret at the start of the document (when editing begins). */
  focusStart(): void {
    const range = this.o.session.documentRange();
    if (!range) return;
    this.o.host.focus({ preventScroll: true });
    this.setSelection(caret(range.anchor.p, 0));
  }

  // ─── Input ─────────────────────────────────────────────────────────────────

  private onBeforeInput(e: InputEvent): void {
    const type = e.inputType;
    // IME composition cannot be cancelled; it is applied when it ends.
    if (type === 'insertCompositionText' || type === 'deleteCompositionText' || type === 'insertFromComposition') return;
    e.preventDefault();
    if (this.composing) return;
    const sel = this.readSelection();
    if (!sel) return;
    const s = this.o.session;
    switch (type) {
      case 'insertText':
      case 'insertReplacementText': {
        const text = e.data ?? e.dataTransfer?.getData('text/plain') ?? '';
        this.typeText(sel, text);
        return;
      }
      case 'insertLineBreak':
        this.typeText(sel, '\n');
        return;
      case 'insertParagraph':
        this.apply(s.insertParagraph(sel), sel);
        return;
      case 'insertFromPaste':
      case 'insertFromPasteAsQuotation': {
        const text = e.dataTransfer?.getData('text/plain');
        if (text) this.apply(s.paste(sel, text), sel);
        return;
      }
      case 'deleteContentBackward':
        this.apply(s.deleteBackward(sel, 'char'), sel);
        return;
      case 'deleteContentForward':
        this.apply(s.deleteForward(sel, 'char'), sel);
        return;
      case 'deleteWordBackward':
        this.apply(s.deleteBackward(sel, 'word'), sel);
        return;
      case 'deleteWordForward':
        this.apply(s.deleteForward(sel, 'word'), sel);
        return;
      case 'deleteSoftLineBackward':
      case 'deleteHardLineBackward':
      case 'deleteSoftLineForward':
      case 'deleteHardLineForward':
      case 'deleteContent':
      case 'deleteByCut': {
        // The browser knows where the visual line starts; take its range.
        const target = isCollapsed(sel) ? this.targetRange(e) : sel;
        if (target) this.apply(s.deleteSelection(target), sel);
        return;
      }
      case 'formatBold':
        return this.toggleFormat(sel, 'bold');
      case 'formatItalic':
        return this.toggleFormat(sel, 'italic');
      case 'formatUnderline':
        return this.toggleFormat(sel, 'underline');
      case 'formatStrikeThrough':
        return this.toggleFormat(sel, 'strike');
      case 'historyUndo':
        return this.undo();
      case 'historyRedo':
        return this.redo();
      default:
        return;
    }
  }

  private typeText(sel: DocSel, text: string): void {
    if (!text) return;
    const pending = this.pending && isCollapsed(sel) && samePos(this.pending.at, sel.focus) ? this.pending.format : undefined;
    const result = this.o.session.insertText(sel, text, pending);
    if (result && pending) this.pending = { at: result.focus, format: pending };
    this.apply(result, sel);
  }

  private onCompositionEnd(e: CompositionEvent): void {
    const sel = this.composing;
    this.composing = undefined;
    if (!sel) return;
    const result = e.data ? this.o.session.insertText(sel, e.data) : undefined;
    this.redraw(result ?? sel);
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (e.isComposing) return;
    const mod = e.metaKey || e.ctrlKey;
    const key = e.key.toLowerCase();
    const sel = this.readSelection();
    if (mod && !e.altKey) {
      const run = (fn: () => void) => {
        e.preventDefault();
        fn();
      };
      if (key === 'z' && !e.shiftKey) return run(() => this.undo());
      if ((key === 'z' && e.shiftKey) || key === 'y') return run(() => this.redo());
      if (key === 'a') {
        return run(() => {
          const all = this.o.session.documentRange();
          if (all) this.setSelection(all);
        });
      }
      if (!sel) return;
      if (key === 'b') return run(() => this.toggleFormat(sel, 'bold'));
      if (key === 'i') return run(() => this.toggleFormat(sel, 'italic'));
      if (key === 'u') return run(() => this.toggleFormat(sel, 'underline'));
      const align = ({ l: 'left', e: 'center', r: 'right', j: 'justify' } as const)[key as 'l' | 'e' | 'r' | 'j'];
      if (align && !e.shiftKey) return run(() => this.paragraphFormat({ align }));
      return;
    }
    if (e.key === 'Tab' && sel) {
      e.preventDefault();
      this.onTab(sel, e.shiftKey);
    }
  }

  /**
   * Tab as in Word: in a table it moves to the next cell (Shift+Tab the
   * previous); at the start of a list item it demotes the item a level;
   * otherwise it types a tab.
   */
  private onTab(sel: DocSel, back: boolean): void {
    const frag = this.fragmentsOf(sel.focus.p)[0];
    const cell = frag?.closest('td');
    if (cell) {
      const all = [...this.o.host.querySelectorAll<HTMLElement>(`[${EDIT_ATTR.para}]`)];
      const i = all.indexOf(frag!);
      const step = back ? -1 : 1;
      for (let k = i + step; k >= 0 && k < all.length; k += step) {
        const td = all[k]!.closest('td');
        if (td && td !== cell) {
          const p = this.nodeOfElement(all[k]!);
          if (p) this.setSelection(caret(p, 0));
          return;
        }
      }
      return;
    }
    const s = this.o.session;
    if (isCollapsed(sel) && sel.focus.off === 0 && s.listKindOf(sel.focus.p)) {
      this.apply(s.indent(sel, back ? -1 : 1), sel);
      return;
    }
    if (back) return;
    this.typeText(sel, '\t');
  }

  private toggleFormat(sel: DocSel, key: 'bold' | 'italic' | 'underline' | 'strike'): void {
    const current = mergeFormat(this.formatAt(sel), this.pending?.format);
    this.format({ [key]: !current[key] });
  }

  private onPaste(e: ClipboardEvent): void {
    e.preventDefault();
    const sel = this.readSelection();
    const text = e.clipboardData?.getData('text/plain');
    if (sel && text) this.apply(this.o.session.paste(sel, text), sel);
  }

  private onCut(e: ClipboardEvent): void {
    const sel = this.readSelection();
    const range = this.o.host.ownerDocument.getSelection();
    if (!sel || isCollapsed(sel) || !range || !e.clipboardData) return;
    e.preventDefault();
    e.clipboardData.setData('text/plain', range.toString());
    const box = this.o.host.ownerDocument.createElement('div');
    box.appendChild(range.getRangeAt(0).cloneContents());
    e.clipboardData.setData('text/html', box.innerHTML);
    this.apply(this.o.session.deleteSelection(sel), sel);
  }

  /** Ctrl/Cmd+click follows a hyperlink, as in Word; a plain click places the caret. */
  private onClick(e: MouseEvent): void {
    const a = (e.target as Element | null)?.closest?.('a[href]') as HTMLAnchorElement | null;
    if (!a) return;
    e.preventDefault();
    if (e.ctrlKey || e.metaKey) window.open(a.href, '_blank', 'noopener');
  }

  private targetRange(e: InputEvent): DocSel | undefined {
    const r = e.getTargetRanges?.()[0];
    if (!r) return undefined;
    const anchor = this.domToPos(r.startContainer, r.startOffset);
    const focus = this.domToPos(r.endContainer, r.endOffset);
    return anchor && focus ? { anchor, focus } : undefined;
  }

  // ─── Applying results ──────────────────────────────────────────────────────

  private lastSel: DocSel | undefined;

  private apply(result: DocSel | undefined, before: DocSel): void {
    this.redraw(result ?? before);
  }

  /** Lay the document out again and put the selection back. */
  private redraw(sel: DocSel | undefined): void {
    this.o.render();
    this.fragments = null;
    // A toolbar command (a <select> above all) may have taken focus.
    const active = this.o.host.ownerDocument.activeElement;
    if (!active || !this.o.host.contains(active)) this.o.host.focus({ preventScroll: true });
    if (sel) {
      this.setSelection(sel);
      this.revealCaret();
    }
    this.report(sel);
  }

  /** Scroll just enough to keep the caret in view, as Word does while typing. */
  private revealCaret(): void {
    const s = this.o.host.ownerDocument.getSelection();
    if (!s || s.rangeCount === 0) return;
    const r = s.getRangeAt(0);
    let rect = typeof r.getBoundingClientRect === 'function' ? r.getBoundingClientRect() : undefined;
    if (!rect?.height) {
      const el = r.startContainer.nodeType === 1 ? (r.startContainer as Element) : r.startContainer.parentElement;
      if (!el) return;
      rect = el.getBoundingClientRect();
    }
    const view = this.o.scroller.getBoundingClientRect();
    const margin = 40;
    if (rect.bottom > view.bottom - margin) this.o.scroller.scrollTop += rect.bottom - view.bottom + margin;
    else if (rect.top < view.top + margin) this.o.scroller.scrollTop -= view.top + margin - rect.top;
  }

  private onSelectionChange(): void {
    if (this.settingSelection || this.composing) return;
    const sel = this.readSelection();
    if (!sel) return;
    if (this.pending && !(isCollapsed(sel) && samePos(this.pending.at, sel.focus))) this.pending = undefined;
    this.lastSel = sel;
    this.report(sel);
  }

  private report(sel: DocSel | undefined): void {
    if (!sel || !this.o.onSelectionChange) return;
    const format = mergeFormat(this.formatAt(sel), this.pending?.format);
    this.o.onSelectionChange(format, this.o.session.paraFormatAt(sel));
  }

  /** The formatting of the character before the caret (or of the selection's first). */
  private formatAt(sel: DocSel): RunFormat {
    const [a, b] = this.o.session.ordered(sel);
    const pm = this.o.session.map(a.p);
    if (!pm) return {};
    const collapsed = isCollapsed(sel);
    const segs = pm.segs.filter((s) => s.kind === 'text' || s.kind === 'obj');
    const pick =
      (collapsed
        ? [...segs].reverse().find((s) => s.start < a.off && a.off <= s.start + s.len) ?? segs.find((s) => s.start >= a.off)
        : segs.find((s) => s.start + s.len > a.off)) ?? segs[segs.length - 1];
    if (!pick) {
      const para = pm.para;
      const out: RunFormat = {};
      if (para.baseFontFamily) out.font = para.baseFontFamily;
      if (para.baseFontSizePt) out.sizePt = para.baseFontSizePt;
      if (para.baseBold) out.bold = true;
      if (para.baseItalic) out.italic = true;
      return out;
    }
    const base = readRunFormat(pick.run);
    if (collapsed || a.p !== b.p) return base;
    // Over a selection, a toggle is on only when every selected run has it.
    for (const s of segs) {
      if (s.start >= b.off || s.start + s.len <= a.off) continue;
      const f = readRunFormat(s.run);
      for (const k of ['bold', 'italic', 'underline', 'strike'] as const) if (!f[k]) base[k] = false;
      if (f.sizePt !== base.sizePt) delete base.sizePt;
      if (f.font !== base.font) delete base.font;
    }
    return base;
  }

  // ─── DOM ↔ position ────────────────────────────────────────────────────────

  private nodeOfElement(el: Element): XmlNode | undefined {
    const node = this.o.ctx.paragraphNode(el.getAttribute(EDIT_ATTR.para));
    return node && this.o.session.map(node) ? node : undefined;
  }

  private fragmentsOf(p: XmlNode): HTMLElement[] {
    if (!this.fragments) {
      this.fragments = new Map();
      for (const el of this.o.host.querySelectorAll<HTMLElement>(`[${EDIT_ATTR.para}]`)) {
        const node = this.nodeOfElement(el);
        if (!node) continue;
        const list = this.fragments.get(node);
        if (list) list.push(el);
        else this.fragments.set(node, [el]);
      }
    }
    return this.fragments.get(p) ?? [];
  }

  /** The stamped pieces of one rendered paragraph fragment, with their atom ranges. */
  private pieces(frag: HTMLElement, pm: ParaMap): Piece[] {
    const out: Piece[] = [];
    for (const el of frag.querySelectorAll(`[${EDIT_ATTR.run}]`)) {
      if (el.closest(`[${EDIT_ATTR.para}]`) !== frag) continue;
      const ref = this.o.ctx.runNode(el.getAttribute(EDIT_ATTR.run));
      const span = ref ? pm.bySource.get(ref.node) : undefined;
      if (!ref || !span) continue;
      const off = ref.off;
      const part = el.getAttribute(EDIT_ATTR.part);
      if (part === 'br' || part === 'tab') {
        const at = part === 'br' ? span.brAt : span.tabAt;
        if (at !== undefined) out.push({ el, start: at, end: at + 1, text: false });
      } else if (span.kind === 'obj') {
        out.push({ el, start: span.at, end: span.at + 1, text: false });
      } else if (span.kind === 'text') {
        const start = span.at + off;
        out.push({ el, start, end: Math.min(span.at + span.len, start + (el.textContent ?? '').length), text: true });
      }
    }
    return out;
  }

  /** The document position of a DOM point, or undefined outside the editable text. */
  domToPos(node: Node, offset: number): DocPos | undefined {
    const start = node.nodeType === 1 ? (node as Element) : node.parentElement;
    let frag = start?.closest<HTMLElement>(`[${EDIT_ATTR.para}]`) ?? null;
    if (!frag && node.nodeType === 1) {
      // A point between blocks: use the paragraph just after it (or before).
      const el = node as Element;
      const next = el.childNodes[offset] as Element | undefined;
      const prev = el.childNodes[offset - 1] as Element | undefined;
      const find = (n: Element | undefined, last: boolean) => {
        if (!n || n.nodeType !== 1) return null;
        if (n.hasAttribute(EDIT_ATTR.para)) return n as HTMLElement;
        const all = n.querySelectorAll<HTMLElement>(`[${EDIT_ATTR.para}]`);
        return (last ? all[all.length - 1] : all[0]) ?? null;
      };
      frag = find(next, false);
      if (frag) {
        const p = this.nodeOfElement(frag);
        const pm = p ? this.o.session.map(p) : undefined;
        const first = pm ? this.pieces(frag, pm)[0] : undefined;
        return p ? { p, off: first?.start ?? 0 } : undefined;
      }
      frag = find(prev, true);
      if (frag) {
        const p = this.nodeOfElement(frag);
        const pm = p ? this.o.session.map(p) : undefined;
        const pieces = pm ? this.pieces(frag, pm) : [];
        return p ? { p, off: pieces[pieces.length - 1]?.end ?? 0 } : undefined;
      }
      return undefined;
    }
    if (!frag) return undefined;
    const p = this.nodeOfElement(frag);
    const pm = p ? this.o.session.map(p) : undefined;
    if (!p || !pm) return undefined;
    const doc = frag.ownerDocument;
    const pieces = this.pieces(frag, pm);
    let lastEnd: number | undefined;
    for (const piece of pieces) {
      const r = doc.createRange();
      r.selectNode(piece.el);
      let cmp: number;
      try {
        cmp = r.comparePoint(node, offset);
      } catch {
        continue;
      }
      if (cmp < 0) return { p, off: lastEnd ?? piece.start };
      if (cmp === 0) {
        if (!piece.text) {
          // Inside an object: snap to whichever side is nearer the point's start.
          return { p, off: offset === 0 && node === piece.el ? piece.start : piece.end };
        }
        const before = doc.createRange();
        before.setStart(piece.el, 0);
        before.setEnd(node, offset);
        const n = before.toString().length;
        return { p, off: Math.max(piece.start, Math.min(piece.end, piece.start + n)) };
      }
      lastEnd = piece.end;
    }
    return { p, off: lastEnd ?? 0 };
  }

  /** The DOM point for a document position. */
  posToDom(pos: DocPos): { node: Node; offset: number } | undefined {
    const pm = this.o.session.map(pos.p);
    const frags = this.fragmentsOf(pos.p);
    if (!pm || frags.length === 0) return undefined;
    let fallback: { node: Node; offset: number } | undefined;
    for (const frag of frags) {
      const pieces = this.pieces(frag, pm);
      // Text first: the caret at a run boundary sits at the end of the left run.
      for (const piece of pieces) {
        if (piece.text && piece.start <= pos.off && pos.off <= piece.end && piece.start < piece.end) {
          const textNode = firstText(piece.el);
          if (textNode) return { node: textNode, offset: Math.min(textNode.data.length, pos.off - piece.start) };
        }
      }
      for (const piece of pieces) {
        const parent = piece.el.parentNode;
        if (!parent) continue;
        const index = Array.prototype.indexOf.call(parent.childNodes, piece.el) as number;
        if (piece.start === pos.off) return { node: parent, offset: index };
        if (piece.end === pos.off) fallback = { node: parent, offset: index + 1 };
      }
      if (pieces.length === 0 && pos.off === 0) {
        // An empty paragraph: its zero-width space holds the line.
        const zw = [...frag.childNodes].find((n) => n.nodeType === 3) as Text | undefined;
        return zw ? { node: zw, offset: 0 } : { node: frag, offset: frag.childNodes.length };
      }
    }
    return fallback ?? { node: frags[0]!, offset: 0 };
  }

  /** The current DOM selection as document positions. */
  readSelection(): DocSel | undefined {
    const s = this.o.host.ownerDocument.getSelection();
    if (!s || s.rangeCount === 0 || !s.anchorNode || !s.focusNode) return undefined;
    if (!this.o.host.contains(s.anchorNode)) return undefined;
    const anchor = this.domToPos(s.anchorNode, s.anchorOffset);
    const focus = this.domToPos(s.focusNode, s.focusOffset) ?? anchor;
    return anchor && focus ? { anchor, focus } : undefined;
  }

  setSelection(sel: DocSel): void {
    const a = this.posToDom(sel.anchor);
    const f = this.posToDom(sel.focus);
    const s = this.o.host.ownerDocument.getSelection();
    if (!a || !f || !s) return;
    this.settingSelection = true;
    try {
      s.setBaseAndExtent(a.node, a.offset, f.node, f.offset);
    } finally {
      this.settingSelection = false;
    }
    this.lastSel = sel;
  }
}

function samePos(a: DocPos, b: DocPos): boolean {
  return a.p === b.p && a.off === b.off;
}

function firstText(el: Node): Text | undefined {
  if (el.nodeType === 3) return el as Text;
  for (const c of el.childNodes) {
    const t = firstText(c);
    if (t) return t;
  }
  return undefined;
}
