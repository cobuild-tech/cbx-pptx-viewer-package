/**
 * DocxEditSession — the seam between a UI and the XML.
 *
 * Word's editing model: the document is one flow of paragraphs, a caret is a
 * position in it ({@link DocPos}) and a selection a pair of them, and every
 * command — typing, Enter, Backspace, a toolbar button — is a transaction that
 * edits `word/document.xml` in place (see ops.ts), re-parses, and hands back
 * where the caret now is. Undo restores whole-part snapshots, so it is exact by
 * construction; consecutive keystrokes share one snapshot, as Word undoes
 * typing a stretch at a time rather than a letter at a time.
 *
 * Every model reference goes stale on each transaction (the re-parse builds
 * fresh paragraphs), but positions stay valid: they are addressed by the
 * `<w:p>` node, which an edit keeps in place.
 */
import type { Snapshot } from '../../oxml/edit/history.js';
import type { DocxDocument } from '../document/document.js';
import type { DocxParagraph, DocxRun, DocxSection } from '../model.js';
import { child, localName, type XmlNode } from '../../oxml/xml.js';
import type { ParaFormat, RunFormat } from '../../oxml/edit/format.js';
import { DocxRelType } from '../relTypes.js';
import { writeParagraphs, wordPrefix, type DocxParaEdit } from './xmlWrite.js';
import { XmlEditor, PPR_ORDER, preserveSpace, textHostAt, type DocPos } from './ops.js';
import {
  mapParagraph,
  storyParagraphs,
  wordStartBefore,
  wordEndAfter,
  prevCharBoundary,
  nextCharBoundary,
  type ParaMap,
} from './story.js';
import {
  listKind,
  createList,
  readParaStyles,
  EMPTY_NUMBERING,
  NUMBERING_CONTENT_TYPE,
  type ParaStyleInfo,
} from './lists.js';
import { cloneNode, createElement } from '../../oxml/xml.js';

export type { DocPos } from './ops.js';

/** A selection: where it was started and where it now ends (equal for a caret). */
export interface DocSel {
  anchor: DocPos;
  focus: DocPos;
}

/** Paragraph formatting as the Home ribbon shows it, plus the paragraph style. */
export interface DocxParaFormat extends ParaFormat {
  styleId?: string;
}

export interface DocxEditSessionOptions {
  /** Called after any change to the document (an edit, undo or redo). */
  onChange?: () => void;
  /** Cap on retained undo steps. */
  historyLimit?: number;
}

/** A paragraph by its position in document order — what survives an undo's re-parse. */
interface Address {
  i: number;
  off: number;
}

interface UndoStep {
  snaps: Snapshot[];
  before: { anchor: Address; focus: Address };
  after?: { anchor: Address; focus: Address };
}

/** Keystrokes closer together than this undo as one step. */
const TYPING_GROUP_MS = 1500;

export function caret(p: XmlNode, off: number): DocSel {
  return { anchor: { p, off }, focus: { p, off } };
}

export function isCollapsed(sel: DocSel): boolean {
  return sel.anchor.p === sel.focus.p && sel.anchor.off === sel.focus.off;
}

export class DocxEditSession {
  private readonly doc: DocxDocument;
  private readonly onChange: DocxEditSessionOptions['onChange'];
  private readonly limit: number;
  private undoStack: UndoStep[] = [];
  private redoStack: UndoStep[] = [];
  /** The last keystroke: typing continues its undo step if it carries on from there. */
  private lastTyping: { kind: string; at: DocPos; time: number } | undefined;
  private byNode = new Map<XmlNode, DocxParagraph>();
  private order = new Map<XmlNode, number>();
  private bodyOrder: XmlNode[] = [];
  private maps = new Map<XmlNode, ParaMap>();
  /** Where the caret belongs after the last undo or redo. */
  lastSelection: DocSel | undefined;
  /**
   * Set by an edit that only changed the text of one `<w:t>`: the model is
   * patched to match instead of re-parsing the whole document.
   */
  private inPlace: { run: DocxRun; t: XmlNode } | undefined;

  constructor(doc: DocxDocument, options: DocxEditSessionOptions = {}) {
    this.doc = doc;
    this.limit = options.historyLimit ?? 100;
    if (options.onChange) this.onChange = options.onChange;
    this.reindex();
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  get hasEdits(): boolean {
    return this.doc.hasEdits;
  }

  /** True if this paragraph is in the main document part (not a header/footer). */
  isEditable(para: DocxParagraph): boolean {
    return this.doc.isEditable(para);
  }

  // ─── Index ─────────────────────────────────────────────────────────────────

  /** Re-read which `<w:p>` each parsed paragraph came from. Call after every re-parse. */
  reindex(): void {
    this.byNode = new Map();
    this.order = new Map();
    this.maps = new Map();
    const blocks = this.doc.sections.flatMap((s) => s.blocks);
    const { body, boxes } = storyParagraphs(blocks);
    this.bodyOrder = [];
    for (const para of [...body, ...boxes]) {
      const src = this.doc.sourceOf(para);
      if (!src || !this.doc.isEditable(para)) continue;
      this.byNode.set(src.node, para);
      this.order.set(src.node, this.order.size);
    }
    for (const para of body) {
      const node = this.doc.sourceOf(para)?.node;
      if (node && this.order.has(node)) this.bodyOrder.push(node);
    }
  }

  /** The parsed paragraph a `<w:p>` currently stands for. */
  paragraphFor(p: XmlNode): DocxParagraph | undefined {
    return this.byNode.get(p);
  }

  /** The `<w:p>` a parsed paragraph came from, if it is editable. */
  nodeOf(para: DocxParagraph): XmlNode | undefined {
    const node = this.doc.sourceOf(para)?.node;
    return node && this.byNode.get(node) === para ? node : undefined;
  }

  /** The atom map of a paragraph, from the current parse. */
  map(p: XmlNode): ParaMap | undefined {
    const cached = this.maps.get(p);
    if (cached) return cached;
    const para = this.byNode.get(p);
    const pm = para ? mapParagraph(this.doc, para) : undefined;
    if (pm) this.maps.set(p, pm);
    return pm;
  }

  /** Atom map of any paragraph, layout copies resolved by the caller. */
  mapOf(para: DocxParagraph): ParaMap | undefined {
    const node = this.nodeOf(para);
    return node ? this.map(node) : undefined;
  }

  /** Document order of two positions. */
  compare(a: DocPos, b: DocPos): number {
    if (a.p === b.p) return a.off - b.off;
    return (this.order.get(a.p) ?? 0) - (this.order.get(b.p) ?? 0);
  }

  /** A selection's ends in document order. */
  ordered(sel: DocSel): [DocPos, DocPos] {
    return this.compare(sel.anchor, sel.focus) <= 0 ? [sel.anchor, sel.focus] : [sel.focus, sel.anchor];
  }

  /** Every editable paragraph from `a` to `b`, in order. */
  paragraphsBetween(a: XmlNode, b: XmlNode): XmlNode[] {
    if (a === b) return [a];
    const ia = this.bodyOrder.indexOf(a);
    const ib = this.bodyOrder.indexOf(b);
    if (ia !== -1 && ib !== -1) return this.bodyOrder.slice(Math.min(ia, ib), Math.max(ia, ib) + 1);
    return [a];
  }

  /** The first and last positions of the document, for Select All. */
  documentRange(): DocSel | undefined {
    const first = this.bodyOrder[0];
    const last = this.bodyOrder[this.bodyOrder.length - 1];
    if (!first || !last) return undefined;
    return { anchor: { p: first, off: 0 }, focus: { p: last, off: this.map(last)?.length ?? 0 } };
  }

  // ─── Transactions ──────────────────────────────────────────────────────────

  /**
   * Run one edit: snapshot (unless this keystroke continues the previous
   * one's undo step), mutate, re-parse, re-index. `fn` returns the new
   * selection, or undefined having changed nothing.
   */
  private transact(
    kind: 'edit' | 'type' | 'delete',
    before: DocSel,
    fn: (x: XmlEditor) => DocSel | undefined,
    extraParts: string[] = [],
  ): DocSel | undefined {
    const root = this.doc.documentXml();
    const body = child(root, 'body');
    if (!root || !body) return undefined;
    const now = Date.now();
    const last = this.lastTyping;
    const continues =
      kind !== 'edit' &&
      last?.kind === kind &&
      isCollapsed(before) &&
      last.at.p === before.focus.p &&
      last.at.off === before.focus.off &&
      now - last.time < TYPING_GROUP_MS &&
      this.undoStack.length > 0;
    const parts = [this.doc.mainPart, ...extraParts];
    const snaps = continues
      ? []
      : parts.flatMap((part) => {
          const xml = this.doc.snapshotPart(part);
          return xml === undefined ? [] : [{ part, xml }];
        });
    const beforeAddr = { anchor: this.address(before.anchor), focus: this.address(before.focus) };

    this.inPlace = undefined;
    const x = new XmlEditor(root, wordPrefix(body));
    const result = fn(x);
    if (!result) return undefined;
    // Read back through a cast: TypeScript cannot see `fn` setting the field.
    const patch = this.inPlace as { run: DocxRun; t: XmlNode } | undefined;
    if (!patch) x.clearMarkers(root);

    if (!continues) {
      this.undoStack.push({ snaps, before: beforeAddr });
      if (this.undoStack.length > this.limit) this.undoStack.shift();
    }
    this.redoStack = [];
    this.doc.markDirty();
    for (const part of extraParts) this.doc.markPartDirty(part);
    if (patch) {
      patch.run.text = patch.t.text;
      this.maps.delete(result.focus.p);
      this.inPlace = undefined;
    } else {
      this.doc.rebuild();
      this.reindex();
    }
    const top = this.undoStack[this.undoStack.length - 1];
    if (top) top.after = { anchor: this.address(result.anchor), focus: this.address(result.focus) };
    this.lastTyping = kind === 'edit' ? undefined : { kind, at: result.focus, time: now };
    this.onChange?.();
    return result;
  }

  /** Every `<w:p>` in the document part, in document order. */
  private allParagraphs(): XmlNode[] {
    const out: XmlNode[] = [];
    const walk = (n: XmlNode) => {
      for (const c of n.children) {
        if (localName(c.name) === 'p') out.push(c);
        walk(c);
      }
    };
    const root = this.doc.documentXml();
    if (root) walk(root);
    return out;
  }

  private address(pos: DocPos): Address {
    return { i: this.allParagraphs().indexOf(pos.p), off: pos.off };
  }

  private resolve(addr: { anchor: Address; focus: Address }): DocSel | undefined {
    const all = this.allParagraphs();
    const pos = (a: Address): DocPos | undefined => {
      const p = all[a.i];
      if (!p) return undefined;
      return { p, off: Math.min(a.off, this.map(p)?.length ?? 0) };
    };
    const anchor = pos(addr.anchor);
    const focus = pos(addr.focus);
    return anchor && focus ? { anchor, focus } : undefined;
  }

  // ─── Text ──────────────────────────────────────────────────────────────────

  /**
   * Type `text` over the selection. `\t` inserts a tab and `\n` a line break;
   * use {@link insertParagraph} or {@link paste} for paragraph breaks.
   * `format` is a pending toolbar format (Ctrl+B with nothing selected).
   */
  insertText(sel: DocSel, text: string, format?: RunFormat): DocSel | undefined {
    if (!text) return undefined;
    const [a, b] = this.ordered(sel);
    const pm = this.map(a.p);
    if (!pm) return undefined;
    const collapsed = a.p === b.p && a.off === b.off;
    return this.transact(collapsed ? 'type' : 'edit', sel, (x) => {
      if (collapsed && !format && !/[\t\n]/.test(text)) {
        const host = textHostAt(pm, a.off);
        if (host) {
          const t = host.seg.node;
          t.text = t.text.slice(0, host.at) + text + t.text.slice(host.at);
          preserveSpace(t);
          this.inPlace = { run: host.seg.run, t };
          return caret(a.p, a.off + text.length);
        }
      }
      const rPr = this.template(x, pm, a.off, !collapsed);
      const m = collapsed ? this.insertionMarker(x, pm, a.off) : this.deleteBetween(x, a, b);
      if (!m) return undefined;
      x.insertBefore(m, x.buildRun(text, rPr, format));
      x.coalesce(x.parent(m)!);
      return caret(a.p, a.off + text.length);
    });
  }

  /**
   * Paste plain text: each line after the first becomes a new paragraph with
   * the properties of the one it was pasted into, as Word pastes text.
   */
  paste(sel: DocSel, text: string): DocSel | undefined {
    const lines = text.replace(/\r\n?/g, '\n').split('\n');
    if (lines.length === 1) return this.insertText(sel, text);
    const [a, b] = this.ordered(sel);
    const pm = this.map(a.p);
    if (!pm) return undefined;
    return this.transact('edit', sel, (x) => {
      const rPr = this.template(x, pm, a.off, a.p !== b.p || a.off !== b.off);
      const m = a.p === b.p && a.off === b.off ? x.marker(pm, a.off) : this.deleteBetween(x, a, b);
      if (!m) return undefined;
      lines.forEach((line, i) => {
        if (line) x.insertBefore(m, x.buildRun(line, rPr));
        if (i < lines.length - 1) x.splitParagraphAt(a.p, m);
      });
      return caret(a.p, lines[lines.length - 1]!.length);
    });
  }

  /** Enter: split the paragraph at the caret (after deleting any selection). */
  insertParagraph(sel: DocSel): DocSel | undefined {
    const [a, b] = this.ordered(sel);
    const pm = this.map(a.p);
    const para = this.paragraphFor(a.p);
    if (!pm || !para) return undefined;
    const collapsed = a.p === b.p && a.off === b.off;
    // Enter on an empty list item ends the list, as in Word.
    if (collapsed && para.listMarker && pm.length === 0) return this.removeListFrom([a.p], sel);
    const atEnd = collapsed && a.off === pm.length && pm.length > 0;
    return this.transact('edit', sel, (x) => {
      const m = collapsed ? x.marker(pm, a.off) : this.deleteBetween(x, a, b);
      if (!m) return undefined;
      x.splitParagraphAt(a.p, m);
      // Enter at the end of a heading starts the style's "next" paragraph.
      if (atEnd) {
        const styles = this.styles();
        const id = this.styleIdOf(a.p) ?? styles.defaultId;
        const next = styles.list.find((s) => s.id === id)?.next;
        if (next && next !== id) x.setParaVal(a.p, 'pStyle', next === styles.defaultId ? undefined : next);
      }
      return caret(a.p, 0);
    });
  }

  /** Delete the selected content. */
  deleteSelection(sel: DocSel): DocSel | undefined {
    if (isCollapsed(sel)) return undefined;
    const [a] = this.ordered(sel);
    return this.transact('edit', sel, (x) => (this.deleteBetween(x, ...this.ordered(sel)) ? caret(a.p, a.off) : undefined));
  }

  /** Backspace (or Ctrl+Backspace with `unit: 'word'`). */
  deleteBackward(sel: DocSel, unit: 'char' | 'word' = 'char'): DocSel | undefined {
    if (!isCollapsed(sel)) return this.deleteSelection(sel);
    const { p, off } = sel.focus;
    const pm = this.map(p);
    const para = this.paragraphFor(p);
    if (!pm || !para) return undefined;
    if (off === 0) {
      // At a list item's start Backspace first takes the bullet off.
      if (para.listMarker) return this.removeListFrom([p], sel);
      const prev = this.siblingParagraph(p, -1);
      const prevMap = prev ? this.map(prev) : undefined;
      if (!prev || !prevMap) return undefined;
      return this.transact('edit', sel, (x) => {
        x.joinParagraphs(prev, p, prevMap.length === 0);
        return caret(prev, prevMap.length);
      });
    }
    const start = unit === 'word' ? wordStartBefore(pm.text, off) : prevCharBoundary(pm.text, off);
    return this.deleteInParagraph(sel, pm, start, off, unit === 'char' ? 'delete' : 'edit');
  }

  /** Delete (or Ctrl+Delete with `unit: 'word'`). */
  deleteForward(sel: DocSel, unit: 'char' | 'word' = 'char'): DocSel | undefined {
    if (!isCollapsed(sel)) return this.deleteSelection(sel);
    const { p, off } = sel.focus;
    const pm = this.map(p);
    if (!pm) return undefined;
    if (off >= pm.length) {
      const next = this.siblingParagraph(p, 1);
      if (!next || !this.map(next)) return undefined;
      return this.transact('edit', sel, (x) => {
        x.joinParagraphs(p, next, pm.length === 0);
        return caret(p, off);
      });
    }
    const end = unit === 'word' ? wordEndAfter(pm.text, off) : nextCharBoundary(pm.text, off);
    return this.deleteInParagraph(sel, pm, off, end, unit === 'char' ? 'delete' : 'edit');
  }

  private deleteInParagraph(sel: DocSel, pm: ParaMap, start: number, end: number, kind: 'delete' | 'edit'): DocSel | undefined {
    if (start >= end) return undefined;
    const p = pm.node;
    return this.transact(kind, sel, (x) => {
      // The common case — a few characters inside one run — edits the text in place.
      const seg = pm.segs.find((s) => s.kind === 'text' && s.start <= start && end <= s.start + s.len);
      if (seg) {
        const t = seg.node;
        t.text = t.text.slice(0, start - seg.start) + t.text.slice(end - seg.start);
        preserveSpace(t);
        if (!t.text) {
          const r = x.parent(t);
          x.remove(t);
          if (r) x.prune(x.parent(r) ?? r);
        } else {
          this.inPlace = { run: seg.run, t };
        }
        return caret(p, start);
      }
      const mE = x.marker(pm, end);
      const mS = x.marker(pm, start);
      x.widenToFields(p, mS, mE);
      x.deleteNodes(x.between(p, mS, mE));
      x.prune(p);
      x.coalesce(p);
      return caret(p, start);
    });
  }

  /**
   * Delete from `a` to `b` and return a marker where they met. Within one
   * container the paragraphs between go and the ends join; across containers
   * (into a table) only the text goes, as Word clears cells it cannot merge.
   */
  private deleteBetween(x: XmlEditor, a: DocPos, b: DocPos): XmlNode | undefined {
    const pmA = this.map(a.p);
    const pmB = this.map(b.p);
    if (!pmA || !pmB) return undefined;
    if (a.p === b.p) {
      const mB = x.marker(pmA, b.off);
      const mA = x.marker(pmA, a.off);
      x.widenToFields(a.p, mA, mB);
      x.deleteNodes(x.between(a.p, mA, mB));
      x.remove(mB);
      x.prune(a.p);
      return mA;
    }
    const parent = x.parent(a.p);
    if (parent && parent === x.parent(b.p)) {
      const mA = x.marker(pmA, a.off);
      const mB = x.marker(pmB, b.off);
      x.deleteNodes(x.between(a.p, mA, undefined));
      x.deleteNodes(x.between(b.p, undefined, mB));
      x.remove(mB);
      const from = parent.children.indexOf(a.p);
      const to = parent.children.indexOf(b.p);
      for (const c of parent.children.slice(from + 1, to)) {
        if (localName(c.name) !== 'sectPr') x.remove(c);
      }
      x.joinParagraphs(a.p, b.p, a.off === 0);
      x.prune(a.p);
      return mA;
    }
    let mA: XmlNode | undefined;
    for (const p of this.paragraphsBetween(a.p, b.p)) {
      const pm = this.map(p);
      if (!pm) continue;
      const s = p === a.p ? a.off : 0;
      const e = p === b.p ? b.off : pm.length;
      const mE = x.marker(pm, e);
      const mS = x.marker(pm, s);
      if (s < e) x.deleteNodes(x.between(p, mS, mE));
      x.remove(mE);
      if (p === a.p) mA = mS;
      else x.remove(mS);
      x.prune(p);
    }
    return mA;
  }

  /** Where typed text goes: after a field rather than into its result. */
  private insertionMarker(x: XmlEditor, pm: ParaMap, off: number): XmlNode {
    const left = [...pm.segs].reverse().find((s) => s.len > 0 && s.start + s.len === off);
    if (left && left.run.fieldCode !== undefined) {
      const m = x.markerAfterField(left);
      if (m) return m;
    }
    return x.marker(pm, off);
  }

  /**
   * The run properties new text takes, cloned before anything is cut: the
   * text it continues, or — replacing a selection — the first selected
   * character's, as Word does.
   */
  private template(x: XmlEditor, pm: ParaMap, off: number, preferRight: boolean): XmlNode | undefined {
    if (preferRight) {
      const right = pm.segs.find((s) => s.start === off && s.len > 0);
      const r = right ? x.parent(right.node) : undefined;
      if (r && localName(r.name) === 'r') {
        const rPr = child(r, 'rPr');
        return rPr ? cloneNode(rPr) : undefined;
      }
    }
    const rPr = x.templateRPr(pm, off);
    return rPr ? cloneNode(rPr) : undefined;
  }

  /** The neighbouring `<w:p>` in the same container, skipping bookmarks; undefined past a table. */
  private siblingParagraph(p: XmlNode, dir: 1 | -1): XmlNode | undefined {
    const root = this.doc.documentXml();
    if (!root) return undefined;
    const x = new XmlEditor(root, '');
    const parent = x.parent(p);
    if (!parent) return undefined;
    for (let i = parent.children.indexOf(p) + dir; i >= 0 && i < parent.children.length; i += dir) {
      const c = parent.children[i]!;
      const local = localName(c.name);
      if (local === 'p') return this.byNode.has(c) ? c : undefined;
      if (local === 'bookmarkStart' || local === 'bookmarkEnd' || local.endsWith('RangeStart') || local.endsWith('RangeEnd')) continue;
      return undefined;
    }
    return undefined;
  }

  // ─── Formatting ────────────────────────────────────────────────────────────

  /** Apply character formatting to the selected text. */
  formatSelection(sel: DocSel, format: RunFormat): DocSel | undefined {
    if (isCollapsed(sel)) return undefined;
    const [a, b] = this.ordered(sel);
    return this.transact('edit', sel, (x) => {
      for (const p of this.paragraphsBetween(a.p, b.p)) {
        const pm = this.map(p);
        if (!pm) continue;
        const s = p === a.p ? a.off : 0;
        const e = p === b.p ? b.off : pm.length;
        if (s < e) {
          const mE = x.marker(pm, e);
          const mS = x.marker(pm, s);
          x.formatNodes(x.between(p, mS, mE), format);
          x.remove(mS);
          x.remove(mE);
          x.coalesce(p);
        }
        // A whole paragraph's mark takes the format too, so text typed at its
        // end — or into it once emptied — keeps it.
        if (s === 0 && e === pm.length) x.formatMark(p, format);
      }
      return sel;
    });
  }

  /** Set paragraph properties (alignment, spacing, style) on every selected paragraph. */
  setParagraphFormat(sel: DocSel, format: DocxParaFormat): DocSel | undefined {
    const [a, b] = this.ordered(sel);
    const paras = this.paragraphsBetween(a.p, b.p);
    const styles = this.styles();
    return this.transact('edit', sel, (x) => {
      for (const p of paras) {
        if (format.align) {
          const jc = { left: 'left', center: 'center', right: 'right', justify: 'both' }[format.align];
          x.setParaVal(p, 'jc', jc);
        }
        if (format.lineSpacingPct !== undefined) {
          x.setParaAttrs(p, 'spacing', { line: String(Math.round(format.lineSpacingPct * 240)), lineRule: 'auto' });
        }
        if (format.spaceBeforePt !== undefined) {
          x.setParaAttrs(p, 'spacing', { before: String(Math.round(format.spaceBeforePt * 20)), beforeAutospacing: undefined });
        }
        if (format.spaceAfterPt !== undefined) {
          x.setParaAttrs(p, 'spacing', { after: String(Math.round(format.spaceAfterPt * 20)), afterAutospacing: undefined });
        }
        if (format.styleId !== undefined) {
          x.setParaVal(p, 'pStyle', format.styleId === styles.defaultId ? undefined : format.styleId);
        }
      }
      return sel;
    });
  }

  /** Turn the selected paragraphs into a list of `kind`, or — if they all are one — out of it. */
  toggleList(sel: DocSel, kind: 'bullet' | 'number'): DocSel | undefined {
    const [a, b] = this.ordered(sel);
    const paras = this.paragraphsBetween(a.p, b.p);
    if (paras.every((p) => this.listKindOf(p) === kind)) return this.removeListFrom(paras, sel);
    const existing = this.doc.relatedXml(DocxRelType.Numbering);
    return this.transact(
      'edit',
      sel,
      (x) => {
        const numbering = this.doc.ensureRelatedPart(DocxRelType.Numbering, 'numbering.xml', NUMBERING_CONTENT_TYPE, EMPTY_NUMBERING);
        // Continue the list just above when it is the same kind, as Word does;
        // otherwise start a new one.
        const prev = this.siblingParagraph(paras[0]!, -1);
        const prevNum = prev && this.listKindOf(prev) === kind ? this.numPrOf(prev) : undefined;
        const numId = prevNum?.numId ?? createList(numbering.xml, kind, x.prefix);
        this.doc.markPartDirty(numbering.part);
        const styles = this.styles();
        const listStyle = styles.list.find((s) => s.id === 'ListParagraph');
        for (const p of paras) {
          const ilvl = this.numPrOf(p)?.ilvl ?? 0;
          this.setNumPr(x, p, numId, ilvl);
          const style = this.styleIdOf(p) ?? styles.defaultId;
          if (listStyle && style === styles.defaultId) x.setParaVal(p, 'pStyle', listStyle.id);
        }
        return sel;
      },
      existing ? [existing.part] : [],
    );
  }

  /** Increase (+1) or decrease (−1) indent: a list item's level, else the left indent by ½ inch. */
  indent(sel: DocSel, delta: 1 | -1): DocSel | undefined {
    const [a, b] = this.ordered(sel);
    const paras = this.paragraphsBetween(a.p, b.p);
    return this.transact('edit', sel, (x) => {
      let changed = false;
      for (const p of paras) {
        const num = this.numPrOf(p);
        if (num && this.listKindOf(p)) {
          const ilvl = Math.max(0, Math.min(8, num.ilvl + delta));
          if (ilvl !== num.ilvl) {
            this.setNumPr(x, p, num.numId, ilvl);
            changed = true;
          }
          continue;
        }
        const para = this.paragraphFor(p);
        const cur = Math.round((para?.indentLeftPx ?? 0) * 15);
        const step = 720;
        const next = delta > 0 ? (Math.floor(cur / step) + 1) * step : Math.max(0, (Math.ceil(cur / step) - 1) * step);
        if (next === cur) continue;
        x.setParaAttrs(p, 'ind', { left: String(next), start: undefined });
        changed = true;
      }
      return changed ? sel : undefined;
    });
  }

  private removeListFrom(paras: XmlNode[], sel: DocSel): DocSel | undefined {
    return this.transact('edit', sel, (x) => {
      for (const p of paras) {
        const pPr = child(p, 'pPr');
        const numPr = child(pPr, 'numPr');
        if (numPr) x.remove(numPr);
        // Numbered by its style: switch the style's list off explicitly.
        if (this.paragraphFor(p)?.listMarker && !numPr) this.setNumPr(x, p, '0', 0);
        // Back to body text: the list style Word gave the item goes with the bullet.
        if (this.styleIdOf(p) === 'ListParagraph') x.setParaVal(p, 'pStyle', undefined);
      }
      return sel;
    });
  }

  private setNumPr(x: XmlEditor, p: XmlNode, numId: string, ilvl: number): void {
    const pPr = x.ensurePPr(p);
    const old = child(pPr, 'numPr');
    if (old) x.remove(old);
    const numPr = createElement(x.q('numPr'), {}, [
      createElement(x.q('ilvl'), { [x.q('val')]: String(ilvl) }),
      createElement(x.q('numId'), { [x.q('val')]: numId }),
    ]);
    x.insertOrdered(pPr, numPr, PPR_ORDER);
  }

  private numPrOf(p: XmlNode): { numId: string; ilvl: number } | undefined {
    const numPr = child(child(p, 'pPr'), 'numPr');
    const numId = numPr ? child(numPr, 'numId')?.attrs : undefined;
    const id = numId ? (Object.entries(numId).find(([k]) => localName(k) === 'val')?.[1]) : undefined;
    if (!id) return undefined;
    const ilvlAttrs = child(numPr, 'ilvl')?.attrs ?? {};
    const ilvl = Number(Object.entries(ilvlAttrs).find(([k]) => localName(k) === 'val')?.[1] ?? 0);
    return { numId: id, ilvl: Number.isFinite(ilvl) ? ilvl : 0 };
  }

  /** Whether a paragraph is a bulleted or numbered list item. */
  listKindOf(p: XmlNode): 'bullet' | 'number' | undefined {
    const para = this.paragraphFor(p);
    if (!para?.listMarker) return undefined;
    const num = this.numPrOf(p);
    const numbering = this.doc.relatedXml(DocxRelType.Numbering)?.xml;
    if (num) return listKind(numbering, num.numId, num.ilvl);
    // Numbered through its style: judge by the marker drawn.
    return /^[\p{L}\p{N}]+[.)]?$/u.test(para.listMarker.trim()) ? 'number' : 'bullet';
  }

  /** The paragraph's own style id (`<w:pStyle>`), undefined for the default style. */
  styleIdOf(p: XmlNode): string | undefined {
    const ps = child(child(p, 'pPr'), 'pStyle');
    return ps ? Object.entries(ps.attrs).find(([k]) => localName(k) === 'val')?.[1] : undefined;
  }

  /** The document's paragraph styles. */
  styles(): { list: ParaStyleInfo[]; defaultId?: string } {
    return readParaStyles(this.doc.relatedXml(DocxRelType.Styles)?.xml);
  }

  /**
   * The paragraph formatting in effect across a selection — a property the
   * paragraphs disagree about is left out.
   */
  paraFormatAt(sel: DocSel): DocxParaFormat {
    const [a, b] = this.ordered(sel);
    const styles = this.styles();
    const formats = this.paragraphsBetween(a.p, b.p).map((p): DocxParaFormat => {
      const para = this.paragraphFor(p);
      const align = para?.align;
      const out: DocxParaFormat = {
        align: align === 'ctr' ? 'center' : align === 'r' ? 'right' : align === 'just' ? 'justify' : 'left',
        list: this.listKindOf(p) ?? 'none',
        styleId: this.styleIdOf(p) ?? styles.defaultId ?? 'Normal',
      };
      if (para?.lineSpacingPct !== undefined && para.lineSpacingPt === undefined) out.lineSpacingPct = para.lineSpacingPct;
      if (para?.spaceBeforePt !== undefined) out.spaceBeforePt = para.spaceBeforePt;
      if (para?.spaceAfterPt !== undefined) out.spaceAfterPt = para.spaceAfterPt;
      const num = this.numPrOf(p);
      if (num) out.level = num.ilvl;
      return out;
    });
    const first = formats[0] ?? {};
    const out: DocxParaFormat = {};
    for (const key of Object.keys(first) as (keyof DocxParaFormat)[]) {
      if (formats.every((f) => f[key] === first[key])) (out as Record<string, unknown>)[key] = first[key];
    }
    return out;
  }

  // ─── History ───────────────────────────────────────────────────────────────

  undo(): DocxSection[] | undefined {
    return this.step(this.undoStack, this.redoStack, 'before');
  }

  redo(): DocxSection[] | undefined {
    return this.step(this.redoStack, this.undoStack, 'after');
  }

  private step(from: UndoStep[], to: UndoStep[], restoreTo: 'before' | 'after'): DocxSection[] | undefined {
    const entry = from.pop();
    if (!entry) return undefined;
    const inverse: Snapshot[] = [];
    for (const snap of entry.snaps) {
      const now = this.doc.snapshotPart(snap.part);
      if (now !== undefined) inverse.push({ part: snap.part, xml: now });
    }
    to.push({ ...entry, snaps: inverse });
    for (const snap of entry.snaps) this.doc.restorePart(snap.part, snap.xml);
    this.lastTyping = undefined;
    const sections = this.doc.rebuild();
    this.reindex();
    const addr = restoreTo === 'before' ? entry.before : (entry.after ?? entry.before);
    this.lastSelection = this.resolve(addr);
    this.onChange?.();
    return sections;
  }

  // ─── Whole-paragraph rewrite (the original, paragraph-at-a-time API) ───────

  /**
   * Replace one source paragraph with the paragraphs it became, then re-parse.
   * Returns the rebuilt sections, or undefined if the paragraph is not
   * editable. Prefer the positional commands above: this rebuilds the
   * paragraph from text segments, so content that is not text is dropped.
   */
  commitParagraph(para: DocxParagraph, edits: DocxParaEdit[]): DocxSection[] | undefined {
    const src = this.doc.sourceOf(para);
    if (!src || !this.doc.isEditable(para)) return undefined;
    const body = child(this.doc.documentXml(), 'body');
    if (!body) return undefined;
    const before = this.doc.snapshotPart(src.part);
    if (before !== undefined) {
      const here = { i: -1, off: 0 };
      this.undoStack.push({ snaps: [{ part: src.part, xml: before }], before: { anchor: here, focus: here } });
      this.redoStack = [];
    }
    writeParagraphs(body, src.node, edits, (model) => this.doc.sourceOf(model));
    this.doc.markDirty();
    this.lastTyping = undefined;
    const sections = this.doc.rebuild();
    this.reindex();
    this.onChange?.();
    return sections;
  }

  /** Re-zip the document with all edits applied. */
  exportBlob(): Blob {
    return this.doc.exportBlob();
  }

  /** Re-zip the document with all edits applied, as raw bytes. */
  exportBytes(): Uint8Array {
    return this.doc.toBytes();
  }
}
