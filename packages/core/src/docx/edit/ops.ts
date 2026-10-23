/**
 * Word editing as operations on the document XML.
 *
 * Every user action — a keystroke, Enter, Backspace, a toolbar command — is
 * applied here directly to `word/document.xml`, addressed by paragraph and
 * atom offset ({@link ParaMap}). Nothing is rebuilt from the edited text: the
 * operations cut runs only where an edit begins or ends, so everything else in
 * a paragraph — pictures, footnote references, equations, fields, bookmarks,
 * tracked changes, properties the parser never read — stays exactly as it was.
 *
 * The working tool is a *marker*: an element placed at a run boundary for an
 * offset (splitting the run that straddles it). Once the markers for an
 * action are down, the action is a plain tree operation between them — delete
 * the nodes between two markers, split the paragraph at one, insert a run
 * before one — and no offset needs re-resolving after the tree has changed.
 * Markers never survive an operation.
 */
import {
  child,
  cloneNode,
  createElement,
  insertInOrder,
  localName,
  setAttr,
  type XmlNode,
} from '../../oxml/xml.js';
import type { RunFormat } from '../../oxml/edit/format.js';
import { applyFormat } from './format.js';
import { isLockedRun, type ParaMap, type Seg } from './story.js';

/** A caret position: an offset into a paragraph's atoms. */
export interface DocPos {
  /** The `<w:p>`. */
  p: XmlNode;
  off: number;
}

const MARKER = 'cbx:marker';

/** `<w:pPr>` child order (CT_PPr). Word rejects properties out of sequence. */
export const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
  'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
  'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
  'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
  'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
  'sectPr', 'pPrChange',
];

/** Inline containers that may be dropped once an edit has emptied them. */
const PRUNABLE = new Set(['hyperlink', 'ins', 'del', 'moveTo', 'moveFrom', 'smartTag', 'customXml', 'dir', 'bdo']);

/** Zero-width structure an edit between two points leaves where it is. */
function isRangeMark(local: string): boolean {
  return (
    local === 'bookmarkStart' ||
    local === 'bookmarkEnd' ||
    local === 'permStart' ||
    local === 'permEnd' ||
    local === 'proofErr' ||
    local.endsWith('RangeStart') ||
    local.endsWith('RangeEnd')
  );
}

/** Property children a container keeps on both halves when it is split. */
function isPropertyChild(local: string): boolean {
  return local === 'pPr' || local === 'rPr' || local === 'sdtPr' || local === 'sdtEndPr';
}

/** Elements whose insides are another story (a text box) or not text at all. */
const OPAQUE = new Set(['drawing', 'pict', 'object', 'AlternateContent', 'oMath', 'oMathPara', 'txbxContent']);

/**
 * The XML of one document, with a parent index kept current through every
 * edit — `XmlNode` has no parent pointer, and every operation here needs one.
 */
export class XmlEditor {
  readonly prefix: string;
  private readonly parents = new Map<XmlNode, XmlNode>();
  /** Runs this operation split or created, the only ones it may coalesce. */
  private readonly touched = new Set<XmlNode>();

  constructor(root: XmlNode, prefix: string) {
    this.prefix = prefix;
    const index = (n: XmlNode) => {
      for (const c of n.children) {
        this.parents.set(c, n);
        index(c);
      }
    };
    index(root);
  }

  q(local: string): string {
    return this.prefix ? `${this.prefix}:${local}` : local;
  }

  parent(n: XmlNode): XmlNode | undefined {
    return this.parents.get(n);
  }

  /** Insert `n` into `parent` at the slot a sequence-ordered schema gives it. */
  insertOrdered(parent: XmlNode, n: XmlNode, order: string[]): void {
    insertInOrder(parent, n, order);
    this.adopt(parent, n);
  }

  private adopt(parent: XmlNode, n: XmlNode): void {
    this.parents.set(n, parent);
    for (const c of n.children) this.adopt(n, c);
  }

  insertAt(parent: XmlNode, index: number, n: XmlNode): void {
    parent.children.splice(index, 0, n);
    this.adopt(parent, n);
  }

  insertBefore(ref: XmlNode, n: XmlNode): void {
    const parent = this.parent(ref)!;
    this.insertAt(parent, parent.children.indexOf(ref), n);
  }

  insertAfter(ref: XmlNode, n: XmlNode): void {
    const parent = this.parent(ref)!;
    this.insertAt(parent, parent.children.indexOf(ref) + 1, n);
  }

  append(parent: XmlNode, n: XmlNode): void {
    this.insertAt(parent, parent.children.length, n);
  }

  remove(n: XmlNode): void {
    const parent = this.parent(n);
    if (!parent) return;
    const i = parent.children.indexOf(n);
    if (i !== -1) parent.children.splice(i, 1);
    this.parents.delete(n);
  }

  /** True if `n` is `anc` or lies inside it. */
  within(n: XmlNode, anc: XmlNode): boolean {
    for (let c: XmlNode | undefined = n; c; c = this.parent(c)) if (c === anc) return true;
    return false;
  }

  // ─── Cutting runs ──────────────────────────────────────────────────────────

  /** Split a `<w:t>` at character `i`; returns the new right-hand `<w:t>`. */
  private splitText(t: XmlNode, i: number): XmlNode {
    const right = createElement(t.name, {}, [], t.text.slice(i));
    t.text = t.text.slice(0, i);
    preserveSpace(t);
    preserveSpace(right);
    this.insertAfter(t, right);
    return right;
  }

  /**
   * Make `c` the first content of its run, moving it and everything after it
   * into a new run (with the same properties) right after. Returns the run `c`
   * now starts.
   */
  private splitRunBefore(r: XmlNode, c: XmlNode): XmlNode {
    const idx = r.children.indexOf(c);
    if (r.children.slice(0, idx).every((k) => localName(k.name) === 'rPr')) return r;
    const rPr = child(r, 'rPr');
    const moved = r.children.slice(idx);
    r.children = r.children.slice(0, idx);
    for (const m of moved) this.parents.delete(m);
    const right = createElement(r.name, { ...r.attrs }, [...(rPr ? [cloneNode(rPr)] : []), ...moved]);
    this.insertAfter(r, right);
    this.touched.add(r);
    this.touched.add(right);
    return right;
  }

  /** The element directly inside a run (`<w:r>`) that holds `node`, and that run. */
  private inRun(node: XmlNode): { r: XmlNode; c: XmlNode } | undefined {
    const r = this.parent(node);
    return r && localName(r.name) === 'r' ? { r, c: node } : undefined;
  }

  /**
   * A run-level gap at paragraph offset `off`: the parent and index where a
   * new run placed there reads at exactly that offset. Splits the run that
   * straddles it. `bias` decides which side of zero-width content the gap
   * falls on.
   */
  private gapAt(pm: ParaMap, off: number): { parent: XmlNode; index: number } {
    const inside = pm.segs.find((s) => s.kind === 'text' && s.start < off && off < s.start + s.len);
    if (inside) {
      const right = this.splitText(inside.node, off - inside.start);
      const host = this.inRun(right)!;
      return this.gapBefore(this.splitRunBefore(host.r, right));
    }
    const next = pm.segs.find((s) => s.start === off && s.len > 0);
    if (next) {
      const host = this.inRun(next.node);
      if (!host) return this.gapBefore(next.node);
      return this.gapBefore(this.splitRunBefore(host.r, next.node));
    }
    const prev = [...pm.segs].reverse().find((s) => s.start + s.len === off && s.len > 0);
    if (prev) {
      const host = this.inRun(prev.node);
      if (!host) return this.gapAfter(prev.node);
      const after = host.r.children[host.r.children.indexOf(prev.node) + 1];
      if (after) return this.gapBefore(this.splitRunBefore(host.r, after));
      return this.gapAfter(host.r);
    }
    // No content at all: the end of the paragraph.
    return { parent: pm.node, index: pm.node.children.length };
  }

  private gapBefore(n: XmlNode): { parent: XmlNode; index: number } {
    const parent = this.parent(n)!;
    return { parent, index: parent.children.indexOf(n) };
  }

  private gapAfter(n: XmlNode): { parent: XmlNode; index: number } {
    const parent = this.parent(n)!;
    return { parent, index: parent.children.indexOf(n) + 1 };
  }

  /** Put a marker at offset `off`. Place a paragraph's later offsets first. */
  marker(pm: ParaMap, off: number): XmlNode {
    const gap = this.gapAt(pm, off);
    const m = createElement(MARKER);
    this.insertAt(gap.parent, gap.index, m);
    return m;
  }

  /**
   * A marker just after the field whose result `seg` is part of — where text
   * typed "after" a field must go, or Word would treat it as part of the
   * result and replace it when the field updates.
   */
  markerAfterField(seg: Seg): XmlNode | undefined {
    const host = this.inRun(seg.node);
    if (!host) return undefined;
    // A simple field holds its result inside it: go after the element.
    for (let c: XmlNode | undefined = host.r; c; c = this.parent(c)) {
      if (localName(c.name) === 'fldSimple') {
        const m = createElement(MARKER);
        this.insertAfter(c, m);
        return m;
      }
      if (localName(c.name) === 'p') break;
    }
    // A complex field: find the <w:fldChar w:fldCharType="end"> that closes it.
    const p = this.paragraphOf(host.r);
    if (!p) return undefined;
    let depth = 1;
    let seen = false;
    let end: XmlNode | undefined;
    const walk = (n: XmlNode): boolean => {
      for (const c of n.children) {
        if (c === host.r) {
          seen = true;
          continue;
        }
        if (seen && localName(c.name) === 'r') {
          for (const f of c.children) {
            if (localName(f.name) !== 'fldChar') continue;
            const type = f.attrs[this.q('fldCharType')] ?? f.attrs['fldCharType'];
            if (type === 'begin') depth++;
            else if (type === 'end' && --depth === 0) {
              end = c;
              return true;
            }
          }
        } else if (!OPAQUE.has(localName(c.name)) && walk(c)) return true;
      }
      return false;
    };
    walk(p);
    if (!end) return undefined;
    const m = createElement(MARKER);
    this.insertAfter(end, m);
    return m;
  }

  /** The `<w:p>` holding `n`. */
  paragraphOf(n: XmlNode): XmlNode | undefined {
    for (let c: XmlNode | undefined = n; c; c = this.parent(c)) if (localName(c.name) === 'p') return c;
    return undefined;
  }

  /**
   * The topmost nodes of `p` strictly between two markers (`from` undefined =
   * the start of the paragraph's content, `to` undefined = its end).
   */
  between(p: XmlNode, from: XmlNode | undefined, to: XmlNode | undefined): XmlNode[] {
    const out: XmlNode[] = [];
    let inside = from === undefined;
    let done = false;
    const visit = (n: XmlNode) => {
      for (const c of [...n.children]) {
        if (done) return;
        if (c === from) {
          inside = true;
          continue;
        }
        if (c === to) {
          done = true;
          return;
        }
        if (n === p && localName(c.name) === 'pPr') continue;
        if ((from && this.within(from, c)) || (to && this.within(to, c))) {
          visit(c);
          continue;
        }
        if (inside) out.push(c);
      }
    };
    visit(p);
    return out;
  }

  /**
   * Widen the range between two markers in `p` to take in whole fields whose
   * result it covers completely — deleting a field's visible text deletes the
   * field, as Word does, rather than leaving its code behind to regenerate it.
   */
  widenToFields(p: XmlNode, mS: XmlNode, mE: XmlNode): void {
    for (let pass = 0; pass < 4; pass++) {
      const seq: XmlNode[] = [];
      const walk = (n: XmlNode) => {
        for (const c of n.children) {
          const local = localName(c.name);
          if (c === mS || c === mE || local === 'r') seq.push(c);
          else if (!OPAQUE.has(local) && local !== 'pPr') walk(c);
        }
      };
      walk(p);
      const iS = seq.indexOf(mS);
      const iE = seq.indexOf(mE);
      const open: { b: number; s?: number }[] = [];
      let lo = iS;
      let hi = iE;
      seq.forEach((r, i) => {
        const fc = r.children.find((c) => localName(c.name) === 'fldChar');
        const type = fc ? (fc.attrs[this.q('fldCharType')] ?? fc.attrs['fldCharType']) : undefined;
        if (type === 'begin') open.push({ b: i });
        else if (type === 'separate' && open.length) open[open.length - 1]!.s = i;
        else if (type === 'end' && open.length) {
          const f = open.pop()!;
          const result = seq
            .slice((f.s ?? f.b) + 1, i)
            .map((n, k) => ({ n, k: k + (f.s ?? f.b) + 1 }))
            .filter(({ n }) => n.children.some((c) => localName(c.name) === 't'));
          const covered = result.length > 0 && result.every(({ k }) => k > iS && k < iE);
          if (covered) {
            lo = Math.min(lo, f.b);
            hi = Math.max(hi, i);
          }
        }
      });
      if (lo === iS && hi === iE) return;
      if (lo < iS) {
        this.remove(mS);
        this.insertBefore(seq[lo]!, mS);
      }
      if (hi > iE) {
        this.remove(mE);
        this.insertAfter(seq[hi]!, mE);
      }
    }
  }

  /**
   * Delete content nodes, leaving bookmarks and range marks where they were.
   * A field's begin/separate/end runs go only when the whole field does —
   * removing half of that structure would corrupt the rest of the paragraph.
   */
  deleteNodes(nodes: XmlNode[]): void {
    const fieldRuns: { r: XmlNode; type?: string }[] = [];
    for (const n of nodes) {
      if (localName(n.name) !== 'r') continue;
      const fc = n.children.find((c) => localName(c.name) === 'fldChar');
      const instr = n.children.some((c) => localName(c.name) === 'instrText');
      if (fc || instr) fieldRuns.push({ r: n, type: fc ? (fc.attrs[this.q('fldCharType')] ?? fc.attrs['fldCharType']) : undefined });
    }
    const keep = new Set(fieldRuns.map((f) => f.r));
    const open: number[] = [];
    fieldRuns.forEach((f, i) => {
      if (f.type === 'begin') open.push(i);
      else if (f.type === 'end' && open.length) {
        const b = open.pop()!;
        for (let k = b; k <= i; k++) keep.delete(fieldRuns[k]!.r);
      }
    });
    for (const n of nodes) {
      const local = localName(n.name);
      if (isRangeMark(local)) continue;
      if (keep.has(n)) {
        // Keep the field's skeleton, drop any text it carries.
        n.children = n.children.filter((c) => !['t', 'tab', 'br', 'cr'].includes(localName(c.name)));
        continue;
      }
      this.remove(n);
    }
  }

  // ─── Building runs ─────────────────────────────────────────────────────────

  /**
   * A run for `text` with properties cloned from `rPr`: `\t` becomes a
   * `<w:tab/>` and `\n` a `<w:br/>`.
   */
  buildRun(text: string, rPr: XmlNode | undefined, format?: RunFormat): XmlNode {
    const kids: XmlNode[] = [];
    const props = rPr ? cloneNode(rPr) : undefined;
    if (props) stripRevisionProps(props);
    if (format) {
      const target = props ?? createElement(this.q('rPr'));
      applyFormat(target, format, this.prefix);
      if (target.children.length) kids.push(target);
    } else if (props && props.children.length) kids.push(props);
    let buf = '';
    const flush = () => {
      if (!buf) return;
      const t = createElement(this.q('t'), {}, [], buf);
      preserveSpace(t);
      kids.push(t);
      buf = '';
    };
    for (const ch of text) {
      if (ch === '\t' || ch === '\n') {
        flush();
        kids.push(createElement(this.q(ch === '\t' ? 'tab' : 'br')));
      } else buf += ch;
    }
    flush();
    const r = createElement(this.q('r'), {}, kids);
    this.touched.add(r);
    return r;
  }

  /** The run properties text typed at `off` should take: those of the text it continues. */
  templateRPr(pm: ParaMap, off: number): XmlNode | undefined {
    const host = (seg: Seg | undefined) => (seg ? this.inRun(seg.node)?.r : undefined);
    const left = [...pm.segs].reverse().find((s) => s.start + s.len === off && s.len > 0);
    const right = pm.segs.find((s) => s.start === off && s.len > 0);
    const r = host(left) ?? host(right);
    if (r) return child(r, 'rPr');
    // An empty paragraph: the paragraph mark's own formatting, as Word does.
    return child(child(pm.node, 'pPr'), 'rPr');
  }

  // ─── Paragraph structure ───────────────────────────────────────────────────

  /**
   * Split paragraph `p` at marker `m`: everything before the marker moves to a
   * new paragraph inserted before `p`, which keeps its identity (and its
   * paragraph mark, with any section break it carries) as the second half.
   */
  splitParagraphAt(p: XmlNode, m: XmlNode): XmlNode {
    const left = createElement(p.name, { ...p.attrs });
    // w14:paraId must be unique per paragraph.
    for (const k of Object.keys(left.attrs)) if (/:(paraId|textId)$/.test(k)) delete left.attrs[k];
    const pPr = child(p, 'pPr');
    if (pPr) {
      const copy = cloneNode(pPr);
      copy.children = copy.children.filter((c) => localName(c.name) !== 'sectPr');
      left.children.push(copy);
    }
    this.insertBefore(p, left);
    this.cutInto(left, p, m);
    this.prune(left);
    this.prune(p);
    return left;
  }

  /** Move the content of `node` before the branch holding `m` into `left`, recursively. */
  private cutInto(left: XmlNode, node: XmlNode, m: XmlNode): void {
    const idx = node.children.findIndex((c) => c === m || this.within(m, c));
    const before = node.children.slice(0, idx).filter((c) => !isPropertyChild(localName(c.name)));
    for (const c of before) {
      this.remove(c);
      this.append(left, c);
    }
    const branch = node.children.find((c) => c === m || this.within(m, c));
    if (!branch || branch === m) return;
    const half = createElement(branch.name, { ...branch.attrs });
    for (const c of branch.children) if (isPropertyChild(localName(c.name))) half.children.push(cloneNode(c));
    this.append(left, half);
    this.cutInto(half, branch, m);
  }

  /**
   * Join `second` onto the end of `first` (Backspace at the start of `second`).
   * The joined paragraph keeps `first`'s properties — unless `first` was
   * empty, when it takes `second`'s, so deleting a blank line never restyles
   * the text below it. A section break on `second`'s mark survives the join.
   */
  joinParagraphs(first: XmlNode, second: XmlNode, firstEmpty: boolean): void {
    const pPr2 = child(second, 'pPr');
    if (firstEmpty) {
      const pPr1 = child(first, 'pPr');
      if (pPr1) this.remove(pPr1);
      if (pPr2) {
        this.remove(pPr2);
        this.insertAt(first, 0, pPr2);
      }
    } else {
      const sect = child(pPr2, 'sectPr');
      if (sect) {
        const pPr1 = this.ensurePPr(first);
        const old = child(pPr1, 'sectPr');
        if (old) this.remove(old);
        this.remove(sect);
        this.insertOrdered(pPr1, sect, PPR_ORDER);
      }
    }
    for (const c of [...second.children]) {
      if (localName(c.name) === 'pPr') continue;
      this.remove(c);
      this.append(first, c);
    }
    this.remove(second);
    this.coalesce(first);
  }

  ensurePPr(p: XmlNode): XmlNode {
    const existing = child(p, 'pPr');
    if (existing) return existing;
    const pPr = createElement(this.q('pPr'));
    this.insertAt(p, 0, pPr);
    return pPr;
  }

  /**
   * Drop what an edit left empty: runs with nothing but properties, and inline
   * containers (hyperlinks, revisions) with nothing inside. Content controls
   * and field structure are never dropped.
   */
  prune(n: XmlNode): void {
    for (const c of [...n.children]) {
      const local = localName(c.name);
      if (OPAQUE.has(local) || local === 'pPr' || local === 'rPr') continue;
      this.prune(c);
      if (local === 'r' && c.children.every((k) => localName(k.name) === 'rPr')) this.remove(c);
      else if (PRUNABLE.has(local) && c.children.length === 0) this.remove(c);
    }
  }

  /**
   * Merge neighbouring runs this operation cut apart when they ended up with
   * identical properties and nothing but text — so that deleting a word does
   * not leave its paragraph in fragments. Untouched runs are left as they are.
   */
  coalesce(n: XmlNode): void {
    for (const c of n.children) if (!OPAQUE.has(localName(c.name)) && localName(c.name) !== 'r') this.coalesce(c);
    for (let i = 0; i + 1 < n.children.length; ) {
      const a = n.children[i]!;
      const b = n.children[i + 1]!;
      if (
        (this.touched.has(a) || this.touched.has(b)) &&
        isPlainTextRun(a) &&
        isPlainTextRun(b) &&
        sameProps(a, b)
      ) {
        const ta = a.children.filter((k) => localName(k.name) === 't');
        const tb = b.children.filter((k) => localName(k.name) === 't');
        const last = ta[ta.length - 1]!;
        last.text += tb.map((t) => t.text).join('');
        preserveSpace(last);
        this.remove(b);
        this.touched.add(a);
        continue;
      }
      i++;
    }
  }

  /** Apply a run format to every run in `nodes` (descending into containers, not into drawings). */
  formatNodes(nodes: XmlNode[], format: RunFormat): void {
    const visit = (n: XmlNode) => {
      const local = localName(n.name);
      if (OPAQUE.has(local) || isRangeMark(local)) return;
      if (local === 'r') {
        let rPr = child(n, 'rPr');
        if (!rPr) {
          rPr = createElement(this.q('rPr'));
          this.insertAt(n, 0, rPr);
        }
        applyFormat(rPr, format, this.prefix);
        this.touched.add(n);
        return;
      }
      for (const c of n.children) visit(c);
    };
    for (const n of nodes) visit(n);
  }

  /** Apply a run format to a paragraph mark (`<w:pPr><w:rPr>`). */
  formatMark(p: XmlNode, format: RunFormat): void {
    const pPr = this.ensurePPr(p);
    let rPr = child(pPr, 'rPr');
    if (!rPr) {
      rPr = createElement(this.q('rPr'));
      this.insertOrdered(pPr, rPr, PPR_ORDER);
    }
    applyFormat(rPr, format, this.prefix);
  }

  /** Remove every marker left under `n`. */
  clearMarkers(n: XmlNode): void {
    n.children = n.children.filter((c) => {
      if (c.name === MARKER) {
        this.parents.delete(c);
        return false;
      }
      this.clearMarkers(c);
      return true;
    });
  }

  /** Set (or with `undefined`, remove) a `<w:pPr>` child with a `w:val`. */
  setParaVal(p: XmlNode, local: string, val: string | undefined): void {
    const pPr = this.ensurePPr(p);
    const old = child(pPr, local);
    if (old) this.remove(old);
    if (val === undefined) return;
    this.insertOrdered(pPr, createElement(this.q(local), { [this.q('val')]: val }), PPR_ORDER);
  }

  /** Merge attributes into a `<w:pPr>` child such as `<w:spacing>` or `<w:ind>`. */
  setParaAttrs(p: XmlNode, local: string, attrs: Record<string, string | undefined>): void {
    const pPr = this.ensurePPr(p);
    let el = child(pPr, local);
    if (!el) {
      el = createElement(this.q(local));
      this.insertOrdered(pPr, el, PPR_ORDER);
    }
    for (const [k, v] of Object.entries(attrs)) {
      const name = this.q(k);
      if (v === undefined) {
        delete el.attrs[name];
        delete el.attrs[k];
      } else setAttr(el, name, v);
    }
  }
}

/** `<w:t>` needs `xml:space="preserve"` whenever its text has outer whitespace, or Word trims it. */
function preserveSpace(t: XmlNode): void {
  if (t.text !== t.text.trim()) t.attrs['xml:space'] = 'preserve';
  else delete t.attrs['xml:space'];
}

/** Revision marks on a paragraph mark's properties belong to that mark, not to new text. */
function stripRevisionProps(rPr: XmlNode): void {
  rPr.children = rPr.children.filter((c) => !['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange'].includes(localName(c.name)));
}

function isPlainTextRun(r: XmlNode): boolean {
  if (localName(r.name) !== 'r') return false;
  if (Object.keys(r.attrs).some((k) => !/:rsid/.test(k))) return false;
  return r.children.every((c) => ['rPr', 't'].includes(localName(c.name)));
}

function sameProps(a: XmlNode, b: XmlNode): boolean {
  const key = (r: XmlNode) => JSON.stringify(child(r, 'rPr') ?? null);
  return key(a) === key(b);
}

/** Atom length of inserted text: every UTF-16 unit, tab and break counts one. */
export function atomLength(text: string): number {
  return text.length;
}

/** The `<w:t>` a paragraph offset falls in, for typing straight into an existing run. */
export function textHostAt(pm: ParaMap, off: number): { seg: Seg; at: number } | undefined {
  const usable = (s: Seg) => s.kind === 'text' && !isLockedRun(s.run) && !s.run.hidden;
  const left = [...pm.segs].reverse().find((s) => s.len > 0 && s.start < off && off <= s.start + s.len);
  if (left && usable(left)) return { seg: left, at: off - left.start };
  // Nothing to continue on the left (a paragraph start, a tab, a picture):
  // type into the start of the text on the right instead.
  if (!left || left.kind !== 'text') {
    const right = pm.segs.find((s) => s.len > 0 && s.start === off);
    if (right && usable(right)) return { seg: right, at: 0 };
  }
  return undefined;
}

export { preserveSpace };
