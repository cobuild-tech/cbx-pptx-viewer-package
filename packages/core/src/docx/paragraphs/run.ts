/**
 * Run parsing: <w:r> (and hyperlink-wrapped runs) -> DocxRun[].
 *
 * A single <w:r> can contain multiple text/break/tab children in document
 * order; we emit one DocxRun per <w:t>, carrying a break/tab flag from any
 * preceding <w:br>/<w:tab> so the renderer reproduces intra-run layout.
 */
import { child, attr, localName, type XmlNode } from '../../oxml/xml.js';
import { halfPtToPt, twipToPt } from '../units.js';
import { logicalChildren } from '../content.js';
import { rPrFrom, mergeRun, type RunProps } from '../styles/styles.js';
import { parseRunDrawing } from '../images/image.js';
import type { DocxRun, DocxAnchor } from '../model.js';
import type { ParseContext } from '../document/context.js';

/**
 * Complex fields (<w:fldChar> begin/separate/end) nest — a table of contents
 * holds one PAGEREF field per entry — so the open fields form a stack. A run
 * belongs to the result of the innermost field that has reached `separate`.
 */
export interface FieldState {
  stack: { instr: string; separated: boolean }[];
}

export function newFieldState(): FieldState {
  return { stack: [] };
}

/** The instruction of the field whose result is currently being read, if any. */
function currentFieldCode(state: FieldState | undefined): string | undefined {
  if (!state) return undefined;
  for (let i = state.stack.length - 1; i >= 0; i--) {
    const f = state.stack[i]!;
    if (f.separated) return f.instr;
  }
  return undefined;
}

/** True while inside a field's instruction (begin..separate): its runs are code, not text. */
function inInstruction(state: FieldState | undefined): boolean {
  const top = state?.stack[state.stack.length - 1];
  return top !== undefined && !top.separated;
}

/** Parse the runs inside a paragraph child (<w:r>, <w:hyperlink> or <w:fldSimple>). */
export function parseRunContainer(
  node: XmlNode,
  baseRun: RunProps,
  ctx: ParseContext,
  fieldState?: FieldState,
  hyperlink?: string,
  anchors?: DocxAnchor[],
): DocxRun[] {
  const name = localName(node.name);
  if (name === 'hyperlink') {
    const rel = ctx.rel(attr(node, 'id'));
    const href = rel?.mode === 'External' ? rel.target : undefined;
    const anchor = attr(node, 'anchor');
    const target = href ?? (anchor ? `#${anchor}` : undefined);
    const out: DocxRun[] = [];
    for (const c of logicalChildren(node, { keepFields: true })) {
      out.push(...parseRunContainer(c, baseRun, ctx, fieldState, target, anchors));
    }
    return out;
  }
  if (name === 'fldSimple') {
    // A simple field is its instruction plus the cached result runs, in one
    // element; treat it as an already-separated complex field.
    const state = fieldState ?? newFieldState();
    state.stack.push({ instr: attr(node, 'instr') ?? '', separated: true });
    const out: DocxRun[] = [];
    for (const c of logicalChildren(node, { keepFields: true })) {
      out.push(...parseRunContainer(c, baseRun, ctx, state, hyperlink, anchors));
    }
    state.stack.pop();
    return out;
  }
  if (name === 'r') return parseRun(node, baseRun, ctx, fieldState, hyperlink, anchors);
  return [];
}

interface Pending {
  breakBefore: boolean;
  pageBreakBefore: boolean;
  columnBreak?: boolean;
  tabBefore: boolean;
  ptab?: DocxRun['ptab'];
}

function parseRun(
  r: XmlNode,
  baseRun: RunProps,
  ctx: ParseContext,
  fieldState?: FieldState,
  hyperlink?: string,
  anchors?: DocxAnchor[],
): DocxRun[] {
  const props = mergeRun(baseRun, resolveRunStyle(r, ctx));
  props.font = ctx.styles.fontName(props.font);
  props.eastAsiaFont = ctx.styles.fontName(props.eastAsiaFont);
  props.csFont = ctx.styles.fontName(props.csFont);
  const out: DocxRun[] = [];
  const rPrChange = child(child(r, 'rPr'), 'rPrChange');
  const pending: Pending = { breakBefore: false, pageBreakBefore: false, tabBefore: false };

  const emit = (text: string, source?: XmlNode): DocxRun => {
    const run = makeRun(text, props, hyperlink, pending);
    const code = currentFieldCode(fieldState);
    if (code !== undefined) run.fieldCode = code;
    out.push(run);
    // The <w:t> identifies the run; the <w:r> carries its <w:rPr>.
    if (source) ctx.recordSource?.(run, source, r);
    if (rPrChange) {
      run.formatChange = { author: attr(rPrChange, 'author') ?? '' };
      const date = attr(rPrChange, 'date');
      if (date) run.formatChange.date = date;
    }
    pending.breakBefore = false;
    pending.pageBreakBefore = false;
    pending.tabBefore = false;
    delete pending.ptab;
    delete pending.columnBreak;
    return run;
  };

  for (const node of r.children) {
    const tagName = localName(node.name);
    switch (tagName) {
      case 'fldChar': {
        const type = attr(node, 'fldCharType');
        if (!fieldState) break;
        if (type === 'begin') fieldState.stack.push({ instr: '', separated: false });
        else if (type === 'separate') {
          const top = fieldState.stack[fieldState.stack.length - 1];
          if (top) top.separated = true;
        } else if (type === 'end') fieldState.stack.pop();
        break;
      }
      case 'instrText': {
        const top = fieldState?.stack[fieldState.stack.length - 1];
        if (top && !top.separated) top.instr += node.text ?? '';
        break;
      }
      case 'br': {
        // A page/column break ends the page; every other kind is a line break.
        // Both keep breakBefore, which is what editing round-trips.
        const type = attr(node, 'type');
        pending.breakBefore = true;
        if (type === 'page' || type === 'column') pending.pageBreakBefore = true;
        if (type === 'column') pending.columnBreak = true;
        break;
      }
      case 'cr':
        pending.breakBefore = true;
        break;
      case 'tab':
        pending.tabBefore = true;
        break;
      case 'ptab': {
        const align = attr(node, 'alignment');
        pending.tabBefore = true;
        pending.ptab = align === 'center' || align === 'right' ? align : 'left';
        break;
      }
      case 't': {
        // Text inside a field's instruction part is code, never displayed.
        if (inInstruction(fieldState)) break;
        emit(node.text ?? '', node);
        break;
      }
      case 'delText':
        // Deleted text (inside <w:del>): shown struck through, never edited.
        if (!inInstruction(fieldState)) emit(node.text ?? '');
        break;
      case 'commentReference': {
        const comment = ctx.comment?.(attr(node, 'id') ?? '');
        if (comment) emit('').comment = comment;
        break;
      }
      case 'footnoteReference':
      case 'endnoteReference': {
        const kind = tagName === 'footnoteReference' ? 'footnote' : 'endnote';
        const note = ctx.note?.(kind, attr(node, 'id') ?? '');
        if (!note) break;
        // customMarkFollows: the run's own text is the mark, so no number.
        const custom = attr(node, 'customMarkFollows');
        const run = emit(custom && custom !== '0' && custom !== 'false' ? '' : note.number);
        run.note = { kind, blocks: note.blocks };
        break;
      }
      case 'footnoteRef':
      case 'endnoteRef':
        // Inside the note itself: the same number its reference shows.
        if (ctx.noteNumber !== undefined) emit(ctx.noteNumber);
        break;
      case 'separator':
      case 'continuationSeparator':
        emit('').separator = true;
        break;
      case 'noBreakHyphen':
        if (!inInstruction(fieldState)) emit('‑');
        break;
      case 'softHyphen':
        if (!inInstruction(fieldState)) emit('­');
        break;
      case 'sym': {
        const code = parseInt(attr(node, 'char') ?? '', 16);
        if (Number.isFinite(code) && !inInstruction(fieldState)) {
          emit(symbolChar(code, attr(node, 'font')));
        }
        break;
      }
      case 'drawing':
      case 'pict':
      case 'object':
      case 'AlternateContent': {
        const res = parseRunDrawing(node, ctx);
        if (res?.inline) {
          const run = emit('');
          run.drawing = res.inline;
        } else if (res?.anchor) {
          anchors?.push(res.anchor);
        }
        break;
      }
      default:
        break;
    }
  }

  // A run that is only a <w:br>/<w:tab> with no text still needs to emit the break.
  if (out.length === 0 && (pending.breakBefore || pending.tabBefore)) emit('');
  return out;
}

function resolveRunStyle(r: XmlNode, ctx: ParseContext): RunProps {
  const rPr = child(r, 'rPr');
  const charStyleId = attr(child(rPr, 'rStyle'), 'val');
  const fromStyle = charStyleId ? ctx.styles.resolveCharStyle(charStyleId) : {};
  return mergeRun(fromStyle, rPrFrom(rPr));
}

function makeRun(text: string, props: RunProps, hyperlink: string | undefined, pending: Pending): DocxRun {
  const run: DocxRun = { text };
  // A right-to-left run is complex-script text, set in the complex-script
  // font, size and weight (<w:rFonts w:cs>, <w:szCs>, <w:bCs>).
  const cs = props.rtl ? props : undefined;
  if (cs?.csFont) props = { ...props, font: cs.csFont };
  if (cs?.csSizeHalfPt !== undefined) props = { ...props, sizeHalfPt: cs.csSizeHalfPt };
  if (cs?.csBold !== undefined) props = { ...props, bold: cs.csBold };
  if (cs?.csItalic !== undefined) props = { ...props, italic: cs.csItalic };
  if (props.rtl) run.rtl = true;
  if (props.eastAsiaFont && props.eastAsiaFont !== props.font) run.eastAsiaFont = props.eastAsiaFont;
  if (props.bold) run.bold = true;
  if (props.italic) run.italic = true;
  if (props.underline) run.underline = true;
  if (props.strike) run.strike = true;
  if (props.sizeHalfPt !== undefined) run.sizePt = halfPtToPt(props.sizeHalfPt);
  if (props.colorHex) run.colorHex = props.colorHex;
  if (props.highlightHex) run.highlightHex = props.highlightHex;
  if (props.font) run.font = props.font;
  if (props.vertAlign) run.vertAlign = props.vertAlign;
  if (props.caps) run.caps = props.caps;
  if (props.hidden) run.hidden = true;
  if (props.charSpacingTwip) run.charSpacingPt = twipToPt(props.charSpacingTwip);
  // <w:kern w:val="0"/> means no kerning; otherwise kern from that size up.
  if (props.kernHalfPt && (props.sizeHalfPt ?? 20) >= props.kernHalfPt) run.kern = true;
  if (hyperlink) run.hyperlink = hyperlink;
  if (pending.breakBefore) run.breakBefore = true;
  if (pending.pageBreakBefore) run.pageBreakBefore = true;
  if (pending.columnBreak) run.columnBreak = true;
  if (pending.tabBefore) run.tabBefore = true;
  if (pending.ptab) run.ptab = pending.ptab;
  return run;
}

/** <w:sym>: a character from a symbol font, mapped to Unicode where we know it. */
function symbolChar(code: number, font: string | undefined): string {
  const slot = code & 0xff;
  const f = font?.toLowerCase();
  if (f === 'symbol') {
    const map: Record<number, string> = { 0xb7: '•', 0xae: '→', 0xde: '⇒', 0xd8: '¬', 0xb1: '±', 0xb4: '×', 0xb8: '÷', 0xa3: '≤', 0xb3: '≥', 0xb9: '≠', 0xa5: '∞', 0xb0: '°' };
    return map[slot] ?? String.fromCharCode(slot);
  }
  if (f?.startsWith('wingdings')) {
    const map: Record<number, string> = { 0xfc: '✓', 0xfb: '✗', 0xfe: '☑', 0xa8: '☐', 0x6f: '☐', 0x78: '☒', 0x6e: '■', 0x6c: '●', 0xa7: '▪', 0xd8: '➢', 0xe8: '➔', 0x28: '☎', 0x2a: '✉' };
    return map[slot] ?? '•';
  }
  return String.fromCodePoint(code >= 0xf000 && code <= 0xf0ff ? slot : code);
}
