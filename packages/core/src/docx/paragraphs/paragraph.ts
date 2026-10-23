/**
 * Paragraph parsing: <w:p> -> DocxParagraph. Inline drawings become drawing
 * runs in the text stream; anchored (floating) drawings attach to the paragraph.
 * Resolves the paragraph property cascade, the run base props its runs inherit,
 * and the pre-rendered list marker from numbering state.
 */
import { child, children, attr, localName, type XmlNode } from '../../oxml/xml.js';
import { ommlToMath } from '../math/omml.js';
import { twipToPx, twipToPt, halfPtToPt, borderSzToPx } from '../units.js';
import { logicalChildren } from '../content.js';
import { pPrFrom, rPrFrom, mergePara, mergeRun, type ParaProps, type RunProps, type RawBorder } from '../styles/styles.js';
import { parseRunContainer, newFieldState, type FieldState } from './run.js';
import { ptToPx } from '../../oxml/units.js';
import type {
  DocxBlock,
  DocxParagraph,
  DocxRun,
  DocxAnchor,
  DocxMarkerFormat,
  DocxTabStop,
  DocxStroke,
  DocxRevision,
  TextAlign,
} from '../model.js';
import type { ParseContext } from '../document/context.js';

/**
 * Formatting a table style contributes to the paragraphs in a cell. It sits
 * BELOW the paragraph's own style and direct props in the cascade (Word applies
 * table styles before paragraph styles), so a cell paragraph can override it.
 */
export interface TableBase {
  pPr: Partial<ParaProps>;
  rPr: Partial<RunProps>;
}

export function parseParagraph(p: XmlNode, ctx: ParseContext, tableBase?: TableBase): DocxBlock[] {
  const pPr = child(p, 'pPr');
  const direct = pPrFrom(pPr);
  const styleId = direct.styleId;

  // Cascade: docDefaults + Normal -> table style -> the paragraph's own style
  // chain -> direct. A table style outranks Normal but not a named style.
  const layers = ctx.styles.paraLayers(styleId);
  const below = tableBase ? mergePara(layers.basePara, tableBase.pPr) : layers.basePara;
  const styled = mergePara(below, layers.ownPara);
  styled.styleId = styleId;
  const resolved = mergePara(styled, direct);

  // Text runs inherit the same layers. The paragraph mark's own rPr
  // (<w:pPr><w:rPr>) formats only the mark — never the runs — so it is
  // layered in only below, where the mark is all there is.
  const runBelow = tableBase ? mergeRun(layers.baseRun, tableBase.rPr) : layers.baseRun;
  const baseRun = mergeRun(runBelow, layers.ownRun);
  const markRun = mergeRun(baseRun, rPrFrom(child(pPr, 'rPr')));

  const fieldState: FieldState = newFieldState();
  const runs: DocxRun[] = [];
  const anchors: DocxAnchor[] = [];
  let mathDisplay: DocxParagraph['mathDisplay'];
  const open = ctx.openComments;
  /** Parse one run container, tagging its runs with a revision and open comments. */
  const take = (node: XmlNode, revision?: DocxRevision) => {
    for (const run of parseRunContainer(node, baseRun, ctx, fieldState, undefined, anchors)) {
      if (revision) run.revision = revision;
      if (open?.size) run.commentIds = [...open];
      runs.push(run);
    }
  };
  for (const node of logicalChildren(p, { keepFields: true, keepRevisions: true })) {
    const name = localName(node.name);
    if (name === 'r' || name === 'hyperlink' || name === 'fldSimple') {
      take(node);
    } else if (name === 'ins' || name === 'del' || name === 'moveTo' || name === 'moveFrom') {
      const revision: DocxRevision = {
        kind: name === 'del' || name === 'moveFrom' ? 'del' : 'ins',
        author: attr(node, 'author') ?? '',
      };
      const date = attr(node, 'date');
      if (date) revision.date = date;
      for (const c of logicalChildren(node, { keepFields: true })) {
        const cn = localName(c.name);
        if (cn === 'r' || cn === 'hyperlink' || cn === 'fldSimple') take(c, revision);
      }
    } else if (name === 'commentRangeStart') {
      open?.add(attr(node, 'id') ?? '');
    } else if (name === 'commentRangeEnd') {
      open?.delete(attr(node, 'id') ?? '');
    } else if (name === 'oMath') {
      const run: DocxRun = { text: '', math: ommlToMath(node) };
      ctx.recordSource?.(run, node);
      runs.push(run);
    } else if (name === 'oMathPara') {
      // A display equation: each <m:oMath> on its own line, centred by default.
      const jc = attr(child(child(node, 'oMathParaPr'), 'jc'), 'val');
      mathDisplay = { align: jc === 'left' ? 'left' : jc === 'right' ? 'right' : 'center' };
      children(node, 'oMath').forEach((m, i) => {
        const run: DocxRun = { text: '', math: ommlToMath(m), ...(i > 0 ? { breakBefore: true } : {}) };
        ctx.recordSource?.(run, m);
        runs.push(run);
      });
    }
  }

  const para: DocxParagraph = {
    kind: 'paragraph',
    runs,
    styleName: ctx.styles.styleName(styleId),
  };

  // The paragraph's font and size are its mark's. Word sizes each line by the
  // text on it, the mark counting on the last line only; a block's own font
  // sets a minimum height for every line in CSS, so it must be the mark's and
  // never a style font no run uses. Runs always carry their own font and size.
  const isEmpty = !runs.some((r) => (r.text && !r.hidden) || r.drawing || r.math);
  const base = isEmpty ? markRun : baseRun;
  const markFont = ctx.styles.fontName(markRun.font);
  if (markFont) para.baseFontFamily = markFont;
  if (markRun.sizeHalfPt !== undefined) para.baseFontSizePt = halfPtToPt(markRun.sizeHalfPt);
  if (base.bold) para.baseBold = true;
  if (base.italic) para.baseItalic = true;
  if (base.colorHex) para.baseColorHex = base.colorHex;

  if (resolved.align) para.align = resolved.align as TextAlign;
  if (resolved.bidi) para.rtl = true;
  if (mathDisplay) para.mathDisplay = mathDisplay;

  const numId = resolved.numId !== undefined && resolved.numId !== 0 ? resolved.numId : undefined;
  applyIndents(para, styled, direct, numId, resolved.ilvl ?? 0, ctx);
  applySpacing(para, resolved);
  applyTabs(para, resolved, ctx);

  if (resolved.contextualSpacing) para.contextualSpacing = true;
  if (resolved.keepNext) para.keepNext = true;
  if (resolved.keepLines) para.keepLines = true;
  // Word treats an unstated widowControl as on (documents write
  // <w:widowControl w:val="0"/> to turn it off), whatever the spec's default.
  if (resolved.widowControl ?? true) para.widowControl = true;
  if (resolved.pageBreakBefore) para.pageBreakBefore = true;
  if (resolved.shadingHex) para.shadingHex = resolved.shadingHex;

  const borders = mapBorders(resolved);
  if (borders) {
    para.paraBorders = borders.strokes;
    para.paraBorderSpacePx = borders.space;
  }

  // List marker (advances numbering counters, in document order).
  if (numId !== undefined) {
    para.level = resolved.ilvl ?? 0;
    const marker = ctx.numbering.marker(numId, para.level);
    if (marker?.text) {
      para.listMarker = marker.text;
      para.listSuffix = marker.suffix;
      // The marker is drawn in the paragraph mark's formatting, overlaid by the
      // level's own rPr (a symbol-font bullet arrives with its font dropped, so
      // it falls back to the mark's text font).
      const fmt = mergeRun(markRun, marker.rPr);
      const markerFormat: DocxMarkerFormat = {};
      const font = ctx.styles.fontName(fmt.font);
      if (font) markerFormat.font = font;
      if (marker.metricsFont) markerFormat.metricsFont = marker.metricsFont;
      if (fmt.sizeHalfPt !== undefined) markerFormat.sizePt = halfPtToPt(fmt.sizeHalfPt);
      if (fmt.bold) markerFormat.bold = true;
      if (fmt.italic) markerFormat.italic = true;
      if (fmt.colorHex) markerFormat.colorHex = fmt.colorHex;
      para.listMarkerFormat = markerFormat;
    }
  }

  if (anchors.length) para.anchors = anchors;
  ctx.recordSource?.(para, p);
  return [para];
}

/**
 * Indents cascade per layer — paragraph style, then the list level, then direct
 * formatting — and within a layer `hanging` and `firstLine` exclude each other:
 * a layer that states one discards the other from the layers below.
 */
function applyIndents(
  para: DocxParagraph,
  styled: ParaProps,
  direct: Partial<ParaProps>,
  numId: number | undefined,
  ilvl: number,
  ctx: ParseContext,
): void {
  let left = styled.indentLeftTwip;
  let right = styled.indentRightTwip;
  let firstLine = styled.indentFirstLineTwip;
  let hanging = styled.hangingTwip;
  if (hanging !== undefined) firstLine = undefined;

  if (numId !== undefined) {
    const lvl = ctx.numbering.levelIndent(numId, ilvl);
    if (lvl.leftPx !== undefined) left = lvl.leftPx * 15;
    if (lvl.hangingPx !== undefined) {
      hanging = lvl.hangingPx * 15;
      firstLine = undefined;
    }
  }

  if (direct.indentLeftTwip !== undefined) left = direct.indentLeftTwip;
  if (direct.indentRightTwip !== undefined) right = direct.indentRightTwip;
  if (direct.hangingTwip !== undefined) {
    hanging = direct.hangingTwip;
    firstLine = undefined;
  } else if (direct.indentFirstLineTwip !== undefined) {
    firstLine = direct.indentFirstLineTwip;
    hanging = undefined;
  }

  if (left !== undefined) para.indentLeftPx = twipToPx(left);
  if (right !== undefined) para.indentRightPx = twipToPx(right);
  if (hanging !== undefined) para.indentFirstLinePx = -twipToPx(hanging);
  else if (firstLine !== undefined) para.indentFirstLinePx = twipToPx(firstLine);
}

function applySpacing(para: DocxParagraph, r: ParaProps): void {
  if (r.spaceBeforeTwip !== undefined) para.spaceBeforePt = twipToPt(r.spaceBeforeTwip);
  if (r.spaceAfterTwip !== undefined) para.spaceAfterPt = twipToPt(r.spaceAfterTwip);
  if (r.line !== undefined) {
    if (r.lineRule === 'auto' || r.lineRule === undefined) para.lineSpacingPct = r.line / 240;
    else {
      para.lineSpacingPt = twipToPt(r.line);
      para.lineSpacingRule = r.lineRule === 'exact' ? 'exact' : 'atLeast';
    }
  }
}

function applyTabs(para: DocxParagraph, r: ParaProps, ctx: ParseContext): void {
  para.defaultTabPx = ctx.settings.defaultTabPx;
  const tabs: DocxTabStop[] = [];
  for (const t of r.tabs ?? []) {
    if (t.val === 'clear') continue;
    const align = TAB_ALIGN[t.val] ?? 'left';
    const stop: DocxTabStop = { posPx: twipToPx(t.posTwip), align };
    const leader = TAB_LEADER[t.leader ?? ''];
    if (leader) stop.leader = leader;
    tabs.push(stop);
  }
  if (tabs.length) para.tabs = tabs.sort((a, b) => a.posPx - b.posPx);
}

const TAB_ALIGN: Record<string, DocxTabStop['align']> = {
  left: 'left',
  start: 'left',
  center: 'center',
  right: 'right',
  end: 'right',
  decimal: 'decimal',
  bar: 'bar',
  num: 'left',
};

const TAB_LEADER: Record<string, DocxTabStop['leader']> = {
  dot: 'dot',
  hyphen: 'hyphen',
  underscore: 'underscore',
  middleDot: 'middleDot',
  heavy: 'heavy',
};

/**
 * <w:contextualSpacing>: "don't add space between paragraphs of the same
 * style". A paragraph with it drops its space-after when the next paragraph has
 * the same style, and its space-before when the previous one does. Needs the
 * neighbours, so it runs over a finished block list.
 */
export function applyContextualSpacing(blocks: DocxBlock[]): void {
  for (let i = 0; i < blocks.length; i++) {
    const p = blocks[i]!;
    if (p.kind !== 'paragraph' || !p.contextualSpacing) continue;
    const prev = blocks[i - 1];
    const next = blocks[i + 1];
    if (next?.kind === 'paragraph' && next.styleName === p.styleName) p.spaceAfterPt = 0;
    if (prev?.kind === 'paragraph' && prev.styleName === p.styleName) p.spaceBeforePt = 0;
  }
}

function mapBorders(
  r: ParaProps,
): { strokes: NonNullable<DocxParagraph['paraBorders']>; space: NonNullable<DocxParagraph['paraBorderSpacePx']> } | undefined {
  if (!r.borders) return undefined;
  const strokes: NonNullable<DocxParagraph['paraBorders']> = {};
  const space: NonNullable<DocxParagraph['paraBorderSpacePx']> = {};
  let any = false;
  for (const side of ['top', 'bottom', 'left', 'right'] as const) {
    const stroke = borderToStroke(r.borders[side]);
    if (stroke) {
      strokes[side] = stroke;
      // <w:space> is the gap between border and text, in points.
      const pt = r.borders[side]?.spacePt;
      if (pt) space[side] = ptToPx(pt);
      any = true;
    }
  }
  return any ? { strokes, space } : undefined;
}

export function borderToStroke(b: RawBorder | undefined): DocxStroke | undefined {
  if (!b || b.val === 'none' || b.val === 'nil') return undefined;
  const stroke: DocxStroke = {
    color: { hex: b.colorHex ?? '000000' },
    width: borderSzToPx(b.sz ?? 4),
  };
  const val = b.val ?? 'single';
  if (/^(double|triple|thinThick|thickThin)/i.test(val)) stroke.style = 'double';
  else if (/dash/i.test(val)) stroke.style = 'dashed';
  else if (/dot/i.test(val)) stroke.style = 'dotted';
  return stroke;
}
