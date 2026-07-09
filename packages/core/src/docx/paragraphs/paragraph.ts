/**
 * Paragraph parsing: <w:p> -> DocxParagraph. Inline drawings become drawing
 * runs in the text stream; anchored (floating) drawings attach to the paragraph.
 * Resolves the paragraph property cascade, the run base props its runs inherit,
 * and the pre-rendered list marker from numbering state.
 */
import { child, localName, type XmlNode } from '../../oxml/xml.js';
import { twipToPx, twipToPt, halfPtToPt, borderSzToPx } from '../units.js';
import { logicalChildren } from '../content.js';
import { pPrFrom, rPrFrom, mergePara, mergeRun, type ParaProps, type RunProps, type RawBorder } from '../styles/styles.js';
import { parseRunContainer, type FieldState } from './run.js';
import type { DocxBlock, DocxParagraph, DocxRun, DocxAnchor, Stroke, TextAlign } from '../model.js';
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

  // Cascade: docDefaults -> table style -> paragraph style -> direct.
  const styleResolved = ctx.styles.resolveParaProps(styleId);
  const resolved = tableBase
    ? mergePara(mergePara(tableBase.pPr as ParaProps, styleResolved), direct)
    : mergePara(styleResolved, direct);

  // Text runs inherit docDefaults -> (table style) -> paragraph style. Word's
  // paragraph-mark rPr formats only the mark; we fold it in for the body path
  // (long-standing behaviour) but not inside table cells, where doing so would
  // wrongly override a conditional format like firstRow's white text.
  const baseRun = tableBase
    ? mergeRun(tableBase.rPr as RunProps, ctx.styles.resolveParaRunProps(styleId))
    : mergeRun(ctx.styles.resolveParaRunProps(styleId), rPrFrom(child(pPr, 'rPr')));

  const fieldState: FieldState = { inField: false, fieldInstr: '', inSeparate: false };
  const runs: DocxRun[] = [];
  const anchors: DocxAnchor[] = [];
  for (const node of logicalChildren(p)) {
    const name = localName(node.name);
    if (name === 'r' || name === 'hyperlink') runs.push(...parseRunContainer(node, baseRun, ctx, fieldState, undefined, anchors));
  }

  const para: DocxParagraph = {
    kind: 'paragraph',
    runs,
    styleName: ctx.styles.styleName(styleId),
  };

  if (baseRun.font) para.baseFontFamily = baseRun.font;
  if (baseRun.sizeHalfPt !== undefined) para.baseFontSizePt = halfPtToPt(baseRun.sizeHalfPt);
  if (baseRun.bold) para.baseBold = true;
  if (baseRun.italic) para.baseItalic = true;
  if (baseRun.colorHex) para.baseColorHex = baseRun.colorHex;

  if (resolved.align) para.align = resolved.align as TextAlign;

  applyIndents(para, resolved, ctx);
  applySpacing(para, resolved);

  if (resolved.contextualSpacing) para.contextualSpacing = true;
  if (resolved.keepTogether) para.keepTogether = true;
  if (resolved.pageBreakBefore) para.pageBreakBefore = true;
  if (resolved.shadingHex) para.shadingHex = resolved.shadingHex;

  const borders = mapBorders(resolved);
  if (borders) para.paraBorders = borders;

  // List marker (advances numbering counters, in document order).
  if (resolved.numId !== undefined) {
    para.level = resolved.ilvl ?? 0;
    const marker = ctx.numbering.marker(resolved.numId, para.level);
    if (marker) para.listMarker = marker;
  }

  if (anchors.length) para.anchors = anchors;
  return [para];
}

function applyIndents(para: DocxParagraph, r: ParaProps, ctx: ParseContext): void {
  let left = r.indentLeftTwip;
  let firstLine = r.indentFirstLineTwip;
  let hanging = r.hangingTwip;

  // Fall back to the numbering level's indent when the paragraph doesn't set one.
  if (r.numId !== undefined && (left === undefined || hanging === undefined)) {
    const lvl = ctx.numbering.levelIndent(r.numId, r.ilvl ?? 0);
    if (left === undefined && lvl.leftPx !== undefined) para.indentLeftPx = lvl.leftPx;
    if (hanging === undefined && lvl.hangingPx !== undefined) para.indentFirstLinePx = -lvl.hangingPx;
  }

  if (left !== undefined) para.indentLeftPx = twipToPx(left);
  if (r.indentRightTwip !== undefined) para.indentRightPx = twipToPx(r.indentRightTwip);
  if (hanging !== undefined) para.indentFirstLinePx = -twipToPx(hanging);
  else if (firstLine !== undefined) para.indentFirstLinePx = twipToPx(firstLine);
}

function applySpacing(para: DocxParagraph, r: ParaProps): void {
  if (r.spaceBeforeTwip !== undefined) para.spaceBeforePt = twipToPt(r.spaceBeforeTwip);
  if (r.spaceAfterTwip !== undefined) para.spaceAfterPt = twipToPt(r.spaceAfterTwip);
  if (r.line !== undefined) {
    if (r.lineRule === 'auto' || r.lineRule === undefined) para.lineSpacingPct = r.line / 240;
    else para.lineSpacingPt = twipToPt(r.line);
  }
}

function mapBorders(r: ParaProps): DocxParagraph['paraBorders'] {
  if (!r.borders) return undefined;
  const out: NonNullable<DocxParagraph['paraBorders']> = {};
  let any = false;
  for (const side of ['top', 'bottom', 'left', 'right'] as const) {
    const stroke = borderToStroke(r.borders[side]);
    if (stroke) {
      out[side] = stroke;
      any = true;
    }
  }
  return any ? out : undefined;
}

export function borderToStroke(b: RawBorder | undefined): Stroke | undefined {
  if (!b || b.val === 'none' || b.val === 'nil') return undefined;
  return {
    color: { hex: b.colorHex ?? '000000' },
    width: borderSzToPx(b.sz ?? 4),
  };
}
