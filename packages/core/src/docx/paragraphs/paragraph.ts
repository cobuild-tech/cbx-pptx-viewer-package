/**
 * Paragraph parser for WordprocessingML.
 *
 * Converts a <w:p> element into a DocxParagraph IR node, resolving style
 * inheritance, list numbering, hyperlinks, and inline images.
 */
import { child, attr, localName, type XmlNode } from '../../oxml/xml.js';
import { twipsToPx } from '../units.js';
import type { DocxParagraph, DocxBlock, DocxRun } from '../model.js';
import type { Stroke } from '../model.js';
import { StyleMap, mergeParaProps, mergeRunProps } from '../styles/styles.js';
import type { ParaBorderSide } from '../styles/styles.js';
import { NumberingMap } from '../numbering/numbering.js';
import { parseRun, runHasPageBreak } from './run.js';
import { parseDrawing } from '../images/image.js';
import { encodeNodeId } from '../edit/nodeId.js';

export interface ParagraphParseCtx {
  styles: StyleMap;
  numbering: NumberingMap;
  /** Counter state for auto-numbered lists; mutated during parsing. */
  listCounters: Map<string, number>;
  resolveImage: (relId: string) => string | undefined;
  resolveHyperlink: (relId: string) => string | undefined;
  /** OPC part this content lives in (for nodeId stamping in editable mode). */
  partPath?: string;
}

export interface ParagraphResult {
  paragraphs: DocxParagraph[];
  /** True if the paragraph ends with an explicit page break. */
  endsWithPageBreak: boolean;
}

/**
 * Parse a <w:p> element. May return multiple DocxParagraphs if the paragraph
 * contains an explicit page break mid-text (rare but possible).
 */
export function parseParagraph(
  pEl: XmlNode,
  ctx: ParagraphParseCtx,
  /** Child-index path to this `<w:p>` from the part root (editable mode). */
  path?: number[],
): ParagraphResult {
  const nodeId = path && ctx.partPath ? encodeNodeId(ctx.partPath, path) : undefined;
  const runNodeId = (...suffix: number[]): string | undefined =>
    path && ctx.partPath ? encodeNodeId(ctx.partPath, [...path, ...suffix]) : undefined;
  const pPr = child(pEl, 'pPr');

  // Resolve base style.
  const styleRef =
    attr(child(pPr, 'pStyle'), 'w:val') ?? attr(child(pPr, 'pStyle'), 'val') ?? 'Normal';
  const baseStyle = ctx.styles.get(styleRef);

  // Merge paragraph properties onto base.
  const para = { ...baseStyle.para };
  if (pPr) mergeParaProps(para, pPr);

  // Resolve list numbering.
  const numId = para.numId;
  const ilvl = para.ilvl ?? 0;
  let bullet = undefined;
  let listIndentLeftPx: number | undefined;
  let listIndentFirstLinePx: number | undefined;

  if (numId !== undefined && numId !== 0) {
    const level = ctx.numbering.resolve(numId, ilvl);
    if (level) {
      bullet = ctx.numbering.toBullet(level, ctx.listCounters, numId, ilvl);
      listIndentLeftPx = level.indentLeftPx;
      listIndentFirstLinePx = level.indentFirstLinePx;
    }
  }

  // Run-level base style from the paragraph style.
  const baseRun = { ...baseStyle.run };
  // Merge paragraph's rPr default run props (pPr/rPr).
  const pRpr = child(pPr, 'rPr');
  if (pRpr) mergeRunProps(baseRun, pRpr);

  // Parse runs, hyperlinks, and drawing elements.
  const runs: DocxRun[] = [];
  let endsWithPageBreak = false;
  const paragraphs: DocxParagraph[] = [];

  // Check for paragraph-level page break.
  const pageBreakBeforeEl = child(pPr, 'pageBreakBefore');
  let pageBreakBefore = pageBreakBeforeEl !== undefined;
  if (pageBreakBeforeEl) {
    const val = attr(pageBreakBeforeEl, 'w:val') ?? attr(pageBreakBeforeEl, 'val');
    pageBreakBefore = val === undefined || (val !== '0' && val !== 'false');
  }

  for (let childIdx = 0; childIdx < pEl.children.length; childIdx++) {
    const node = pEl.children[childIdx]!;
    const name = localName(node.name);

    if (name === 'r') {
      // Check for a page break inside this run — flush what we have and signal.
      if (runHasPageBreak(node)) {
        // Flush current runs as a paragraph before the break.
        if (runs.length > 0) {
          paragraphs.push(buildParagraph(runs.slice(), baseStyle.name, para, baseRun, bullet, ilvl, pageBreakBefore, listIndentLeftPx, listIndentFirstLinePx, nodeId));
          runs.length = 0;
          pageBreakBefore = false;
          bullet = undefined;
        }
        endsWithPageBreak = true;
        continue;
      }
      const run = parseRun(node, baseRun, ctx.resolveImage, ctx.styles) as DocxRun | null;
      if (run) {
        run.nodeId = runNodeId(childIdx);
        runs.push(run);
      }

    } else if (name === 'hyperlink') {
      const relId = attr(node, 'r:id') ?? attr(node, 'id');
      const url = relId ? ctx.resolveHyperlink(relId) : undefined;
      for (let hIdx = 0; hIdx < node.children.length; hIdx++) {
        const rEl = node.children[hIdx]!;
        if (localName(rEl.name) !== 'r') continue;
        const run = parseRun(rEl, baseRun, ctx.resolveImage, ctx.styles) as DocxRun | null;
        if (run) {
          if (url) run.hyperlink = url;
          run.nodeId = runNodeId(childIdx, hIdx);
          runs.push(run);
        }
      }

    } else if (name === 'drawing') {
      const image = parseDrawing(node, ctx.resolveImage);
      if (image) {
        // Flush preceding text runs as a paragraph, then the image is its own block.
        // For simplicity, inline images are appended as a run with a special marker.
        // We embed a placeholder text run with image metadata via a custom approach:
        // wrap the image inline with the surrounding text by just noting it.
        // Real approach: push current runs, push image block separately.
        // Since DocxParagraph only has runs, we add the image as a top-level block
        // outside this paragraph. Signal via side-channel? Instead we return the
        // paragraphs array and include inline images as separate blocks.
        // We flush runs so far, then the image block follows outside this function.
        // → handled at body level by checking for 'drawing' nodes during paragraph iteration.
      }
    }
  }

  paragraphs.push(buildParagraph(runs, baseStyle.name, para, baseRun, bullet, ilvl, pageBreakBefore, listIndentLeftPx, listIndentFirstLinePx, nodeId));

  return { paragraphs, endsWithPageBreak };
}

function buildParagraph(
  runs: DocxRun[],
  styleName: string,
  para: ReturnType<typeof import('../styles/styles.js').StyleMap.prototype.get>['para'],
  baseRun: ReturnType<typeof import('../styles/styles.js').StyleMap.prototype.get>['run'],
  bullet: DocxParagraph['bullet'],
  level: number,
  pageBreakBefore: boolean,
  listIndentLeftPx: number | undefined,
  listIndentFirstLinePx: number | undefined,
  nodeId: string | undefined,
): DocxParagraph {
  return {
    kind: 'paragraph',
    runs,
    nodeId,
    styleName,
    baseFontFamily: baseRun.fontAscii ?? baseRun.fontHAnsi,
    baseFontSizePt: baseRun.sizePt,
    baseBold: baseRun.bold,
    baseItalic: baseRun.italic,
    baseColorHex: baseRun.colorHex,
    align: para.align,
    indentLeftPx: listIndentLeftPx ?? para.indentLeftPx,
    indentFirstLinePx: listIndentFirstLinePx ?? para.indentFirstLinePx,
    spaceBeforePt: para.spaceBeforePt,
    spaceAfterPt: para.spaceAfterPt,
    lineSpacingPct: para.lineSpacingPct,
    lineSpacingPt: para.lineSpacingPt,
    bullet,
    level: bullet ? level : undefined,
    pageBreakBefore: pageBreakBefore || undefined,
    shadingHex: para.shadingHex,
    indentRightPx: para.indentRightPx,
    contextualSpacing: para.contextualSpacing,
    paraBorders: para.pBdr ? buildParaBorders(para.pBdr) : undefined,
  };
}

function buildParaBorders(
  pBdr: Partial<Record<'top' | 'bottom' | 'left' | 'right', ParaBorderSide>>,
): DocxParagraph['paraBorders'] {
  const out: NonNullable<DocxParagraph['paraBorders']> = {};
  for (const side of ['top', 'bottom', 'left', 'right'] as const) {
    const s = pBdr[side];
    if (!s) continue;
    const isDashed = s.type === 'dashed' || s.type === 'dotted';
    const stroke: Stroke = {
      color: { hex: s.colorHex },
      width: s.widthPx,
      ...(isDashed ? { dash: [4, 3] } : {}),
    };
    out[side] = stroke;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}
