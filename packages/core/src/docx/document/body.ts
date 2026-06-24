/**
 * Body parser and paginator for WordprocessingML.
 *
 * Two-phase approach:
 *  1. Walk <w:body> and parse every element into a flat block list, tagging
 *     each block with any explicit page-break intent.
 *  2. Run a height-accumulating paginator: estimate each block's rendered
 *     height and flush to a new page when the content area would overflow.
 *
 * Height estimation uses resolved font sizes, paragraph spacing, and a
 * character-width heuristic for line wrapping. It won't be pixel-perfect
 * (that would require a full layout engine), but it distributes content
 * across the correct number of pages for typical documents.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { twipsToPx } from '../units.js';
import { OpcPackage } from '../../oxml/package.js';
import type { DocxBlock, DocxPage, DocxPageSize, DocxPageMargins, DocxParagraph } from '../model.js';
import { StyleMap } from '../styles/styles.js';
import { NumberingMap } from '../numbering/numbering.js';
import { parseParagraph } from '../paragraphs/paragraph.js';
import { parseTable } from '../tables/table.js';
import { parseDrawing } from '../images/image.js';

// Default: US Letter 8.5" × 11" at 96 DPI.
const DEFAULT_PG_SZ: DocxPageSize = { wPx: twipsToPx(12240), hPx: twipsToPx(15840) };
const DEFAULT_PG_MARGINS: DocxPageMargins = {
  topPx: twipsToPx(1440),
  rightPx: twipsToPx(1440),
  bottomPx: twipsToPx(1440),
  leftPx: twipsToPx(1440),
  headerPx: twipsToPx(720),
  footerPx: twipsToPx(720),
};

export interface BodyBuildCtx {
  pkg: OpcPackage;
  docPart: string;
  styles: StyleMap;
  numbering: NumberingMap;
}

// ─── Tagged block (internal) ──────────────────────────────────────────────────

interface TaggedBlock {
  block: DocxBlock;
  /** Force a page break before this block. */
  breakBefore: boolean;
  /** Force a page break after this block. */
  breakAfter: boolean;
  /** When a section change applies after this block. */
  nextSize?: DocxPageSize;
  nextMargins?: DocxPageMargins;
  nextHeader?: DocxParagraph[];
  nextFooter?: DocxParagraph[];
}

// ─── Entry point ─────────────────────────────────────────────────────────────

export function buildPages(bodyEl: XmlNode, ctx: BodyBuildCtx): DocxPage[] {
  const tagged = collectBlocks(bodyEl, ctx);
  return paginate(tagged);
}

// ─── Phase 1: collect all blocks with break metadata ─────────────────────────

function collectBlocks(bodyEl: XmlNode, ctx: BodyBuildCtx): TaggedBlock[] {
  const out: TaggedBlock[] = [];
  const listCounters = new Map<string, number>();

  function resolveImage(relId: string): string | undefined {
    return ctx.pkg.resolveRel(ctx.docPart, relId)?.target;
  }
  function resolveHyperlink(relId: string): string | undefined {
    const rel = ctx.pkg.resolveRel(ctx.docPart, relId);
    return rel?.mode === 'External' ? rel.rawTarget : undefined;
  }

  const paraCtx = { styles: ctx.styles, numbering: ctx.numbering, listCounters, resolveImage, resolveHyperlink };

  for (const node of bodyEl.children) {
    const name = localName(node.name);

    if (name === 'p') {
      const pPr = child(node, 'pPr');
      const sectPr = child(pPr, 'sectPr');

      const { paragraphs, endsWithPageBreak } = parseParagraph(node, paraCtx);

      // Inline drawings: extract before the paragraph text blocks.
      for (const childNode of node.children) {
        if (localName(childNode.name) === 'r') {
          for (const gc of childNode.children) {
            if (localName(gc.name) === 'drawing') {
              const img = parseDrawing(gc, resolveImage);
              if (img) out.push({ block: img, breakBefore: false, breakAfter: false });
            }
          }
        }
      }

      for (let i = 0; i < paragraphs.length; i++) {
        const para = paragraphs[i]!;
        const isLast = i === paragraphs.length - 1;
        out.push({
          block: para,
          breakBefore: i === 0 && !!para.pageBreakBefore,
          breakAfter: isLast && endsWithPageBreak,
        });
      }

      // Section break in pPr → tag the last added block with next-section info.
      if (sectPr && out.length > 0) {
        const { size, margins } = readSectPr(sectPr);
        const sectType = attr(child(sectPr, 'type'), 'w:val') ?? attr(child(sectPr, 'type'), 'val') ?? 'nextPage';
        const { header, footer } = readHeaderFooter(sectPr, ctx);
        const last = out[out.length - 1]!;
        if (sectType !== 'continuous') {
          last.breakAfter = true;
        }
        last.nextSize = size;
        last.nextMargins = margins;
        last.nextHeader = header;
        last.nextFooter = footer;
      }

    } else if (name === 'tbl') {
      out.push({ block: parseTable(node, paraCtx), breakBefore: false, breakAfter: false });

    } else if (name === 'sectPr') {
      // Final body-level sectPr: tag the last block (or create a sentinel).
      const { size, margins } = readSectPr(node);
      const { header, footer } = readHeaderFooter(node, ctx);
      if (out.length > 0) {
        const last = out[out.length - 1]!;
        last.nextSize = size;
        last.nextMargins = margins;
        last.nextHeader = header;
        last.nextFooter = footer;
      }
    }
  }

  return out;
}

// ─── Phase 2: height-accumulating paginator ───────────────────────────────────

function paginate(tagged: TaggedBlock[]): DocxPage[] {
  const pages: DocxPage[] = [];
  let currentBlocks: DocxBlock[] = [];
  let currentSize: DocxPageSize = DEFAULT_PG_SZ;
  let currentMargins: DocxPageMargins = DEFAULT_PG_MARGINS;
  let currentHeader: DocxParagraph[] | undefined;
  let currentFooter: DocxParagraph[] | undefined;
  let accumulatedH = 0;

  function contentH(): number {
    return Math.max(200, currentSize.hPx - currentMargins.topPx - currentMargins.bottomPx);
  }
  function contentW(): number {
    return Math.max(200, currentSize.wPx - currentMargins.leftPx - currentMargins.rightPx);
  }

  function flush(): void {
    pages.push({
      index: pages.length,
      size: currentSize,
      margins: currentMargins,
      elements: currentBlocks,
      header: currentHeader,
      footer: currentFooter,
    });
    currentBlocks = [];
    accumulatedH = 0;
  }

  for (const tagged_ of tagged) {
    const { block, breakBefore, breakAfter, nextSize, nextMargins, nextHeader, nextFooter } = tagged_;
    const blockH = estimateBlockHeight(block, contentW());

    // Explicit break before: flush whatever we have.
    if (breakBefore && currentBlocks.length > 0) {
      flush();
    }

    // Auto-break: block would overflow the current page.
    // Skip auto-break on the very first block of a page to avoid infinite loops.
    if (
      currentBlocks.length > 0 &&
      accumulatedH + blockH > contentH()
    ) {
      flush();
    }

    currentBlocks.push(block);
    accumulatedH += blockH;

    // Explicit break after (page break run or non-continuous section break).
    if (breakAfter) {
      flush();
      if (nextSize) currentSize = nextSize;
      if (nextMargins) currentMargins = nextMargins;
      if (nextHeader !== undefined) currentHeader = nextHeader;
      if (nextFooter !== undefined) currentFooter = nextFooter;
    } else if (nextSize) {
      // Continuous section: geometry update only, no flush.
      currentSize = nextSize;
      currentMargins = nextMargins ?? currentMargins;
    }
  }

  // Flush any remaining content.
  if (currentBlocks.length > 0 || pages.length === 0) {
    flush();
  }

  return pages;
}

// ─── Height estimation ────────────────────────────────────────────────────────

function estimateBlockHeight(block: DocxBlock, contentWidthPx: number): number {
  switch (block.kind) {
    case 'paragraph': return estimateParaHeight(block, contentWidthPx);
    case 'table':     return estimateTableHeight(block, contentWidthPx);
    case 'image':     return block.heightPx + 8;
  }
}

const PT_TO_PX = 96 / 72;

/** Resolved font size in px for a paragraph, falling back to style defaults. */
function paraFontPx(para: DocxParagraph): number {
  // Prefer per-run size, then style-chain base, then 11pt fallback.
  for (const run of para.runs) {
    if (run.sizePt && run.sizePt > 0) return run.sizePt * PT_TO_PX;
  }
  if (para.baseFontSizePt && para.baseFontSizePt > 0) return para.baseFontSizePt * PT_TO_PX;
  return 11 * PT_TO_PX;
}

function estimateParaHeight(para: DocxParagraph, contentWidthPx: number): number {
  const fontPx = paraFontPx(para);
  const lineH = fontPx * (para.lineSpacingPct ?? 1.15);

  // Estimate total text length across all runs.
  const text = para.runs.map((r) => r.text).join('');

  // Average character width ≈ 0.42× font height (proportional Latin text approximation).
  const avgCharW = fontPx * 0.42;
  const usableW = Math.max(1, contentWidthPx - (para.indentLeftPx ?? 0));
  const charsPerLine = Math.max(1, Math.floor(usableW / avgCharW));

  // At least 1 line even for empty paragraphs.
  const lines = text.length === 0 ? 1 : Math.max(1, Math.ceil(text.length / charsPerLine));

  const beforePx = (para.spaceBeforePt ?? 0) * PT_TO_PX;
  const afterPx = (para.spaceAfterPt ?? 0) * PT_TO_PX;

  return lines * lineH + beforePx + afterPx;
}

function estimateTableHeight(table: import('../model.js').DocxTable, contentWidthPx: number): number {
  let total = 0;
  for (const row of table.rows) {
    let rowH = 0;
    for (const cell of row) {
      if (!cell) continue;
      let cellH = 8; // top + bottom padding
      for (const para of cell.content) {
        cellH += estimateParaHeight(para, contentWidthPx / Math.max(1, table.colWidths.length));
      }
      rowH = Math.max(rowH, cellH);
    }
    total += rowH;
  }
  return total + 4; // table margin
}

// ─── Section helpers ──────────────────────────────────────────────────────────

function readSectPr(sectPr: XmlNode): { size: DocxPageSize; margins: DocxPageMargins } {
  const pgSz = child(sectPr, 'pgSz');
  const pgMar = child(sectPr, 'pgMar');

  const w = attrNum(pgSz, 'w:w') ?? attrNum(pgSz, 'w') ?? 12240;
  const h = attrNum(pgSz, 'w:h') ?? attrNum(pgSz, 'h') ?? 15840;

  const top    = attrNum(pgMar, 'w:top')    ?? attrNum(pgMar, 'top')    ?? 1440;
  const right  = attrNum(pgMar, 'w:right')  ?? attrNum(pgMar, 'right')  ?? 1440;
  const bottom = attrNum(pgMar, 'w:bottom') ?? attrNum(pgMar, 'bottom') ?? 1440;
  const left   = attrNum(pgMar, 'w:left')   ?? attrNum(pgMar, 'left')   ?? 1440;
  const header = attrNum(pgMar, 'w:header') ?? attrNum(pgMar, 'header') ?? 720;
  const footer = attrNum(pgMar, 'w:footer') ?? attrNum(pgMar, 'footer') ?? 720;

  return {
    size: { wPx: twipsToPx(w), hPx: twipsToPx(h) },
    margins: {
      topPx:    twipsToPx(top),
      rightPx:  twipsToPx(right),
      bottomPx: twipsToPx(bottom),
      leftPx:   twipsToPx(left),
      headerPx: twipsToPx(header),
      footerPx: twipsToPx(footer),
    },
  };
}

function readHeaderFooter(
  sectPr: XmlNode,
  ctx: BodyBuildCtx,
): { header?: DocxParagraph[]; footer?: DocxParagraph[] } {
  const headerRef = children(sectPr, 'headerReference').find(
    (el) => (attr(el, 'w:type') ?? attr(el, 'type')) === 'default',
  );
  const footerRef = children(sectPr, 'footerReference').find(
    (el) => (attr(el, 'w:type') ?? attr(el, 'type')) === 'default',
  );

  let header: DocxParagraph[] | undefined;
  let footer: DocxParagraph[] | undefined;

  if (headerRef) {
    const rId = attr(headerRef, 'r:id') ?? attr(headerRef, 'id');
    const part = rId ? ctx.pkg.resolveRel(ctx.docPart, rId)?.target : undefined;
    if (part) {
      const xml = ctx.pkg.getXml(part);
      if (xml) header = parseHeaderFooterPart(xml, ctx, part);
    }
  }

  if (footerRef) {
    const rId = attr(footerRef, 'r:id') ?? attr(footerRef, 'id');
    const part = rId ? ctx.pkg.resolveRel(ctx.docPart, rId)?.target : undefined;
    if (part) {
      const xml = ctx.pkg.getXml(part);
      if (xml) footer = parseHeaderFooterPart(xml, ctx, part);
    }
  }

  return { header, footer };
}

function parseHeaderFooterPart(xml: XmlNode, ctx: BodyBuildCtx, part: string): DocxParagraph[] {
  const paras: DocxParagraph[] = [];
  const listCounters = new Map<string, number>();
  const resolveImage = (relId: string) => ctx.pkg.resolveRel(part, relId)?.target;
  const resolveHyperlink = (relId: string) => {
    const rel = ctx.pkg.resolveRel(part, relId);
    return rel?.mode === 'External' ? rel.rawTarget : undefined;
  };
  const paraCtx = { styles: ctx.styles, numbering: ctx.numbering, listCounters, resolveImage, resolveHyperlink };
  for (const node of xml.children) {
    if (localName(node.name) === 'p') {
      const { paragraphs } = parseParagraph(node, paraCtx);
      paras.push(...paragraphs);
    }
  }
  return paras;
}
