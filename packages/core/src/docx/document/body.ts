/**
 * Body parser and paginator for WordprocessingML.
 *
 * Two-phase approach:
 *  1. Walk <w:body> and parse every element into a flat block list, tagging
 *     each block with any explicit page-break intent.
 *  2. Run a height-accumulating paginator that calls a supplied heightFn for
 *     each block.  When the viewer supplies per-row heights for a table (an
 *     array), the paginator splits the table at row boundaries so that it
 *     flows naturally across pages — exactly as Word does.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { twipsToPx } from '../units.js';
import { OpcPackage } from '../../oxml/package.js';
import { encodeNodeId } from '../edit/nodeId.js';
import type { DocxBlock, DocxTable, DocxPage, DocxPageSize, DocxPageMargins, DocxParagraph } from '../model.js';
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

// ─── Public flat-content types ────────────────────────────────────────────────

export interface DocxTaggedBlock {
  block: DocxBlock;
  breakBefore: boolean;
  breakAfter: boolean;
  nextSize?: DocxPageSize;
  nextMargins?: DocxPageMargins;
  nextHeader?: DocxBlock[];
  nextFooter?: DocxBlock[];
}

export interface DocxFlatContent {
  taggedBlocks: DocxTaggedBlock[];
  initSize: DocxPageSize;
  initMargins: DocxPageMargins;
  initHeader?: DocxBlock[];
  initFooter?: DocxBlock[];
}

/**
 * Height function signature.
 * - Returns a single number for non-table blocks.
 * - Returns number[] (one entry per model row) for table blocks when the caller
 *   has per-row measurements.  The paginator uses this array to split the table
 *   at page boundaries.  Return a plain number to treat the table as a single
 *   indivisible block (heuristic fallback behaviour).
 */
export type BlockHeightFn = (block: DocxBlock, contentWidthPx: number) => number | number[];

// ─── Entry points ─────────────────────────────────────────────────────────────

export function collectDocxContent(
  bodyEl: XmlNode,
  ctx: BodyBuildCtx,
  /** Child-index path to `<w:body>` from the part root (editable mode). */
  bodyPath: number[] = [],
): DocxFlatContent {
  const bodySectPrEl = bodyEl.children.slice().reverse().find(
    (n) => localName(n.name) === 'sectPr',
  );
  const { size: initSize, margins: initMargins } = bodySectPrEl
    ? readSectPr(bodySectPrEl)
    : { size: DEFAULT_PG_SZ, margins: DEFAULT_PG_MARGINS };
  const { header: initHeader, footer: initFooter } = bodySectPrEl
    ? readHeaderFooter(bodySectPrEl, ctx)
    : {};

  const taggedBlocks = collectBlocks(bodyEl, ctx, bodyPath);
  return { taggedBlocks, initSize, initMargins, initHeader, initFooter };
}

export function paginateDocxContent(
  content: DocxFlatContent,
  heightFn: BlockHeightFn = estimateBlockHeight,
): DocxPage[] {
  return paginate(
    content.taggedBlocks,
    content.initSize,
    content.initMargins,
    content.initHeader,
    content.initFooter,
    heightFn,
  );
}

export function buildPages(bodyEl: XmlNode, ctx: BodyBuildCtx): DocxPage[] {
  return paginateDocxContent(collectDocxContent(bodyEl, ctx));
}

// ─── Phase 1: collect all blocks with break metadata ─────────────────────────

function collectBlocks(bodyEl: XmlNode, ctx: BodyBuildCtx, bodyPath: number[]): DocxTaggedBlock[] {
  const out: DocxTaggedBlock[] = [];
  const listCounters = new Map<string, number>();

  function resolveImage(relId: string): string | undefined {
    return ctx.pkg.resolveRel(ctx.docPart, relId)?.target;
  }
  function resolveHyperlink(relId: string): string | undefined {
    const rel = ctx.pkg.resolveRel(ctx.docPart, relId);
    return rel?.mode === 'External' ? rel.rawTarget : undefined;
  }

  const paraCtx = { styles: ctx.styles, numbering: ctx.numbering, listCounters, resolveImage, resolveHyperlink, partPath: ctx.docPart };

  for (let i = 0; i < bodyEl.children.length; i++) {
    const node = bodyEl.children[i]!;
    const name = localName(node.name);

    if (name === 'p') {
      const pPr = child(node, 'pPr');
      const sectPr = child(pPr, 'sectPr');

      const { paragraphs, endsWithPageBreak } = parseParagraph(node, paraCtx, [...bodyPath, i]);

      const inlineImages: import('../model.js').DocxInlineImage[] = [];
      node.children.forEach((childNode, rIdx) => {
        if (localName(childNode.name) === 'r') {
          for (const gc of childNode.children) {
            if (localName(gc.name) === 'drawing') {
              const img = parseDrawing(gc, resolveImage);
              if (img) {
                img.nodeId = encodeNodeId(ctx.docPart, [...bodyPath, i, rIdx]);
                inlineImages.push(img);
              }
            }
          }
        }
      });

      const hasText = paragraphs.some(p => p.runs.some(r => r.text.trim()));
      const skipPara = inlineImages.length > 0 && !hasText;

      if (!skipPara) {
        for (let i = 0; i < paragraphs.length; i++) {
          const para = paragraphs[i]!;
          const isLast = i === paragraphs.length - 1;
          out.push({
            block: para,
            breakBefore: i === 0 && !!para.pageBreakBefore,
            breakAfter: isLast && endsWithPageBreak,
          });
        }
      }

      for (let ii = 0; ii < inlineImages.length; ii++) {
        out.push({
          block: inlineImages[ii]!,
          breakBefore: false,
          breakAfter: skipPara && ii === inlineImages.length - 1 && endsWithPageBreak,
        });
      }

      if (sectPr && out.length > 0) {
        const { size, margins } = readSectPr(sectPr);
        const sectType = attr(child(sectPr, 'type'), 'w:val') ?? attr(child(sectPr, 'type'), 'val') ?? 'nextPage';
        const { header, footer } = readHeaderFooter(sectPr, ctx);
        const last = out[out.length - 1]!;
        if (sectType !== 'continuous') last.breakAfter = true;
        last.nextSize = size;
        last.nextMargins = margins;
        last.nextHeader = header;
        last.nextFooter = footer;
      }

    } else if (name === 'tbl') {
      out.push({ block: parseTable(node, paraCtx, [...bodyPath, i]), breakBefore: false, breakAfter: false });

    } else if (name === 'sectPr') {
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

// ─── Phase 2: paginator with row-level table splitting ────────────────────────

function paginate(
  tagged: DocxTaggedBlock[],
  initSize: DocxPageSize,
  initMargins: DocxPageMargins,
  initHeader: DocxBlock[] | undefined,
  initFooter: DocxBlock[] | undefined,
  heightFn: BlockHeightFn,
): DocxPage[] {
  const pages: DocxPage[] = [];
  let currentBlocks: DocxBlock[] = [];
  let currentSize = initSize;
  let currentMargins = initMargins;
  let currentHeader = initHeader;
  let currentFooter = initFooter;
  let accumulatedH = 0;

  const contentH = () => Math.max(200, currentSize.hPx - currentMargins.topPx - currentMargins.bottomPx);
  const contentW = () => Math.max(200, currentSize.wPx - currentMargins.leftPx - currentMargins.rightPx);

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

  function applySectionChange(
    nextSize?: DocxPageSize,
    nextMargins?: DocxPageMargins,
    nextHeader?: DocxBlock[],
    nextFooter?: DocxBlock[],
  ): void {
    if (nextSize) currentSize = nextSize;
    if (nextMargins) currentMargins = nextMargins;
    if (nextHeader !== undefined) currentHeader = nextHeader;
    if (nextFooter !== undefined) currentFooter = nextFooter;
  }

  for (const tagged_ of tagged) {
    const { block, breakBefore, breakAfter, nextSize, nextMargins, nextHeader, nextFooter } = tagged_;

    if (breakBefore && currentBlocks.length > 0) flush();

    if (block.kind === 'table') {
      // ── Table: split row-by-row when per-row heights are available ──────
      const rawH = heightFn(block, contentW());

      if (!Array.isArray(rawH)) {
        // Heuristic single-height: treat as indivisible (legacy behaviour).
        if (currentBlocks.length > 0 && accumulatedH + rawH > contentH()) flush();
        currentBlocks.push(block);
        accumulatedH += rawH;
      } else {
        // DOM-measured per-row heights: distribute rows across pages.
        const rowHs = rawH;
        let rowStart = 0;

        while (rowStart < block.rows.length) {
          const avail = contentH() - accumulatedH;

          // Count rows that fit in the available space.
          let rowsFit = 0;
          let heightFit = 0;
          for (let ri = rowStart; ri < rowHs.length; ri++) {
            const rh = rowHs[ri] ?? 0;
            if (heightFit + rh > avail && rowsFit > 0) break;
            rowsFit++;
            heightFit += rh;
          }

          if (rowsFit === 0) {
            // Nothing fits in the remaining space.
            if (currentBlocks.length > 0) {
              // Flush current page and retry with a fresh page.
              flush();
              continue;
            }
            // Already on a fresh page — force at least one row to prevent
            // an infinite loop (happens when a single row is taller than the page).
            rowsFit = 1;
            heightFit = rowHs[rowStart] ?? 0;
          }

          const rowEnd = Math.min(rowStart + rowsFit, block.rows.length);
          const subTable: DocxTable = {
            kind: 'table',
            widthPx: block.widthPx,
            colWidths: block.colWidths,
            rows: block.rows.slice(rowStart, rowEnd),
          };

          currentBlocks.push(subTable);
          accumulatedH += heightFit;
          rowStart = rowEnd;

          // If more rows remain, flush to start a fresh page.
          if (rowStart < block.rows.length) flush();
        }
      }

    } else {
      // ── Non-table block ────────────────────────────────────────────────
      const blockH = heightFn(block, contentW()) as number;
      if (currentBlocks.length > 0 && accumulatedH + blockH > contentH()) flush();
      currentBlocks.push(block);
      accumulatedH += blockH;
    }

    if (breakAfter) {
      flush();
      applySectionChange(nextSize, nextMargins, nextHeader, nextFooter);
    } else if (nextSize) {
      applySectionChange(nextSize, nextMargins, nextHeader, nextFooter);
    }
  }

  if (currentBlocks.length > 0 || pages.length === 0) flush();

  return pages;
}

// ─── Heuristic height estimator (fallback when no DOM available) ──────────────

export function estimateBlockHeight(block: DocxBlock, contentWidthPx: number): number {
  switch (block.kind) {
    case 'paragraph': return estimateParaHeight(block, contentWidthPx);
    case 'table':     return estimateTableHeight(block, contentWidthPx);
    case 'image':     return block.heightPx + 8;
  }
}

const PT_TO_PX = 96 / 72;

function paraFontPx(para: DocxParagraph): number {
  for (const run of para.runs) {
    if (run.sizePt && run.sizePt > 0) return run.sizePt * PT_TO_PX;
  }
  if (para.baseFontSizePt && para.baseFontSizePt > 0) return para.baseFontSizePt * PT_TO_PX;
  return 11 * PT_TO_PX;
}

function estimateParaHeight(para: DocxParagraph, contentWidthPx: number): number {
  const fontPx = paraFontPx(para);
  const lineH = fontPx * (para.lineSpacingPct ?? 1.15);
  const text = para.runs.map((r) => r.text).join('');
  const avgCharW = fontPx * 0.5;
  const usableW = Math.max(1, contentWidthPx - (para.indentLeftPx ?? 0));
  const charsPerLine = Math.max(1, Math.floor(usableW / avgCharW));
  const lines = text.length === 0 ? 1 : Math.max(1, Math.ceil(text.length / charsPerLine));
  const beforePx = (para.spaceBeforePt ?? 0) * PT_TO_PX;
  const afterPx  = (para.spaceAfterPt  ?? 0) * PT_TO_PX;
  return lines * lineH + beforePx + afterPx;
}

function estimateTableHeight(table: DocxTable, contentWidthPx: number): number {
  const totalColW = table.colWidths.reduce((s, w) => s + w, 0) || contentWidthPx;
  let total = 0;
  for (let r = 0; r < table.rows.length; r++) {
    const row = table.rows[r]!;
    let rowH = 0;
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (!cell) continue;
      const colW = table.colWidths[c] ?? (totalColW / Math.max(1, table.colWidths.length));
      let cellH = 4;
      for (const para of cell.content) {
        cellH += estimateParaHeight(para, colW - 14);
      }
      rowH = Math.max(rowH, cellH);
    }
    total += rowH;
  }
  return total;
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
): { header?: DocxBlock[]; footer?: DocxBlock[] } {
  const headerRef = children(sectPr, 'headerReference').find(
    (el) => (attr(el, 'w:type') ?? attr(el, 'type')) === 'default',
  );
  const footerRef = children(sectPr, 'footerReference').find(
    (el) => (attr(el, 'w:type') ?? attr(el, 'type')) === 'default',
  );

  let header: DocxBlock[] | undefined;
  let footer: DocxBlock[] | undefined;

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

function parseHeaderFooterPart(xml: XmlNode, ctx: BodyBuildCtx, part: string): DocxBlock[] {
  const blocks: DocxBlock[] = [];
  const listCounters = new Map<string, number>();
  const resolveImage = (relId: string) => ctx.pkg.resolveRel(part, relId)?.target;
  const resolveHyperlink = (relId: string) => {
    const rel = ctx.pkg.resolveRel(part, relId);
    return rel?.mode === 'External' ? rel.rawTarget : undefined;
  };
  const paraCtx = { styles: ctx.styles, numbering: ctx.numbering, listCounters, resolveImage, resolveHyperlink, partPath: part };
  for (let i = 0; i < xml.children.length; i++) {
    const node = xml.children[i]!;
    if (localName(node.name) === 'p') {
      const { paragraphs } = parseParagraph(node, paraCtx, [i]);

      const imgs: import('../model.js').DocxInlineImage[] = [];
      node.children.forEach((childNode, rIdx) => {
        if (localName(childNode.name) === 'r') {
          for (const gc of childNode.children) {
            if (localName(gc.name) === 'drawing') {
              const img = parseDrawing(gc, resolveImage);
              if (img) {
                img.nodeId = encodeNodeId(part, [i, rIdx]);
                imgs.push(img);
              }
            }
          }
        }
      });

      const hasText = paragraphs.some(p => p.runs.some(r => r.text.trim()));
      if (imgs.length === 0 || hasText) blocks.push(...paragraphs);
      blocks.push(...imgs);
    }
  }
  return blocks;
}
