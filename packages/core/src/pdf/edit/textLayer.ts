/**
 * PDF text layer extraction.
 *
 * Uses pdf.js getTextContent() to extract positioned text items from a page,
 * then groups only horizontally-adjacent items (within the same logical text run)
 * into editable PdfTextBlock units.
 *
 * Coordinate systems:
 *  - PDF user space:  origin bottom-left, y up. Matches pdf-lib coordinates.
 *  - CSS space:       origin top-left, y down. Used for DOM overlay positioning.
 * At scale 1.0 the magnitudes are identical (1 PDF point = 1 CSS pixel).
 */
import type { PDFDocumentProxy } from 'pdfjs-dist';
import type { TextItem } from 'pdfjs-dist/types/src/display/api.js';
import type { PdfTextItem, PdfTextBlock } from '../model.js';

function isTextItem(item: unknown): item is TextItem {
  return typeof item === 'object' && item !== null && 'str' in item;
}

/**
 * Extract all text items from a single PDF page and convert them into
 * CSS-positioned PdfTextItem objects.
 */
export async function extractPageTextItems(
  proxy: PDFDocumentProxy,
  pageIndex: number,
  pageHeightPx: number,
): Promise<PdfTextItem[]> {
  const page = await proxy.getPage(pageIndex + 1);
  const textContent = await page.getTextContent();
  page.cleanup();

  const items: PdfTextItem[] = [];
  let itemIdx = 0;

  for (const raw of textContent.items) {
    if (!isTextItem(raw) || raw.str.trim() === '') continue;

    const transform = raw.transform as number[];
    const a = transform[0] ?? 0;
    const b = transform[1] ?? 0;
    const d = transform[3] ?? 0;
    const e = transform[4] ?? 0;
    const f = transform[5] ?? 0;

    // Approximate font size from the scale component of the transform matrix.
    const fontSize = Math.max(Math.sqrt(a * a + b * b), Math.abs(d), 1);
    const pdfWidth  = Math.max(raw.width, 0);
    // pdfHeight approximates the full line height (ascender + descender).
    const pdfHeight = Math.max(raw.height, fontSize);

    // CSS y: flip PDF y-axis (baseline from bottom → top from top).
    // We position the div top at (pageHeight - baseline - ascender).
    // Approximate ascender as 80% of height, descender as 20%.
    const cssX = e;
    const cssY = pageHeightPx - f - pdfHeight;

    items.push({
      id: `${pageIndex}-${itemIdx}`,
      pageIndex,
      str: raw.str,
      cssX,
      cssY,
      cssWidth: Math.max(pdfWidth, 4),
      cssHeight: Math.max(pdfHeight, fontSize),
      pdfX: e,
      pdfY: f,            // baseline in PDF coords (for pdf-lib drawText)
      pdfWidth: Math.max(pdfWidth, 4),
      pdfHeight: Math.max(pdfHeight, fontSize),
      fontSize,
      fontName: raw.fontName,
      transform,
    });

    itemIdx++;
  }

  return items;
}

/**
 * Group text items into editable blocks.
 *
 * Rules:
 *  1. Items with baselines within (avgFontSize × 0.5) are considered the same line.
 *  2. Within a line, items are merged into a block only when the gap between them
 *     is ≤ (avgFontSize × 2). Larger gaps mean different columns / separate fields.
 *  3. Within each block, items are sorted left-to-right and joined with a space
 *     only when there is a visible gap.
 *
 * This prevents cross-column merging in tables, which was causing the
 * single-block-per-row mess visible in the screenshots.
 */
export function groupItemsIntoBlocks(items: PdfTextItem[]): PdfTextBlock[] {
  if (items.length === 0) return [];

  // Sort top-to-bottom (descending PDF y), then left-to-right.
  const sorted = [...items].sort((a, b) => {
    const dy = b.pdfY - a.pdfY;
    if (Math.abs(dy) > 1) return dy;
    return a.pdfX - b.pdfX;
  });

  // Phase 1: bucket items into visual lines.
  const lines: PdfTextItem[][] = [];
  let currentLine: PdfTextItem[] = [sorted[0]!];
  let lineBaselineY = sorted[0]!.pdfY;
  let lineFontSize  = sorted[0]!.fontSize;

  for (let i = 1; i < sorted.length; i++) {
    const item = sorted[i]!;
    const threshold = (lineFontSize + item.fontSize) * 0.4;
    if (Math.abs(item.pdfY - lineBaselineY) <= threshold) {
      currentLine.push(item);
    } else {
      lines.push(currentLine);
      currentLine   = [item];
      lineBaselineY = item.pdfY;
      lineFontSize  = item.fontSize;
    }
  }
  lines.push(currentLine);

  // Phase 2: within each line, split into column-aware blocks by horizontal gap.
  const blocks: PdfTextBlock[] = [];
  let blockIdx = 0;

  for (const line of lines) {
    // Sort left-to-right within the line.
    line.sort((a, b) => a.pdfX - b.pdfX);

    const avgFontSize = line.reduce((s, it) => s + it.fontSize, 0) / line.length;
    // Maximum gap before we consider items to be in separate columns/fields.
    const columnGapThreshold = avgFontSize * 2.5;

    let group: PdfTextItem[] = [line[0]!];

    const flushGroup = () => {
      if (group.length === 0) return;
      blocks.push(buildBlock(group, blockIdx++));
      group = [];
    };

    for (let i = 1; i < line.length; i++) {
      const prev = line[i - 1]!;
      const item = line[i]!;
      const gap  = item.pdfX - (prev.pdfX + prev.pdfWidth);

      if (gap > columnGapThreshold) {
        flushGroup();
        group = [item];
      } else {
        group.push(item);
      }
    }
    flushGroup();
  }

  return blocks;
}

function buildBlock(items: PdfTextItem[], idx: number): PdfTextBlock {
  const pageIndex = items[0]!.pageIndex;

  // Concatenate text, inserting a space where there is a visible gap.
  let text = items[0]!.str;
  for (let i = 1; i < items.length; i++) {
    const prev = items[i - 1]!;
    const curr = items[i]!;
    const gap  = curr.pdfX - (prev.pdfX + prev.pdfWidth);
    if (gap > prev.fontSize * 0.3) text += ' ';
    text += curr.str;
  }

  const cssXMin  = Math.min(...items.map(it => it.cssX));
  const cssYMin  = Math.min(...items.map(it => it.cssY));
  const cssXMax  = Math.max(...items.map(it => it.cssX + it.cssWidth));
  const cssYMax  = Math.max(...items.map(it => it.cssY + it.cssHeight));

  const pdfXMin  = Math.min(...items.map(it => it.pdfX));
  // pdfYBaseline: the actual text baseline (used for pdf-lib drawText).
  // Take the mode/max since all items on the same line share nearly the same baseline.
  const pdfYBaseline = Math.max(...items.map(it => it.pdfY));
  const pdfXMax  = Math.max(...items.map(it => it.pdfX + it.pdfWidth));
  // pdfYTop: top of the text = baseline + largest ascent in the group.
  const pdfYTop  = Math.max(...items.map(it => it.pdfY + it.pdfHeight));
  const fontSize = Math.max(...items.map(it => it.fontSize));

  return {
    id: `block-${pageIndex}-${idx}`,
    pageIndex,
    items,
    text,
    cssX:    cssXMin,
    cssY:    cssYMin,
    cssWidth:  Math.max(cssXMax - cssXMin, 6),
    cssHeight: Math.max(cssYMax - cssYMin, fontSize),
    pdfX:    pdfXMin,
    pdfY:    pdfYBaseline,   // baseline — matches pdf-lib drawText y param
    pdfWidth:  Math.max(pdfXMax - pdfXMin, 6),
    pdfHeight: Math.max(pdfYTop - pdfYBaseline, fontSize),
    fontSize,
  };
}
