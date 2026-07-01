/**
 * PDF export with applied edits, block style overrides, and annotations using pdf-lib.
 *
 * Strategy for text block edits (replaceText) and style overrides (styleBlock):
 *  1. Draw a filled white rectangle over the original text at the original position.
 *  2. Render the replacement text at the (possibly new) position with the overridden style.
 *
 * Strategy for annotations (addAnnotation):
 *  1. Resolve the font via the display name → PDF Standard14 family mapping.
 *  2. Calculate per-line x offset for center/right alignment.
 *  3. Draw each line — no background rectangle.
 */
import { PDFDocument, PDFFont, StandardFonts, rgb, type RGB } from 'pdf-lib';
import type { PdfTextBlock, PdfAnnotation, PdfBlockStyle } from '../model.js';
import { resolvePdfFamily } from './fonts.js';

// ── Helpers ────────────────────────────────────────────────────────────────

function parseCssColor(color: string): RGB {
  const rgbMatch = color.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
  if (rgbMatch) {
    return rgb(
      parseInt(rgbMatch[1]!, 10) / 255,
      parseInt(rgbMatch[2]!, 10) / 255,
      parseInt(rgbMatch[3]!, 10) / 255,
    );
  }
  const hexMatch = color.match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
  if (hexMatch) {
    return rgb(
      parseInt(hexMatch[1]!, 16) / 255,
      parseInt(hexMatch[2]!, 16) / 255,
      parseInt(hexMatch[3]!, 16) / 255,
    );
  }
  return rgb(0, 0, 0);
}

function resolveStandardFont(fontName: string, bold: boolean, italic: boolean): StandardFonts {
  const family = resolvePdfFamily(fontName);
  if (family === 'times') {
    if (bold && italic) return StandardFonts.TimesRomanBoldItalic;
    if (bold)           return StandardFonts.TimesRomanBold;
    if (italic)         return StandardFonts.TimesRomanItalic;
    return StandardFonts.TimesRoman;
  }
  if (family === 'courier') {
    if (bold && italic) return StandardFonts.CourierBoldOblique;
    if (bold)           return StandardFonts.CourierBold;
    if (italic)         return StandardFonts.CourierOblique;
    return StandardFonts.Courier;
  }
  // helvetica (default)
  if (bold && italic) return StandardFonts.HelveticaBoldOblique;
  if (bold)           return StandardFonts.HelveticaBold;
  if (italic)         return StandardFonts.HelveticaOblique;
  return StandardFonts.Helvetica;
}

// ── Export ─────────────────────────────────────────────────────────────────

/**
 * Apply text edits, block style overrides, and annotations to a PDF and return the new bytes.
 *
 * @param originalBytes  Raw bytes of the unedited PDF.
 * @param blocks         All PdfTextBlock[][] indexed by page.
 * @param edits          Map of blockId → newText (only changed blocks).
 * @param annotations    Map of annotationId → PdfAnnotation (user-added text boxes).
 * @param blockStyles    Map of blockId → PdfBlockStyle (position/style overrides).
 */
export async function exportPdfWithEdits(
  originalBytes: Uint8Array,
  blocks:        PdfTextBlock[][],
  edits:         Map<string, string>,
  annotations:   Map<string, PdfAnnotation>,
  blockStyles:   Map<string, PdfBlockStyle> = new Map(),
): Promise<Uint8Array> {
  if (edits.size === 0 && annotations.size === 0 && blockStyles.size === 0) return originalBytes;

  const pdfDoc   = await PDFDocument.load(originalBytes, { ignoreEncryption: true });
  const pdfPages = pdfDoc.getPages();

  // Lazy font cache: embed each variant only once.
  const fontCache = new Map<string, Promise<PDFFont>>();
  const getFont   = (fontName: string, bold: boolean, italic: boolean): Promise<PDFFont> => {
    const key = `${fontName}-${bold}-${italic}`;
    if (!fontCache.has(key)) {
      fontCache.set(key, pdfDoc.embedFont(resolveStandardFont(fontName, bold, italic)));
    }
    return fontCache.get(key)!;
  };

  // Build a flat block lookup map.
  const blockLookup = new Map<string, PdfTextBlock>();
  for (const pageBlocks of blocks) {
    for (const block of pageBlocks) blockLookup.set(block.id, block);
  }

  // ── 1. Apply text block edits and style overrides ──────────────────────────

  // Collect all blocks that need any processing.
  const blocksToProcess = new Set<string>();
  for (const id of edits.keys())       blocksToProcess.add(id);
  for (const id of blockStyles.keys()) blocksToProcess.add(id);

  for (const blockId of blocksToProcess) {
    const block      = blockLookup.get(blockId);
    const blockStyle = blockStyles.get(blockId);

    if (!block) {
      console.warn(`[pdf-export] Block not found for id="${blockId}" — skipping.`);
      continue;
    }

    const page = pdfPages[block.pageIndex];
    if (!page) {
      console.warn(`[pdf-export] Page ${block.pageIndex} not found — skipping block "${blockId}".`);
      continue;
    }

    // ── Erase original position ──────────────────────────────────────────────
    const { pdfX, pdfY, pdfWidth, pdfHeight, fontSize } = block;
    const eraseFontSize = Math.max(fontSize, 6);
    const descender     = eraseFontSize * 0.22;
    const erasePad      = 2;

    page.drawRectangle({
      x: pdfX - erasePad,           y: pdfY - descender - erasePad,
      width: pdfWidth + erasePad * 2, height: pdfHeight + descender + erasePad * 2,
      color: rgb(1, 1, 1),           opacity: 1,
    });

    // ── Determine draw text ──────────────────────────────────────────────────
    const displayText = edits.get(blockId) ?? block.text;
    if (!displayText.trim()) continue;

    // ── Resolve draw position (override or original) ─────────────────────────
    // CSS Y increases downward; PDF Y increases upward. Shift = cssY delta.
    const drawX = blockStyle?.cssX !== undefined
      ? block.pdfX + (blockStyle.cssX - block.cssX)
      : pdfX;
    const drawY = blockStyle?.cssY !== undefined
      ? block.pdfY - (blockStyle.cssY - block.cssY)
      : pdfY;
    const drawWidth = blockStyle?.cssWidth ?? pdfWidth;

    // ── Resolve font and size ────────────────────────────────────────────────
    const fontName   = blockStyle?.fontFamily ?? 'helvetica';
    const bold       = blockStyle?.fontWeight === 'bold';
    const italic     = blockStyle?.fontStyle  === 'italic';
    const drawFontSize = Math.max(blockStyle?.fontSize ?? fontSize, 6);
    const font       = await getFont(fontName, bold, italic);

    const lineHeight = drawFontSize * 1.2;
    const lines      = displayText.split('\n');
    const textColor  = blockStyle?.color ? parseCssColor(blockStyle.color) : rgb(0, 0, 0);
    const textAlign  = blockStyle?.textAlign ?? 'left';

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (!line?.trim()) continue;

      let lineX = drawX;
      if (textAlign !== 'left') {
        const lineW = font.widthOfTextAtSize(line, drawFontSize);
        lineX = textAlign === 'center'
          ? drawX + (drawWidth - lineW) / 2
          : drawX + drawWidth - lineW;
      }

      page.drawText(line, {
        x: lineX, y: drawY - i * lineHeight,
        size: drawFontSize, font, color: textColor,
        maxWidth: Math.max(drawWidth * 2, 300),
      });
    }
  }

  // ── 2. Draw annotations ───────────────────────────────────────────────────

  for (const ann of annotations.values()) {
    const page = pdfPages[ann.pageIndex];
    if (!page) {
      console.warn(`[pdf-export] Page ${ann.pageIndex} not found — skipping annotation "${ann.id}".`);
      continue;
    }

    const fontName  = ann.fontFamily || 'helvetica';
    const bold      = ann.fontWeight === 'bold';
    const italic    = ann.fontStyle  === 'italic';
    const font      = await getFont(fontName, bold, italic);
    const pageH     = page.getSize().height;
    const fs        = Math.max(ann.fontSize, 6);
    const lineH     = fs * 1.4;
    const lines     = ann.text.split('\n');
    const textColor = parseCssColor(ann.color);
    const align     = ann.textAlign ?? 'left';

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (!line) continue;

      // Alignment: compute per-line x offset.
      let lineX = ann.cssX;
      if (align !== 'left') {
        const lineW = font.widthOfTextAtSize(line, fs);
        lineX = align === 'center'
          ? ann.cssX + (ann.width - lineW) / 2
          : ann.cssX + ann.width - lineW; // right
      }

      // Baseline y in PDF coords (origin bottom-left, y up).
      const baselineY = pageH - ann.cssY - fs - i * lineH;

      page.drawText(line, { x: lineX, y: baselineY, size: fs, font, color: textColor });
    }
  }

  return pdfDoc.save();
}
