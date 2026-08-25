/**
 * Reading formatting back out of a cell, for the toolbar.
 *
 * The write side lives in styleWrite.ts; this is the other half — what the
 * toolbar should show for whatever cell the user is standing on.
 */
import type { RunFormat } from '../../oxml/edit/format.js';
import type { XlsxAlignment, XlsxCell } from '../model.js';
import type { CellFormatPatch } from './styleWrite.js';

function hexOf(color: string | undefined): string | undefined {
  if (!color) return undefined;
  const clean = color.replace(/^#/, '').toUpperCase();
  return clean.length === 8 ? clean.slice(2) : clean;
}

/** Character formatting in effect on a cell. */
export function readCellFormat(cell: XlsxCell | undefined): RunFormat {
  const font = cell?.style?.font;
  if (!font) return {};
  const out: RunFormat = {};
  if (font.bold !== undefined) out.bold = font.bold;
  if (font.italic !== undefined) out.italic = font.italic;
  if (font.underline !== undefined) out.underline = font.underline;
  if (font.strike !== undefined) out.strike = font.strike;
  if (font.sizePt !== undefined) out.sizePt = font.sizePt;
  if (font.name !== undefined) out.font = font.name;
  const color = hexOf(font.colorHex);
  if (color) out.colorHex = color;
  return out;
}

/** Everything the toolbar reflects for a cell: font, fill and alignment. */
export function readCellPatch(cell: XlsxCell | undefined): CellFormatPatch {
  const out: CellFormatPatch = readCellFormat(cell);
  const fill = cell?.style?.fill;
  if (fill?.patternType === 'solid') {
    const hex = hexOf(fill.fgColorHex);
    if (hex) out.fillHex = hex;
  }
  const align = cell?.style?.alignment;
  if (align && (align.horizontal || align.vertical || align.wrapText)) {
    out.alignment = { ...align } as XlsxAlignment;
  }
  if (cell?.style?.numFmtCode) out.numFmtCode = cell.style.numFmtCode;
  return out;
}
