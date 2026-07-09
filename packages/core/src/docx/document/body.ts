/**
 * Body walk: <w:body> -> DocxPage[].
 *
 * Splits the body into sections at each <w:sectPr> (a section-ending paragraph
 * mark, or the final body-level sectPr) and emits one continuous-flow page per
 * section with that section's page size and margins. Measured pagination
 * (breaking a section across fixed-height pages) is a later fidelity pass.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { twipToPx } from '../units.js';
import { parseParagraph } from '../paragraphs/paragraph.js';
import { parseTable } from '../tables/table.js';
import type { DocxBlock, DocxPage, DocxPageSize, DocxPageMargins } from '../model.js';
import type { ParseContext } from './context.js';

/** US Letter, the Word default when a section omits <w:pgSz>. */
const DEFAULT_PAGE: DocxPageSize = { wPx: twipToPx(12240), hPx: twipToPx(15840) };
const DEFAULT_MARGINS: DocxPageMargins = {
  topPx: twipToPx(1440),
  rightPx: twipToPx(1440),
  bottomPx: twipToPx(1440),
  leftPx: twipToPx(1440),
  headerPx: twipToPx(720),
  footerPx: twipToPx(720),
};

export function parseBody(body: XmlNode, ctx: ParseContext): DocxPage[] {
  const pages: DocxPage[] = [];
  let current: DocxBlock[] = [];

  const flush = (sectPr: XmlNode | undefined) => {
    const { size, margins } = readSectPr(sectPr);
    pages.push({ index: pages.length, size, margins, elements: current });
    current = [];
  };

  for (const node of body.children) {
    const name = localName(node.name);
    if (name === 'p') {
      current.push(...parseParagraph(node, ctx));
      const sectPr = child(child(node, 'pPr'), 'sectPr');
      if (sectPr) flush(sectPr);
    } else if (name === 'tbl') {
      current.push(parseTable(node, ctx));
    } else if (name === 'sectPr') {
      // Final section's properties (direct body child).
      flush(node);
    }
  }

  // Body ended without a trailing sectPr child (unusual) — emit what's left.
  if (current.length > 0 || pages.length === 0) flush(undefined);

  return pages;
}

function readSectPr(sectPr: XmlNode | undefined): { size: DocxPageSize; margins: DocxPageMargins } {
  if (!sectPr) return { size: { ...DEFAULT_PAGE }, margins: { ...DEFAULT_MARGINS } };

  const pgSz = child(sectPr, 'pgSz');
  let w = attrNum(pgSz, 'w') ?? 12240;
  let h = attrNum(pgSz, 'h') ?? 15840;
  if (attr(pgSz, 'orient') === 'landscape' && w < h) [w, h] = [h, w];

  const pgMar = child(sectPr, 'pgMar');
  const mar = (name: string, fallback: number) => attrNum(pgMar, name) ?? fallback;

  return {
    size: { wPx: twipToPx(w), hPx: twipToPx(h) },
    margins: {
      topPx: twipToPx(mar('top', 1440)),
      rightPx: twipToPx(mar('right', 1440)),
      bottomPx: twipToPx(mar('bottom', 1440)),
      leftPx: twipToPx(mar('left', 1440)),
      headerPx: twipToPx(mar('header', 720)),
      footerPx: twipToPx(mar('footer', 720)),
    },
  };
}
