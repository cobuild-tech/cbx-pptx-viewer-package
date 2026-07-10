/**
 * Body walk: <w:body> -> DocxSection[].
 *
 * Splits the body into sections at each <w:sectPr> (a section-ending paragraph
 * mark, or the final body-level sectPr). Each section carries its page size,
 * margins, block flow, and resolved header/footer content. The viewer's
 * paginator then flows each section's blocks into fixed-size pages.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { twipToPx } from '../units.js';
import { logicalChildren } from '../content.js';
import { parseParagraph } from '../paragraphs/paragraph.js';
import { parseTable, fitTableWidth } from '../tables/table.js';
import { fitImageWidth } from '../images/image.js';
import type { DocxBlock, DocxSection, DocxPageSize, DocxPageMargins } from '../model.js';
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

export function parseBody(body: XmlNode, ctx: ParseContext): DocxSection[] {
  const sections: DocxSection[] = [];
  let current: DocxBlock[] = [];

  const flush = (sectPr: XmlNode | undefined) => {
    const { size, margins, titlePg } = readSectPr(sectPr);
    const hf = readHeaderFooter(sectPr, ctx);
    // Autofit top-level tables to the section's content width so oversized/pct
    // tables don't overflow the right margin (Word scales them to fit).
    const contentW = size.wPx - margins.leftPx - margins.rightPx;
    for (const block of current) {
      if (block.kind === 'table') fitTableWidth(block, contentW);
      else if (block.kind === 'image') fitImageWidth(block, contentW);
    }
    sections.push({ index: sections.length, size, margins, blocks: current, ...(titlePg ? { titlePg } : {}), ...hf });
    current = [];
  };

  for (const node of logicalChildren(body)) {
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
  if (current.length > 0 || sections.length === 0) flush(undefined);

  return sections;
}

/** Parse the block children (paragraphs, tables) of a body/header/footer node. */
export function parseBlocks(container: XmlNode | undefined, ctx: ParseContext): DocxBlock[] {
  if (!container) return [];
  const out: DocxBlock[] = [];
  for (const node of logicalChildren(container)) {
    const name = localName(node.name);
    if (name === 'p') out.push(...parseParagraph(node, ctx));
    else if (name === 'tbl') out.push(parseTable(node, ctx));
  }
  return out;
}

function readHeaderFooter(
  sectPr: XmlNode | undefined,
  ctx: ParseContext,
): { header?: DocxBlock[]; footer?: DocxBlock[]; firstHeader?: DocxBlock[]; firstFooter?: DocxBlock[] } {
  if (!sectPr) return {};
  const out: { header?: DocxBlock[]; footer?: DocxBlock[]; firstHeader?: DocxBlock[]; firstFooter?: DocxBlock[] } = {};

  const hdrs = children(sectPr, 'headerReference');
  const ftrs = children(sectPr, 'footerReference');

  const header = resolveRef(hdrs, 'default', 'hdr', ctx);
  if (header?.length) out.header = header;
  const firstHeader = resolveRef(hdrs, 'first', 'hdr', ctx);
  if (firstHeader?.length) out.firstHeader = firstHeader;

  const footer = resolveRef(ftrs, 'default', 'ftr', ctx);
  if (footer?.length) out.footer = footer;
  const firstFooter = resolveRef(ftrs, 'first', 'ftr', ctx);
  if (firstFooter?.length) out.firstFooter = firstFooter;

  return out;
}

/**
 * Resolve the header/footer reference of a given type ('default' | 'first' |
 * 'even') and parse its block flow. Types are kept distinct — a first-page
 * header must not fall back to the default, or <w:titlePg> would render the
 * wrong banner on page 1.
 */
function resolveRef(
  refs: XmlNode[],
  type: 'default' | 'first' | 'even',
  root: string,
  ctx: ParseContext,
): DocxBlock[] | undefined {
  const pick = refs.find((r) => (attr(r, 'type') ?? 'default') === type);
  if (!pick) return undefined;
  const rel = ctx.rel(attr(pick, 'id'));
  if (!rel) return undefined;
  const xml = ctx.getPartXml(rel.target);
  if (!xml) return undefined;
  // Header/footer parts have their own rels — resolve their images/hyperlinks
  // against the header/footer part, not word/document.xml.
  const partCtx = ctx.forPart(rel.target);
  const container = child(xml, root) ?? xml;
  return parseBlocks(container, partCtx);
}

function readSectPr(sectPr: XmlNode | undefined): {
  size: DocxPageSize;
  margins: DocxPageMargins;
  titlePg: boolean;
} {
  if (!sectPr) return { size: { ...DEFAULT_PAGE }, margins: { ...DEFAULT_MARGINS }, titlePg: false };

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
    titlePg: child(sectPr, 'titlePg') !== undefined,
  };
}
