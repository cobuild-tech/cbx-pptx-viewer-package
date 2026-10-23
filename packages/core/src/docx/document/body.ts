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
import { parseParagraph, applyContextualSpacing } from '../paragraphs/paragraph.js';
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
  let inherited: HeaderFooter = {};
  /** Empty paragraphs whose only content is a section break. */
  const breakMarks = new WeakSet<DocxBlock>();

  const flush = (sectPr: XmlNode | undefined) => {
    const { size, margins, titlePg, extra } = readSectPr(sectPr);
    // A section that does not reference a header/footer of some type carries
    // on with the previous section's.
    const own = readHeaderFooter(sectPr, ctx);
    const hf: HeaderFooter = {};
    for (const key of HF_KEYS) {
      const blocks = key in own ? own[key] : inherited[key];
      if (blocks) hf[key] = blocks;
    }
    inherited = hf;
    const shown: HeaderFooter = {};
    for (const key of HF_KEYS) if (hf[key]?.length) shown[key] = hf[key];
    applyContextualSpacing(current);
    // Size top-level tables against the section: percentage widths resolve
    // against the column, and only a table running off the page is shrunk.
    const contentW = size.wPx - margins.leftPx - margins.rightPx;
    for (const block of current) {
      if (block.kind === 'table') fitTableWidth(block, contentW, size.wPx - margins.leftPx);
      else if (block.kind === 'image') fitImageWidth(block, contentW);
    }
    sections.push({
      index: sections.length,
      size,
      margins,
      blocks: current,
      ...(titlePg ? { titlePg } : {}),
      ...(ctx.settings.evenAndOddHeaders ? { evenAndOddHeaders: true } : {}),
      ...extra,
      ...shown,
    });
    current = [];
  };

  for (const node of logicalChildren(body)) {
    const name = localName(node.name);
    if (name === 'p') {
      const parsed = parseParagraph(node, ctx);
      current.push(...parsed);
      const sectPr = child(child(node, 'pPr'), 'sectPr');
      if (sectPr) {
        const mark = parsed[parsed.length - 1];
        if (mark?.kind === 'paragraph' && !mark.runs.some((r) => r.text || r.drawing || r.math)) breakMarks.add(mark);
        flush(sectPr);
      }
    } else if (name === 'tbl') {
      current.push(parseTable(node, ctx));
    } else if (name === 'commentRangeStart') {
      ctx.openComments?.add(attr(node, 'id') ?? '');
    } else if (name === 'commentRangeEnd') {
      ctx.openComments?.delete(attr(node, 'id') ?? '');
    } else if (name === 'sectPr') {
      // Final section's properties (direct body child).
      flush(node);
    }
  }

  // Body ended without a trailing sectPr child (unusual) — emit what's left.
  if (current.length > 0 || sections.length === 0) flush(undefined);

  // An empty paragraph that only carries a section break takes no room when
  // the next section continues on the same page: Word lays the two sections
  // out back to back.
  for (let i = 1; i < sections.length; i++) {
    if (sections[i]!.start !== 'continuous') continue;
    const prev = sections[i - 1]!.blocks;
    const last = prev[prev.length - 1];
    if (last && breakMarks.has(last)) prev.pop();
  }

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
  applyContextualSpacing(out);
  return out;
}

type HeaderFooter = Partial<
  Record<'header' | 'footer' | 'firstHeader' | 'firstFooter' | 'evenHeader' | 'evenFooter', DocxBlock[]>
>;
const HF_KEYS = ['header', 'footer', 'firstHeader', 'firstFooter', 'evenHeader', 'evenFooter'] as const;

/**
 * The header/footer parts a sectPr references. A key is present when the
 * section has a reference of that type — even one resolving to nothing — and
 * absent when it has none, which is what inheritance keys off.
 */
function readHeaderFooter(sectPr: XmlNode | undefined, ctx: ParseContext): HeaderFooter {
  if (!sectPr) return {};
  const out: HeaderFooter = {};
  const hdrs = children(sectPr, 'headerReference');
  const ftrs = children(sectPr, 'footerReference');
  const set = (key: keyof HeaderFooter, refs: XmlNode[], type: 'default' | 'first' | 'even', root: string) => {
    const blocks = resolveRef(refs, type, root, ctx);
    if (blocks) out[key] = blocks;
  };
  set('header', hdrs, 'default', 'hdr');
  set('firstHeader', hdrs, 'first', 'hdr');
  set('footer', ftrs, 'default', 'ftr');
  set('firstFooter', ftrs, 'first', 'ftr');
  set('evenHeader', hdrs, 'even', 'hdr');
  set('evenFooter', ftrs, 'even', 'ftr');
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
  if (!rel) return [];
  const xml = ctx.getPartXml(rel.target);
  if (!xml) return [];
  // Header/footer parts have their own rels — resolve their images/hyperlinks
  // against the header/footer part, not word/document.xml.
  const partCtx = ctx.forPart(rel.target);
  const container = child(xml, root) ?? xml;
  return parseBlocks(container, partCtx);
}

type SectionExtra = Pick<DocxSection, 'start' | 'pageNumberStart' | 'pageNumberFormat' | 'columns'>;

function readSectPr(sectPr: XmlNode | undefined): {
  size: DocxPageSize;
  margins: DocxPageMargins;
  titlePg: boolean;
  extra: SectionExtra;
} {
  if (!sectPr) return { size: { ...DEFAULT_PAGE }, margins: { ...DEFAULT_MARGINS }, titlePg: false, extra: {} };

  const extra: SectionExtra = {};
  const type = attr(child(sectPr, 'type'), 'val');
  if (type === 'continuous' || type === 'evenPage' || type === 'oddPage' || type === 'nextColumn' || type === 'nextPage') {
    extra.start = type;
  }
  const pgNumType = child(sectPr, 'pgNumType');
  const numStart = attrNum(pgNumType, 'start');
  if (numStart !== undefined) extra.pageNumberStart = numStart;
  const numFmt = attr(pgNumType, 'fmt');
  if (numFmt) extra.pageNumberFormat = numFmt;

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
    titlePg: child(sectPr, 'titlePg') !== undefined && attr(child(sectPr, 'titlePg'), 'val') !== '0',
    extra: { ...extra, ...readColumns(child(sectPr, 'cols'), twipToPx(w) - twipToPx(mar('left', 1440)) - twipToPx(mar('right', 1440))) },
  };
}

/**
 * <w:cols>: equal columns (num, space) or explicit ones (<w:col w:w w:space>),
 * laid out left to right across the text area.
 */
function readColumns(cols: XmlNode | undefined, textW: number): Pick<DocxSection, 'columns'> {
  if (!cols) return {};
  const sep = attr(cols, 'sep');
  const separator = sep !== undefined && sep !== '0' && sep !== 'false';
  const explicit = children(cols, 'col');
  const equal = attr(cols, 'equalWidth');
  const out: { xPx: number; widthPx: number }[] = [];
  if (explicit.length > 1 && equal !== '1' && equal !== 'true') {
    let x = 0;
    for (const c of explicit) {
      const w = twipToPx(attrNum(c, 'w') ?? 0);
      out.push({ xPx: x, widthPx: w });
      x += w + twipToPx(attrNum(c, 'space') ?? 0);
    }
  } else {
    const n = attrNum(cols, 'num') ?? 1;
    if (n <= 1) return {};
    const space = twipToPx(attrNum(cols, 'space') ?? 720);
    const w = (textW - space * (n - 1)) / n;
    for (let i = 0; i < n; i++) out.push({ xPx: i * (w + space), widthPx: w });
  }
  return out.length > 1 ? { columns: { cols: out, separator } } : {};
}
