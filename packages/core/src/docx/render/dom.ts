/**
 * DOM renderer: DocxPage model -> HTML/CSS. Consumes only the model (never
 * XML). Each page is a fixed-width sheet with margin padding; the viewer scales
 * pages to fit the container width and stacks them vertically.
 */
import { ptToPx } from '../../oxml/units.js';
import type {
  DocxPage,
  DocxBlock,
  DocxParagraph,
  DocxTable,
  DocxTableCell,
  DocxInlineImage,
  DocxFloat,
  DocxRun,
  Stroke,
  TextAlign,
} from '../model.js';

export interface RenderDeps {
  /** Resolve an embedded media part path to a displayable URL. */
  imageUrl(part: string): string | undefined;
}

const px = (n: number) => `${n}px`;

const ALIGN: Record<TextAlign, string> = { l: 'left', ctr: 'center', r: 'right', just: 'justify' };

/** Base typographic context shared by the sheet and the paginator's measurer. */
export const SHEET_FONT = {
  fontFamily: 'Calibri, "Segoe UI", Arial, sans-serif',
  fontSize: '16px',
  lineHeight: '1.15',
} as const;

/**
 * Render one paginated page as a fixed-size sheet: content flows in the padding
 * box between the margins, with the header drawn in the top margin band and the
 * footer in the bottom margin band (matching Word).
 */
export function renderPage(page: DocxPage, deps: RenderDeps): HTMLDivElement {
  const sheet = document.createElement('div');
  const s = sheet.style;
  s.position = 'relative';
  s.boxSizing = 'border-box';
  s.width = px(page.size.wPx);
  s.height = px(page.size.hPx);
  s.paddingTop = px(page.margins.topPx);
  s.paddingRight = px(page.margins.rightPx);
  s.paddingBottom = px(page.margins.bottomPx);
  s.paddingLeft = px(page.margins.leftPx);
  s.background = '#fff';
  s.color = '#000';
  s.margin = '0 auto';
  s.boxShadow = '0 1px 6px rgba(0,0,0,0.35)';
  s.fontFamily = SHEET_FONT.fontFamily;
  s.fontSize = SHEET_FONT.fontSize;
  s.lineHeight = SHEET_FONT.lineHeight;
  s.overflow = 'hidden';
  s.overflowWrap = 'break-word';

  const contentW = page.size.wPx - page.margins.leftPx - page.margins.rightPx;

  // Floating (anchored) images: absolutely positioned in page coordinates,
  // painted first so behindDoc banners sit under the content.
  if (page.floats) {
    for (const f of page.floats) sheet.appendChild(renderFloat(f, deps));
  }

  if (page.header && page.header.length) {
    sheet.appendChild(marginBand(page.header, deps, contentW, page.margins.leftPx, { top: px(page.margins.headerPx) }));
  }
  if (page.footer && page.footer.length) {
    sheet.appendChild(marginBand(page.footer, deps, contentW, page.margins.leftPx, { bottom: px(page.margins.footerPx) }));
  }

  for (const block of page.elements) sheet.appendChild(renderBlock(block, deps));
  return sheet;
}

/** A header/footer positioned inside the page's top/bottom margin band. */
function marginBand(
  blocks: DocxBlock[],
  deps: RenderDeps,
  contentW: number,
  leftPx: number,
  pos: { top?: string; bottom?: string },
): HTMLDivElement {
  const band = document.createElement('div');
  const s = band.style;
  s.position = 'absolute';
  s.left = px(leftPx);
  s.width = px(contentW);
  if (pos.top) s.top = pos.top;
  if (pos.bottom) s.bottom = pos.bottom;
  s.color = '#000';
  for (const block of blocks) band.appendChild(renderBlock(block, deps));
  return band;
}

/** Render a single block to a DOM element (also used by the paginator to measure). */
export function renderBlock(block: DocxBlock, deps: RenderDeps): HTMLElement {
  switch (block.kind) {
    case 'paragraph':
      return renderParagraph(block);
    case 'table':
      return renderTable(block, deps);
    case 'image':
      return renderImage(block, deps);
  }
}

function renderParagraph(p: DocxParagraph): HTMLElement {
  const el = document.createElement('div');
  const s = el.style;
  s.position = 'relative';

  if (p.align) s.textAlign = ALIGN[p.align];
  if (p.indentLeftPx !== undefined) s.marginLeft = px(p.indentLeftPx);
  if (p.indentRightPx !== undefined) s.marginRight = px(p.indentRightPx);
  if (p.indentFirstLinePx !== undefined) s.textIndent = px(p.indentFirstLinePx);
  if (p.spaceBeforePt !== undefined) s.marginTop = px(ptToPx(p.spaceBeforePt));
  if (p.spaceAfterPt !== undefined) s.marginBottom = px(ptToPx(p.spaceAfterPt));
  if (p.lineSpacingPct !== undefined) s.lineHeight = String(p.lineSpacingPct);
  else if (p.lineSpacingPt !== undefined) s.lineHeight = px(ptToPx(p.lineSpacingPt));

  if (p.baseFontFamily) s.fontFamily = quoteFont(p.baseFontFamily);
  if (p.baseFontSizePt !== undefined) s.fontSize = px(ptToPx(p.baseFontSizePt));
  if (p.baseBold) s.fontWeight = 'bold';
  if (p.baseItalic) s.fontStyle = 'italic';
  if (p.baseColorHex) s.color = `#${p.baseColorHex}`;
  if (p.shadingHex) s.background = `#${p.shadingHex}`;

  applyBorders(el, p.paraBorders);

  // List marker: sits before the text; the paragraph's hanging indent (negative
  // text-indent) pulls it into the left margin gutter, matching Word.
  if (p.listMarker) {
    const marker = document.createElement('span');
    marker.textContent = p.listMarker;
    marker.style.marginRight = '0.4em';
    el.appendChild(marker);
  }

  if (p.runs.length === 0) {
    // Preserve the height of an empty line.
    el.appendChild(document.createTextNode('​'));
  } else {
    for (const run of p.runs) appendRun(el, run);
  }
  return el;
}

function appendRun(parent: HTMLElement, run: DocxRun): void {
  if (run.breakBefore) parent.appendChild(document.createElement('br'));
  if (run.tabBefore) {
    const tab = document.createElement('span');
    tab.style.display = 'inline-block';
    tab.style.width = '0.5in';
    parent.appendChild(tab);
  }
  if (!run.text) return;

  const span = document.createElement('span');
  span.textContent = run.text;
  const s = span.style;
  if (run.bold) s.fontWeight = 'bold';
  if (run.italic) s.fontStyle = 'italic';
  const deco: string[] = [];
  if (run.underline) deco.push('underline');
  if (run.strike) deco.push('line-through');
  if (deco.length) s.textDecoration = deco.join(' ');
  if (run.sizePt !== undefined) s.fontSize = px(ptToPx(run.sizePt));
  if (run.colorHex) s.color = `#${run.colorHex}`;
  if (run.highlightHex) s.background = `#${run.highlightHex}`;
  if (run.font) s.fontFamily = quoteFont(run.font);
  if (run.caps === 'all') s.textTransform = 'uppercase';
  else if (run.caps === 'small') s.fontVariant = 'small-caps';
  if (run.vertAlign === 'super') {
    s.verticalAlign = 'super';
    s.fontSize = 'smaller';
  } else if (run.vertAlign === 'sub') {
    s.verticalAlign = 'sub';
    s.fontSize = 'smaller';
  }

  if (run.hyperlink) {
    const a = document.createElement('a');
    a.href = run.hyperlink;
    a.target = '_blank';
    a.rel = 'noopener noreferrer';
    a.style.color = run.colorHex ? `#${run.colorHex}` : '#0563c1';
    a.appendChild(span);
    parent.appendChild(a);
  } else {
    parent.appendChild(span);
  }
}

function renderTable(table: DocxTable, deps: RenderDeps): HTMLElement {
  const t = document.createElement('table');
  t.style.borderCollapse = 'collapse';
  t.style.tableLayout = 'fixed';
  if (table.widthPx) t.style.width = px(table.widthPx);

  if (table.colWidths.length) {
    const cg = document.createElement('colgroup');
    for (const w of table.colWidths) {
      const col = document.createElement('col');
      if (w) col.style.width = px(w);
      cg.appendChild(col);
    }
    t.appendChild(cg);
  }

  const tbody = document.createElement('tbody');
  for (const row of table.rows) {
    const tr = document.createElement('tr');
    for (const cell of row) {
      if (cell === null) continue; // covered by a span
      tr.appendChild(renderCell(cell, deps));
    }
    tbody.appendChild(tr);
  }
  t.appendChild(tbody);
  return t;
}

function renderCell(cell: DocxTableCell, deps: RenderDeps): HTMLElement {
  const td = document.createElement('td');
  if (cell.colSpan > 1) td.colSpan = cell.colSpan;
  if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
  const s = td.style;
  s.verticalAlign = cell.vAlign ?? 'top';
  if (cell.fillHex) s.background = `#${cell.fillHex}`;
  if (cell.cellPaddingPx) {
    const p = cell.cellPaddingPx;
    s.padding = `${p.top}px ${p.right}px ${p.bottom}px ${p.left}px`;
  } else {
    s.padding = '2px 5px';
  }
  if (cell.borders) {
    if (cell.borders.l) s.borderLeft = strokeCss(cell.borders.l);
    if (cell.borders.t) s.borderTop = strokeCss(cell.borders.t);
    if (cell.borders.r) s.borderRight = strokeCss(cell.borders.r);
    if (cell.borders.b) s.borderBottom = strokeCss(cell.borders.b);
  }

  for (const block of cell.content) td.appendChild(renderBlock(block, deps));
  if (!cell.content.length) td.appendChild(document.createTextNode('​'));
  return td;
}

function renderFloat(f: DocxFloat, deps: RenderDeps): HTMLElement {
  const url = deps.imageUrl(f.part);
  const el = document.createElement(url ? 'img' : 'div') as HTMLElement;
  const s = el.style;
  s.position = 'absolute';
  s.left = px(f.xPx);
  s.top = px(f.yPx);
  s.width = px(f.wPx);
  s.height = px(f.hPx);
  s.zIndex = f.behindDoc ? '0' : '2';
  if (url) {
    (el as HTMLImageElement).src = url;
    if (f.alt) (el as HTMLImageElement).alt = f.alt;
  }
  return el;
}

function renderImage(img: DocxInlineImage, deps: RenderDeps): HTMLElement {
  const url = deps.imageUrl(img.part);
  if (!url) {
    const ph = document.createElement('div');
    ph.textContent = img.alt ?? '[image]';
    ph.style.cssText = 'color:#999;font-style:italic;padding:4px 0;';
    return ph;
  }
  const el = document.createElement('img');
  el.src = url;
  if (img.widthPx) el.style.width = px(img.widthPx);
  if (img.heightPx) el.style.height = px(img.heightPx);
  el.style.maxWidth = '100%';
  if (img.alt) el.alt = img.alt;
  el.style.display = 'block';
  return el;
}

// ─── Helpers ───────────────────────────────────────────────────────────────────

function applyBorders(el: HTMLElement, borders: DocxParagraph['paraBorders']): void {
  if (!borders) return;
  if (borders.top) el.style.borderTop = strokeCss(borders.top);
  if (borders.bottom) el.style.borderBottom = strokeCss(borders.bottom);
  if (borders.left) el.style.borderLeft = strokeCss(borders.left);
  if (borders.right) el.style.borderRight = strokeCss(borders.right);
}

function strokeCss(stroke: Stroke): string {
  return `${Math.max(1, stroke.width)}px solid #${stroke.color.hex}`;
}

/** Known serif families, so an uninstalled font falls back to the right generic. */
const SERIF = /times|georgia|cambria|garamond|minion|book antiqua|palatino|serif|roman|constantia/i;

/**
 * Build a CSS font stack with a generic fallback. Word's default fonts (Calibri,
 * etc.) are rarely installed on non-Windows machines; without a generic the
 * browser falls back to serif, which looks nothing like Word. Pick serif vs
 * sans-serif by family name so the substitute is close.
 */
function quoteFont(name: string): string {
  const quoted = /\s/.test(name) ? `"${name}"` : name;
  const generic = SERIF.test(name) ? 'serif' : 'sans-serif';
  return `${quoted}, ${generic}`;
}
