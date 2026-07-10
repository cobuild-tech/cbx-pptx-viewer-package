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
  DocxDrawing,
  DocxShape,
  DocxAnchor,
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

// relativeFrom values framed to the page/margins (not the text flow). An anchor
// whose VERTICAL reference is one of these is positioned in page coordinates on
// the sheet — its flow position doesn't determine where it sits (e.g. a
// full-page-width header banner anchored relativeFrom="page"). Anchors framed to
// the paragraph/line/column stay attached to their paragraph instead.
const PAGE_FRAME = new Set([
  'page',
  'margin',
  'topMargin',
  'bottomMargin',
  'leftMargin',
  'rightMargin',
  'insideMargin',
  'outsideMargin',
]);

/** Whether an anchor is positioned against the page frame (vs. the text flow). */
export function isPageAnchored(a: DocxAnchor): boolean {
  return PAGE_FRAME.has(a.relV);
}

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
  // Body content box: grown past the raw margins when a tall header/footer
  // would otherwise overlap it (Word reserves space for the banner/footer).
  s.paddingTop = px(page.contentTopPx ?? page.margins.topPx);
  s.paddingRight = px(page.margins.rightPx);
  s.paddingBottom = px(page.contentBottomPx ?? page.margins.bottomPx);
  s.paddingLeft = px(page.margins.leftPx);
  s.background = '#fff';
  s.color = '#000';
  s.margin = '0 auto';
  s.boxShadow = '0 1px 6px rgba(0,0,0,0.35)';
  s.fontFamily = SHEET_FONT.fontFamily;
  s.fontSize = SHEET_FONT.fontSize;
  s.lineHeight = SHEET_FONT.lineHeight;
  s.overflow = 'hidden';
  s.overflowWrap = 'normal';

  const contentW = page.size.wPx - page.margins.leftPx - page.margins.rightPx;

  // Page-framed anchors (banners relativeFrom="page"/margins) painted first, in
  // page coordinates on the sheet, so behindDoc ones sit under the content and
  // full-bleed banners ignore the text margins.
  const pageAnchors: DocxAnchor[] = [];
  collectPageAnchors(page.elements, pageAnchors);
  if (page.header) collectPageAnchors(page.header, pageAnchors);
  if (page.footer) collectPageAnchors(page.footer, pageAnchors);
  for (const a of pageAnchors) sheet.appendChild(renderSheetAnchor(a, deps, page));

  if (page.header && page.header.length) {
    sheet.appendChild(marginBand(page.header, deps, contentW, page.margins.leftPx, { top: px(page.margins.headerPx) }, page.index, page.resolvedStyles));
  }
  if (page.footer && page.footer.length) {
    sheet.appendChild(marginBand(page.footer, deps, contentW, page.margins.leftPx, { bottom: px(page.margins.footerPx) }, page.index, page.resolvedStyles));
  }

  for (const block of page.elements) sheet.appendChild(renderBlock(block, deps, page.index, page.resolvedStyles));
  return sheet;
}

/** Gather page-framed anchors from a block tree (paragraph anchors + cells). */
function collectPageAnchors(blocks: DocxBlock[], out: DocxAnchor[]): void {
  for (const b of blocks) {
    if (b.kind === 'paragraph') {
      for (const a of b.anchors ?? []) if (isPageAnchored(a)) out.push(a);
    } else if (b.kind === 'table') {
      for (const row of b.rows) for (const cell of row) if (cell) collectPageAnchors(cell.content, out);
    }
  }
}

/** A header/footer positioned inside the page's top/bottom margin band. */
function marginBand(
  blocks: DocxBlock[],
  deps: RenderDeps,
  contentW: number,
  leftPx: number,
  pos: { top?: string; bottom?: string },
  pageIndex?: number,
  resolvedStyles?: Record<string, string>,
): HTMLDivElement {
  const band = document.createElement('div');
  const s = band.style;
  s.position = 'absolute';
  s.left = px(leftPx);
  s.width = px(contentW);
  if (pos.top) s.top = pos.top;
  if (pos.bottom) s.bottom = pos.bottom;
  s.color = '#000';
  for (const block of blocks) band.appendChild(renderBlock(block, deps, pageIndex, resolvedStyles));
  return band;
}

/** Render a single block to a DOM element (also used by the paginator to measure). */
export function renderBlock(
  block: DocxBlock,
  deps: RenderDeps,
  pageIndex?: number,
  resolvedStyles?: Record<string, string>,
): HTMLElement {
  switch (block.kind) {
    case 'paragraph':
      return renderParagraph(block, deps, pageIndex, resolvedStyles);
    case 'table':
      return renderTable(block, deps, pageIndex, resolvedStyles);
    case 'image':
      return renderImage(block, deps);
  }
}

function renderParagraph(
  p: DocxParagraph,
  deps: RenderDeps,
  pageIndex?: number,
  resolvedStyles?: Record<string, string>,
  flowAnchorSink?: DocxAnchor[],
): HTMLElement {
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
    const hasTab = p.runs.some((r) => r.tabBefore);
    if (hasTab) {
      s.display = 'flex';
      s.justifyContent = 'space-between';
      s.alignItems = 'center';
      s.width = '100%';

      // Group runs into segments separated by tabs
      const segments: DocxRun[][] = [];
      let currentSegment: DocxRun[] = [];
      for (const run of p.runs) {
        if (run.tabBefore) {
          segments.push(currentSegment);
          currentSegment = [];
        }
        currentSegment.push(run);
      }
      if (currentSegment.length > 0) {
        segments.push(currentSegment);
      }

      // Render each segment as a flex item. The segment itself is inline-block
      // (not inline-flex) so its runs flow as normal inline text — otherwise each
      // run becomes a flex item and flexbox trims the trailing whitespace at the
      // item boundary, collapsing e.g. "Page " + "1" into "Page1".
      for (const segment of segments) {
        const segEl = document.createElement('span');
        segEl.style.display = 'inline-block';
        segEl.style.verticalAlign = 'middle';
        for (const run of segment) {
          const cleanRun = { ...run, tabBefore: false };
          appendRun(segEl, cleanRun, deps, pageIndex, resolvedStyles);
        }
        el.appendChild(segEl);
      }
    } else {
      for (const run of p.runs) appendRun(el, run, deps, pageIndex, resolvedStyles);
    }
  }

  // Flow-framed anchors (relativeFrom column/paragraph) position within this
  // paragraph. Page-framed anchors are hoisted to the sheet by renderPage. When
  // a sink is supplied (paragraph inside a table cell), flow anchors are hoisted
  // to the cell instead, so a vertically-centered cell doesn't drag them down.
  if (p.anchors) {
    for (const a of p.anchors) {
      if (isPageAnchored(a)) continue;
      if (flowAnchorSink) flowAnchorSink.push(a);
      else el.appendChild(renderAnchor(a, deps));
    }
  }
  return el;
}

function appendRun(
  parent: HTMLElement,
  run: DocxRun,
  deps: RenderDeps,
  pageIndex?: number,
  resolvedStyles?: Record<string, string>,
): void {
  if (run.breakBefore) parent.appendChild(document.createElement('br'));
  if (run.tabBefore) {
    const tab = document.createElement('span');
    tab.style.display = 'inline-block';
    tab.style.width = '0.5in';
    parent.appendChild(tab);
  }

  // An inline drawing (picture / shape) occupying this run's slot.
  if (run.drawing) {
    parent.appendChild(renderDrawing(run.drawing, deps));
    return;
  }

  let displayText = run.text;
  if (run.fieldCode) {
    const code = run.fieldCode.trim();
    if (/^page$/i.test(code)) {
      displayText = pageIndex !== undefined ? String(pageIndex + 1) : '1';
    } else {
      const match = /^\s*styleref\s+(?:"([^"]+)"|'([^']+)'|([^\s]+))/i.exec(code);
      if (match) {
        const styleName = (match[1] || match[2] || match[3] || '').toLowerCase();
        displayText = resolvedStyles?.[styleName] ?? run.text;
      }
    }
  }

  if (!displayText) return;

  const span = document.createElement('span');
  span.textContent = displayText;
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

function renderTable(
  table: DocxTable,
  deps: RenderDeps,
  pageIndex?: number,
  resolvedStyles?: Record<string, string>,
): HTMLElement {
  const t = document.createElement('table');
  t.style.borderCollapse = 'collapse';
  t.style.tableLayout = 'fixed';
  if (table.widthPx) t.style.width = px(table.widthPx);
  if (table.indentPx) t.style.marginLeft = px(table.indentPx);

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
      tr.appendChild(renderCell(cell, deps, pageIndex, resolvedStyles));
    }
    tbody.appendChild(tr);
  }
  t.appendChild(tbody);
  return t;
}

function renderCell(
  cell: DocxTableCell,
  deps: RenderDeps,
  pageIndex?: number,
  resolvedStyles?: Record<string, string>,
): HTMLElement {
  const td = document.createElement('td');
  if (cell.colSpan > 1) td.colSpan = cell.colSpan;
  if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
  const s = td.style;
  // Positioning context for the cell's floating anchors (hoisted off their
  // paragraphs so a centered cell doesn't shift them).
  s.position = 'relative';
  s.verticalAlign = cell.vAlign ?? 'top';
  if (cell.fillHex) {
    s.background = `#${cell.fillHex}`;
    // Word's "automatic" font color flips to white on a dark fill for contrast.
    // Set it as the cell default so explicit run/paragraph colors still win.
    if (isDarkFill(cell.fillHex)) s.color = '#fff';
  }
  if (cell.cellPaddingPx) {
    const p = cell.cellPaddingPx;
    s.padding = `${p.top}px ${p.right}px ${p.bottom}px ${p.left}px`;
  } else {
    s.padding = '2px 5px';
  }
  s.wordBreak = 'keep-all';
  s.overflowWrap = 'normal';
  if (cell.borders) {
    if (cell.borders.l) s.borderLeft = strokeCss(cell.borders.l);
    if (cell.borders.t) s.borderTop = strokeCss(cell.borders.t);
    if (cell.borders.r) s.borderRight = strokeCss(cell.borders.r);
    if (cell.borders.b) s.borderBottom = strokeCss(cell.borders.b);
  }

  const flowAnchors: DocxAnchor[] = [];
  for (const block of cell.content) {
    if (block.kind === 'paragraph') {
      td.appendChild(renderParagraph(block, deps, pageIndex, resolvedStyles, flowAnchors));
    } else {
      td.appendChild(renderBlock(block, deps, pageIndex, resolvedStyles));
    }
  }
  // Cell-hoisted floating anchors, positioned against the cell (top-aligned).
  for (const a of flowAnchors) td.appendChild(renderAnchor(a, deps));
  if (!cell.content.length) td.appendChild(document.createTextNode('​'));
  return td;
}

/**
 * Render a floating (anchored) drawing absolutely inside its containing,
 * position:relative paragraph/cell. `left`/`top` resolve the anchor's
 * relativeFrom + offset/align against the paragraph's content box using CSS
 * calc — so an image anchored in a right-hand cell lands on the right without
 * the renderer needing to know the box width.
 */
function renderAnchor(anchor: DocxAnchor, deps: RenderDeps): HTMLElement {
  const el = renderDrawing(anchor.drawing, deps);
  const s = el.style;
  s.position = 'absolute';
  s.margin = '0';
  s.width = px(anchor.wPx);
  s.height = px(anchor.hPx);
  s.maxWidth = 'none';
  s.left = anchorLeft(anchor);
  s.top = anchorTop(anchor);
  // behindDoc drawings sit under the text (negative z); others float above it.
  s.zIndex = anchor.behindDoc ? '-1' : '1';
  return el;
}

/**
 * Render a page-framed anchor (relativeFrom page/margins) in absolute page
 * coordinates on the sheet, so full-bleed banners ignore the text margins and
 * relativeFrom="page" offsets are measured from the page edge.
 */
function renderSheetAnchor(anchor: DocxAnchor, deps: RenderDeps, page: DocxPage): HTMLElement {
  const el = renderDrawing(anchor.drawing, deps);
  const s = el.style;
  const { size, margins } = page;
  const contentW = size.wPx - margins.leftPx - margins.rightPx;
  s.position = 'absolute';
  s.margin = '0';
  s.width = px(anchor.wPx);
  s.height = px(anchor.hPx);
  s.maxWidth = 'none';

  // Horizontal.
  let left: number;
  if (anchor.hAlign === 'center') {
    left = anchor.relH === 'page' ? (size.wPx - anchor.wPx) / 2 : margins.leftPx + (contentW - anchor.wPx) / 2;
  } else if (anchor.hAlign === 'right' || anchor.hAlign === 'outside') {
    left = anchor.relH === 'page' ? size.wPx - anchor.wPx : size.wPx - margins.rightPx - anchor.wPx;
  } else if (anchor.hAlign === 'left' || anchor.hAlign === 'inside') {
    left = anchor.relH === 'page' ? 0 : margins.leftPx;
  } else {
    const off = anchor.hOffsetPx ?? 0;
    if (anchor.relH === 'page') left = off;
    else if (anchor.relH === 'rightMargin' || anchor.relH === 'outsideMargin') left = size.wPx - margins.rightPx + off;
    else left = margins.leftPx + off; // leftMargin / margin / column / character
  }

  // Vertical.
  let top: number;
  if (anchor.vAlign === 'center') {
    top = (size.hPx - anchor.hPx) / 2;
  } else if (anchor.vAlign === 'bottom') {
    top = anchor.relV === 'page' ? size.hPx - anchor.hPx : size.hPx - margins.bottomPx - anchor.hPx;
  } else if (anchor.vAlign === 'top') {
    top = anchor.relV === 'page' ? 0 : margins.topPx;
  } else {
    const off = anchor.vOffsetPx ?? 0;
    if (anchor.relV === 'page') top = off;
    else if (anchor.relV === 'bottomMargin' || anchor.relV === 'outsideMargin') top = size.hPx - margins.bottomPx + off;
    else top = margins.topPx + off; // topMargin / margin / text
  }

  s.left = px(left);
  s.top = px(top);
  s.zIndex = anchor.behindDoc ? '-1' : '5';
  return el;
}

/** CSS `left` for an anchor, relative to the containing box's content edge. */
function anchorLeft(a: DocxAnchor): string {
  const w = a.wPx;
  if (a.hAlign === 'center') return `calc((100% - ${w}px) / 2)`;
  if (a.hAlign === 'right' || a.hAlign === 'outside') return `calc(100% - ${w}px)`;
  if (a.hAlign === 'left' || a.hAlign === 'inside') return '0px';
  const off = a.hOffsetPx ?? 0;
  // rightMargin/outsideMargin measure from the right content edge.
  if (a.relH === 'rightMargin' || a.relH === 'outsideMargin') return `calc(100% + ${off}px)`;
  return px(off);
}

/** CSS `top` for an anchor, relative to the containing paragraph's top. */
function anchorTop(a: DocxAnchor): string {
  if (a.vAlign === 'center') return `calc((100% - ${a.hPx}px) / 2)`;
  if (a.vAlign === 'bottom') return `calc(100% - ${a.hPx}px)`;
  if (a.vAlign === 'top') return '0px';
  return px(a.vOffsetPx ?? 0);
}

/** A drawing (picture or DrawingML shape) as an inline-block element. */
function renderDrawing(drawing: DocxDrawing, deps: RenderDeps): HTMLElement {
  return drawing.kind === 'shape' ? renderShape(drawing, deps) : renderImage(drawing, deps);
}

function renderImage(img: DocxInlineImage, deps: RenderDeps): HTMLElement {
  const url = deps.imageUrl(img.part);
  if (!url) {
    const ph = document.createElement('div');
    ph.textContent = img.alt ?? '[image]';
    ph.style.cssText = 'color:#999;font-style:italic;padding:4px 0;display:inline-block;';
    return ph;
  }
  const el = document.createElement('img');
  el.src = url;
  if (img.widthPx) el.style.width = px(img.widthPx);
  if (img.heightPx) el.style.height = px(img.heightPx);
  el.style.maxWidth = '100%';
  if (img.alt) el.alt = img.alt;
  el.style.display = 'inline-block';
  el.style.verticalAlign = 'middle';
  return el;
}

/** A DrawingML shape (filled/outlined box, optionally a text box). */
function renderShape(shape: DocxShape, deps: RenderDeps): HTMLElement {
  const el = document.createElement('div');
  const s = el.style;
  s.display = 'inline-block';
  s.verticalAlign = 'middle';
  s.boxSizing = 'border-box';
  if (shape.widthPx) s.width = px(shape.widthPx);
  if (shape.heightPx) s.height = px(shape.heightPx);
  s.maxWidth = '100%';

  if (shape.geom === 'line') {
    // A connector/line: render as a rule using the outline color.
    s.borderTop = `${shape.lineWidthPx ? Math.max(1, shape.lineWidthPx) : 1}px solid ${shape.lineHex ? `#${shape.lineHex}` : '#000'}`;
    s.height = px(shape.heightPx || 0);
    return el;
  }

  if (shape.fillHex) {
    s.background = `#${shape.fillHex}`;
    // Word's automatic text color flips to white on a dark fill (matches the
    // table-cell rule); explicit run/paragraph colors still win.
    if (isDarkFill(shape.fillHex)) s.color = '#fff';
  }
  if (shape.lineHex) s.border = `${shape.lineWidthPx ?? 1}px solid #${shape.lineHex}`;
  if (shape.geom === 'roundRect') s.borderRadius = px(Math.min(shape.widthPx, shape.heightPx) * 0.16);
  else if (shape.geom === 'ellipse') s.borderRadius = '50%';

  // Vertically anchor the text box content.
  s.display = 'inline-flex';
  s.flexDirection = 'column';
  s.justifyContent = shape.vAnchor === 'top' ? 'flex-start' : shape.vAnchor === 'bottom' ? 'flex-end' : 'center';
  const inner = document.createElement('div');
  inner.style.width = '100%';
  inner.style.padding = '0 0.1in';
  for (const block of shape.content) inner.appendChild(renderBlock(block, deps));
  el.appendChild(inner);
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

/** Perceived-luminance test for cell fills, to pick auto white/black text. */
function isDarkFill(hex: string): boolean {
  const h = hex.replace(/^#/, '');
  if (h.length < 6) return false;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return 0.299 * r + 0.587 * g + 0.114 * b < 128;
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
