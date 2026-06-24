/**
 * DOM renderer for DOCX documents.
 *
 * Converts DocxPage IR nodes into HTML elements. The page is a white card
 * sized to the section's paper dimensions (width fixed; height auto so all
 * content is visible). The DocxViewer scales it to fit the container by width.
 *
 * Rendering mirrors the PPTX approach — feature renderers are isolated per
 * block type and share the RenderDeps interface for media resolution.
 */
import type {
  DocxPage,
  DocxBlock,
  DocxParagraph,
  DocxTable,
  DocxTableCell,
  DocxInlineImage,
  TextRun,
  Fill,
  Stroke,
  Bullet,
} from '../model.js';
import type { RenderDeps } from '../../pptx/render/primitives.js';
import { colorToCss } from '../../pptx/color.js';

export type { RenderDeps } from '../../pptx/render/primitives.js';

// ─── Page ─────────────────────────────────────────────────────────────────────

/**
 * Render a single DocxPage into an HTML div.
 * Width is fixed to the page's paper width; height is auto (content-driven).
 */
export function renderPage(page: DocxPage, deps: RenderDeps): HTMLDivElement {
  const root = document.createElement('div');
  root.className = 'docx-page';
  root.style.position = 'relative';
  root.style.width = `${page.size.wPx}px`;
  root.style.minHeight = `${page.size.hPx}px`;
  root.style.background = 'white';
  root.style.color = '#000000';
  root.style.boxSizing = 'border-box';
  root.style.boxShadow = '0 2px 12px rgba(0,0,0,0.18)';
  root.style.overflowX = 'hidden';

  // Header — absolutely positioned so it doesn't push the body content down.
  // topPx is measured from the paper top and already includes the header zone.
  if (page.header?.length) {
    const headerEl = document.createElement('div');
    headerEl.className = 'docx-header';
    headerEl.style.position = 'absolute';
    headerEl.style.top = `${page.margins.headerPx}px`;
    headerEl.style.left = `${page.margins.leftPx}px`;
    headerEl.style.right = `${page.margins.rightPx}px`;
    headerEl.style.maxHeight = `${page.margins.topPx - page.margins.headerPx}px`;
    headerEl.style.overflow = 'hidden';
    for (const block of page.header) {
      const el = renderBlock(block, deps);
      if (el) headerEl.appendChild(el);
    }
    root.appendChild(headerEl);
  }

  // Content area — full page margins as padding; header/footer live above/below via position:absolute.
  const content = document.createElement('div');
  content.className = 'docx-content';
  content.style.paddingTop = `${page.margins.topPx}px`;
  content.style.paddingRight = `${page.margins.rightPx}px`;
  content.style.paddingBottom = `${page.margins.bottomPx}px`;
  content.style.paddingLeft = `${page.margins.leftPx}px`;

  let prevParaStyle: string | undefined;
  for (const block of page.elements) {
    const el = renderBlock(block, deps);
    if (el) {
      if (
        block.kind === 'paragraph' &&
        block.contextualSpacing &&
        prevParaStyle === block.styleName
      ) {
        el.style.marginTop = '0';
      }
      prevParaStyle = block.kind === 'paragraph' ? block.styleName : undefined;
      content.appendChild(el);
    }
  }
  root.appendChild(content);

  // Footer — absolutely positioned at the bottom of the page.
  if (page.footer?.length) {
    const footerEl = document.createElement('div');
    footerEl.className = 'docx-footer';
    footerEl.style.position = 'absolute';
    footerEl.style.bottom = `${page.margins.footerPx}px`;
    footerEl.style.left = `${page.margins.leftPx}px`;
    footerEl.style.right = `${page.margins.rightPx}px`;
    footerEl.style.maxHeight = `${page.margins.bottomPx - page.margins.footerPx}px`;
    footerEl.style.overflow = 'hidden';
    for (const block of page.footer) {
      const el = renderBlock(block, deps);
      if (el) footerEl.appendChild(el);
    }
    root.appendChild(footerEl);
  }

  return root;
}

// ─── Block dispatcher ─────────────────────────────────────────────────────────

export function renderBlock(block: DocxBlock, deps: RenderDeps): HTMLElement | null {
  switch (block.kind) {
    case 'paragraph': return renderParagraph(block, deps);
    case 'table':     return renderDocxTable(block, deps);
    case 'image':     return renderInlineImageBlock(block, deps);
  }
}

// ─── Paragraph ────────────────────────────────────────────────────────────────

function renderParagraph(para: DocxParagraph, deps: RenderDeps): HTMLDivElement {
  const el = document.createElement('div');
  el.className = `docx-para docx-style-${cssClass(para.styleName)}`;

  applyParagraphStyles(el, para);

  // Bullet/number prefix — rendered inline so the paragraph's own
  // indentLeftPx (paddingLeft) + indentFirstLinePx (textIndent) produce
  // the correct hanging-indent layout, identical to Word.
  if (para.bullet && para.bullet.type !== 'none') {
    const prefix = document.createElement('span');
    prefix.className = 'docx-bullet';
    prefix.style.userSelect = 'none';

    if (para.bullet.type === 'char') {
      prefix.textContent = para.bullet.char + ' ';
      if (para.bullet.font) prefix.style.fontFamily = `'${para.bullet.font}', sans-serif`;
      if (para.bullet.color) prefix.style.color = colorToCss(para.bullet.color);
    } else {
      prefix.textContent = '• ';
    }

    el.appendChild(prefix);
  }

  for (const run of para.runs) {
    el.appendChild(renderRun(run));
  }

  return el;
}

function applyParagraphStyles(el: HTMLElement, para: DocxParagraph): void {
  // Apply style-chain defaults; runs override per-span.
  if (para.baseFontFamily) {
    el.style.fontFamily = `'${para.baseFontFamily}', sans-serif`;
  }
  if (para.baseFontSizePt !== undefined) {
    el.style.fontSize = `${para.baseFontSizePt}pt`;
  }
  if (para.baseBold) el.style.fontWeight = 'bold';
  if (para.baseItalic) el.style.fontStyle = 'italic';
  if (para.baseColorHex) el.style.color = `#${para.baseColorHex}`;

  if (para.align) {
    const map: Record<string, string> = { l: 'left', ctr: 'center', r: 'right', just: 'justify' };
    el.style.textAlign = map[para.align] ?? 'left';
  }

  if (para.indentLeftPx) {
    el.style.paddingLeft = `${para.indentLeftPx}px`;
  }
  if (para.indentFirstLinePx !== undefined && para.indentFirstLinePx !== 0) {
    el.style.textIndent = `${para.indentFirstLinePx}px`;
  }
  if (para.indentRightPx) {
    el.style.paddingRight = `${para.indentRightPx}px`;
  }
  if (para.spaceBeforePt) el.style.marginTop = `${para.spaceBeforePt}pt`;
  if (para.spaceAfterPt) el.style.marginBottom = `${para.spaceAfterPt}pt`;

  if (para.lineSpacingPct !== undefined) {
    el.style.lineHeight = String(para.lineSpacingPct);
  } else if (para.lineSpacingPt !== undefined) {
    el.style.lineHeight = `${para.lineSpacingPt}pt`;
  }

  // Empty paragraphs (blank lines) should still take up space.
  if (para.runs.every((r) => !r.text.trim())) {
    el.style.minHeight = '1em';
  }

  if (para.shadingHex) {
    el.style.backgroundColor = `#${para.shadingHex}`;
  }
  if (para.paraBorders) {
    if (para.paraBorders.top) {
      el.style.borderTop = strokeToCssBorder(para.paraBorders.top);
      el.style.paddingTop = '1pt';
    }
    if (para.paraBorders.bottom) {
      el.style.borderBottom = strokeToCssBorder(para.paraBorders.bottom);
      el.style.paddingBottom = '1pt';
    }
    if (para.paraBorders.left) {
      el.style.borderLeft = strokeToCssBorder(para.paraBorders.left);
      el.style.paddingLeft = `${(para.indentLeftPx ?? 0) + 4}px`;
    }
    if (para.paraBorders.right) {
      el.style.borderRight = strokeToCssBorder(para.paraBorders.right);
    }
  }
}

// ─── Run ──────────────────────────────────────────────────────────────────────

function renderRun(run: TextRun): HTMLElement {
  const el = run.hyperlink
    ? document.createElement('a')
    : document.createElement('span');
  el.className = 'docx-run';

  if (run.text.includes('\n') || run.text.includes('\t')) {
    const parts = run.text.split(/(\n|\t)/);
    for (const part of parts) {
      if (part === '\n') {
        el.appendChild(document.createElement('br'));
      } else if (part === '\t') {
        // Approximate Word's 0.5-inch default tab stop
        const tab = document.createElement('span');
        tab.style.display = 'inline-block';
        tab.style.minWidth = '48px';
        el.appendChild(tab);
      } else if (part) {
        el.appendChild(document.createTextNode(part));
      }
    }
  } else {
    el.textContent = run.text;
  }

  applyRunStyles(el, run);

  if (run.hyperlink && el instanceof HTMLAnchorElement) {
    el.href = run.hyperlink;
    el.target = '_blank';
    el.rel = 'noopener noreferrer';
    // Color comes from the document's Hyperlink character style (already in run.color).
    // Only fall back to the Word default if no color was specified in the document.
    if (!run.color) el.style.color = '#0563C1';
  }

  return el;
}

function applyRunStyles(el: HTMLElement, run: TextRun): void {
  if (run.bold) el.style.fontWeight = 'bold';
  if (run.italic) el.style.fontStyle = 'italic';
  if (run.underline) el.style.textDecoration = 'underline';
  if (run.strike) el.style.textDecoration = (el.style.textDecoration ? el.style.textDecoration + ' ' : '') + 'line-through';
  if (run.sizePt !== undefined) el.style.fontSize = `${run.sizePt}pt`;
  if (run.color) el.style.color = colorToCss(run.color);
  if (run.font) el.style.fontFamily = `'${run.font}', sans-serif`;
  if (run.caps === 'all') el.style.textTransform = 'uppercase';
  if (run.caps === 'small') el.style.fontVariant = 'small-caps';
  if (run.highlight) el.style.backgroundColor = colorToCss(run.highlight);
  if (run.baseline !== undefined && run.baseline !== 0) {
    el.style.verticalAlign = run.baseline > 0 ? 'super' : 'sub';
    el.style.fontSize = '0.75em';
  }
  if (run.letterSpacingPt !== undefined) {
    el.style.letterSpacing = `${run.letterSpacingPt}pt`;
  }
}

// ─── Table ────────────────────────────────────────────────────────────────────

function renderDocxTable(table: DocxTable, deps: RenderDeps): HTMLTableElement {
  const tableEl = document.createElement('table');
  tableEl.className = 'docx-table';
  tableEl.style.borderCollapse = 'collapse';
  tableEl.style.tableLayout = 'fixed';
  if (table.widthPx) tableEl.style.width = `${table.widthPx}px`;

  const colGroup = document.createElement('colgroup');
  for (const w of table.colWidths) {
    const col = document.createElement('col');
    col.style.width = `${w}px`;
    colGroup.appendChild(col);
  }
  tableEl.appendChild(colGroup);

  const tbody = document.createElement('tbody');
  for (const row of table.rows) {
    const trEl = document.createElement('tr');
    for (const cell of row) {
      if (cell === null) continue;
      trEl.appendChild(renderTableCell(cell, deps));
    }
    tbody.appendChild(trEl);
  }
  tableEl.appendChild(tbody);
  return tableEl;
}

function renderTableCell(cell: DocxTableCell, deps: RenderDeps): HTMLTableCellElement {
  const td = document.createElement('td');
  td.className = 'docx-cell';
  if (cell.colSpan > 1) td.colSpan = cell.colSpan;
  if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;

  // Use document cell margins; fall back to DOCX default (0 top/bottom, ~7px left/right).
  if (cell.cellPaddingPx) {
    const { top, right, bottom, left } = cell.cellPaddingPx;
    td.style.padding = `${top}px ${right}px ${bottom}px ${left}px`;
  } else {
    td.style.padding = '0 7px';
  }
  td.style.verticalAlign = cell.vAlign ?? 'top';
  td.style.wordBreak = 'break-word';

  // Fill
  if (cell.fill.type === 'solid') {
    td.style.backgroundColor = colorToCss(cell.fill.color);
  }

  // Borders — fully document-driven; no fallback added by the renderer.
  if (cell.borders) {
    const { l, t, r, b } = cell.borders;
    if (l) td.style.borderLeft = strokeToCssBorder(l);
    if (t) td.style.borderTop = strokeToCssBorder(t);
    if (r) td.style.borderRight = strokeToCssBorder(r);
    if (b) td.style.borderBottom = strokeToCssBorder(b);
  }

  for (const para of cell.content) {
    td.appendChild(renderParagraph(para, deps));
  }

  return td;
}

// ─── Inline image block ───────────────────────────────────────────────────────

function renderInlineImageBlock(image: DocxInlineImage, deps: RenderDeps): HTMLElement {
  const wrapper = document.createElement('div');
  wrapper.className = 'docx-image';
  wrapper.style.display = 'block';

  const url = deps.imageUrl(image.part);
  if (url) {
    const img = document.createElement('img');
    img.src = url;
    img.alt = image.alt ?? '';
    img.style.width = `${image.widthPx}px`;
    img.style.height = `${image.heightPx}px`;
    img.style.maxWidth = '100%';
    img.style.display = 'block';
    wrapper.appendChild(img);
  }

  return wrapper;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function strokeToCssBorder(stroke: Stroke): string {
  const style = stroke.dash ? 'dashed' : 'solid';
  return `${stroke.width}px ${style} ${colorToCss(stroke.color)}`;
}

function cssClass(name: string): string {
  return name.toLowerCase().replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '');
}
