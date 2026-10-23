/**
 * DOM renderer: DocxPage model -> HTML/CSS. Consumes only the model (never
 * XML). Each page is a fixed-width sheet with margin padding; the viewer scales
 * pages to fit the container width and stacks them vertically.
 *
 * Three Word behaviours shape the CSS here:
 *  - The gap between two paragraphs is the *larger* of the first's space-after
 *    and the second's space-before — exactly CSS sibling-margin collapsing. But
 *    a cell's or text box's first paragraph keeps its space-before inside the
 *    box, so block containers are `flow-root` (no collapsing through them).
 *  - "Single" line spacing is the font's own line height (ascent + descent +
 *    gap), which differs per font — 1.22 for Calibri, 1.15 for Times New Roman.
 *    Line heights come from oxml/fontMetrics.ts, never a fixed multiplier.
 *  - Tabs jump to tab stops measured from the text column's edge. A stop's
 *    position depends on where the text before it ended, so tabs are laid out
 *    after the page is in the DOM — see {@link layoutTabs}.
 */
import { EDIT_ATTR } from '../../oxml/edit/attrs.js';
import type { DocxEditRenderContext } from '../edit/context.js';
import { ptToPx } from '../../oxml/units.js';
import { fontStack, lineHeightRatio, eastAsianLineHeightRatio, EAST_ASIAN_TEXT } from '../../oxml/fontMetrics.js';
import { presetPath, OPEN_PRESETS, EVENODD_PRESETS } from '../../oxml/drawingml/geometry/presets.js';
import { renderChart } from '../../oxml/drawingml/chart/render.js';
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
  DocxStroke,
  DocxTabStop,
  DocxComment,
  DocxBand,
  DocxChart,
  DocxGroup,
  MathNode,
  TextAlign,
} from '../model.js';

export interface RenderDeps {
  /** Resolve an embedded media part path to a displayable URL. */
  imageUrl(part: string): string | undefined;
  /** Present only when the document is being rendered for editing. */
  edit?: DocxEditRenderContext;
  /**
   * How tracked changes show: 'markup' (default) marks insertions and
   * deletions in the author's colour, 'final' shows the text as if every
   * change were accepted, 'original' as if every one were rejected.
   */
  revisions?: RevisionView;
  /** Highlight commented text (default true). The cards are drawn by {@link renderCommentRail}. */
  comments?: boolean;
  /**
   * The text column being laid out: its left edge from the page edge and its
   * width. Wrapped drawings are positioned against it.
   */
  column?: { leftPx: number; widthPx: number };
}

/** A drawing text wraps around, laid out as a CSS float. */
export interface FloatBox {
  side: 'left' | 'right';
  /** Offset from the paragraph's content edge on the float's own side, px. */
  inset: number;
  /** Gap kept on the text side, px. */
  gap: number;
  top: number;
  bottom: number;
  w: number;
  h: number;
  /** Text may not sit beside it at all (top-and-bottom wrapping). */
  fullWidth: boolean;
}

/**
 * The float a wrapped, paragraph-framed drawing becomes — undefined for one
 * text ignores (no wrapping, behind or in front of text) or one positioned
 * against the page, which is drawn absolutely instead.
 *
 * Text goes to the side Word would put it: the side `wrapText` names, else
 * the wider of the two. (Word can flow text on both sides of a drawing in the
 * middle of a column; CSS floats cannot, so such text takes the wider side.)
 */
export function anchorFloat(
  a: DocxAnchor,
  column: RenderDeps['column'],
  indentLeft: number,
  indentRight: number,
): FloatBox | undefined {
  if (isPageAnchored(a) || a.wrap === 'none') return undefined;
  const colW = column?.widthPx ?? 0;
  const colLeft = column?.leftPx ?? 0;
  const w = a.wPx;
  let x: number;
  if (a.hAlign === 'center') x = (colW - w) / 2;
  else if (a.hAlign === 'right' || a.hAlign === 'outside') x = colW - w;
  else if (a.hAlign === 'left' || a.hAlign === 'inside') x = a.relH === 'page' ? -colLeft : 0;
  else {
    const off = a.hOffsetPx ?? 0;
    if (a.relH === 'page' || a.relH === 'leftMargin' || a.relH === 'insideMargin') x = off - colLeft;
    else if (a.relH === 'rightMargin' || a.relH === 'outsideMargin') x = colW + off;
    else x = off;
  }
  const dist = a.distPx ?? { top: 0, right: 0, bottom: 0, left: 0 };
  const top = a.vOffsetPx ?? 0;
  const bottom = dist.bottom;
  if (a.wrap === 'topAndBottom') {
    return { side: 'left', inset: x - indentLeft, gap: 0, top: top - dist.top, bottom, w, h: a.hPx + dist.top, fullWidth: true };
  }
  const roomLeft = x;
  const roomRight = colW - x - w;
  const side =
    a.wrapText === 'left' ? 'right' : a.wrapText === 'right' ? 'left' : roomRight >= roomLeft ? 'left' : 'right';
  return side === 'left'
    ? { side, inset: x - indentLeft, gap: dist.right, top, bottom, w, h: a.hPx, fullWidth: false }
    : { side, inset: colW - x - w - indentRight, gap: dist.left, top, bottom, w, h: a.hPx, fullWidth: false };
}

export type RevisionView = 'markup' | 'final' | 'original';

/** Word-like per-author markup colours, assigned in the order authors are met. */
const AUTHOR_COLORS = ['#b4009e', '#2e7d32', '#c62828', '#1565c0', '#ef6c00', '#00838f', '#6a1b9a', '#558b2f'];
const authorColors = new Map<string, string>();

/** The markup colour for a reviewer, stable for the session. */
export function authorColor(author: string): string {
  let c = authorColors.get(author);
  if (!c) {
    c = AUTHOR_COLORS[authorColors.size % AUTHOR_COLORS.length]!;
    authorColors.set(author, c);
  }
  return c;
}

/** Whether a run's text shows at all under the chosen view of tracked changes. */
function revisionShown(run: DocxRun, view: RevisionView): boolean {
  if (!run.revision) return true;
  if (view === 'final') return run.revision.kind !== 'del';
  if (view === 'original') return run.revision.kind !== 'ins';
  return true;
}

/**
 * Per-page values that generated text (fields) draws from. Absent outside a
 * paginated page — fields then show the text Word cached when it saved.
 */
export interface FieldContext {
  /** The page number as displayed (PAGE). */
  pageNumber?: string;
  /** Total page count (NUMPAGES). */
  pageCount?: number;
  /** Pages in the current section (SECTIONPAGES). */
  sectionPageCount?: number;
  /** STYLEREF values, keyed by lowercased style name. */
  resolvedStyles?: Record<string, string>;
}

const px = (n: number) => `${n}px`;

const ALIGN: Record<TextAlign, string> = { l: 'left', ctr: 'center', r: 'right', just: 'justify' };

/** Word's built-in body font when a document states none: Times New Roman 10pt. */
const DEFAULT_FONT = 'Times New Roman';
const DEFAULT_SIZE_PT = 10;

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
  fontFamily: fontStack(DEFAULT_FONT),
  fontSize: px(ptToPx(DEFAULT_SIZE_PT)),
  lineHeight: String(lineHeightRatio(DEFAULT_FONT)),
} as const;

/**
 * Word sets text without pair kerning (unless a run asks for it with
 * <w:kern>) and without ligatures; browsers apply both by default, which
 * narrows lines enough to change where they break.
 */
export function applyWordTextDefaults(el: HTMLElement): void {
  el.style.fontKerning = 'none';
  el.style.fontVariantLigatures = 'none';
}

/**
 * Lay a container's block children out as Word does: adjacent paragraph
 * spacing collapses to the larger value, but never through the container.
 */
export function stackBlocks(el: HTMLElement): void {
  el.style.display = 'flow-root';
}

/**
 * Render one paginated page as a fixed-size sheet: content flows in the padding
 * box between the margins, with the header drawn in the top margin band and the
 * footer in the bottom margin band (matching Word).
 *
 * Call {@link layoutTabs} on the sheet once it is attached to the document.
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
  applyWordTextDefaults(sheet);
  s.overflow = 'hidden';
  // Its own stacking context, so a behind-text drawing (z-index -1) sits above
  // the white sheet but beneath the text rather than vanishing behind the page.
  s.isolation = 'isolate';
  stackBlocks(sheet);

  const contentW = page.size.wPx - page.margins.leftPx - page.margins.rightPx;
  deps = { ...deps, column: { leftPx: page.margins.leftPx, widthPx: contentW } };
  const field: FieldContext = {
    ...(page.pageNumber !== undefined ? { pageNumber: page.pageNumber } : {}),
    ...(page.pageCount !== undefined ? { pageCount: page.pageCount } : {}),
    ...(page.sectionPageCount !== undefined ? { sectionPageCount: page.sectionPageCount } : {}),
    ...(page.resolvedStyles ? { resolvedStyles: page.resolvedStyles } : {}),
  };
  if (field.pageNumber === undefined) field.pageNumber = String(page.index + 1);

  // Page-framed anchors (banners relativeFrom="page"/margins) painted first, in
  // page coordinates on the sheet, so behindDoc ones sit under the content and
  // full-bleed banners ignore the text margins.
  const pageAnchors: DocxAnchor[] = [];
  collectPageAnchors(page.elements, pageAnchors);
  if (page.header) collectPageAnchors(page.header, pageAnchors);
  if (page.footer) collectPageAnchors(page.footer, pageAnchors);
  for (const a of pageAnchors) sheet.appendChild(renderSheetAnchor(a, deps, page));

  if (page.header && page.header.length) {
    sheet.appendChild(marginBand(page.header, deps, contentW, page.margins.leftPx, { top: px(page.margins.headerPx) }, field));
  }
  if (page.footer && page.footer.length) {
    sheet.appendChild(marginBand(page.footer, deps, contentW, page.margins.leftPx, { bottom: px(page.margins.footerPx) }, field));
  }

  if (page.bands) {
    for (const band of page.bands) sheet.appendChild(renderBand(band, page, deps, field));
  } else {
    for (const block of page.elements) sheet.appendChild(renderBlock(block, deps, field));
  }

  // Footnotes sit at the foot of the text area, just above the bottom margin;
  // the paginator reserved their height.
  if (page.footnotes?.length) {
    const band = marginBand(page.footnotes, deps, contentW, page.margins.leftPx, { bottom: px(page.contentBottomPx ?? page.margins.bottomPx) }, field);
    sheet.appendChild(band);
  }
  return sheet;
}

/**
 * One band of a page: a single column flows as ordinary blocks; several sit
 * side by side in a box as tall as the paginator measured, with Word's rule
 * down the middle of each gap when the section asks for one.
 */
function renderBand(band: DocxBand, page: DocxPage, deps: RenderDeps, field: FieldContext): HTMLElement {
  const box = document.createElement('div');
  if (band.columns.length === 1) {
    stackBlocks(box);
    for (const block of band.blocks[0] ?? []) box.appendChild(renderBlock(block, deps, field));
    return box;
  }
  box.style.position = 'relative';
  box.style.height = px(band.heightPx);
  band.columns.forEach((c, i) => {
    const col = document.createElement('div');
    stackBlocks(col);
    col.style.position = 'absolute';
    col.style.top = '0';
    col.style.left = px(c.xPx);
    col.style.width = px(c.widthPx);
    const colDeps = { ...deps, column: { leftPx: page.margins.leftPx + c.xPx, widthPx: c.widthPx } };
    for (const block of band.blocks[i] ?? []) col.appendChild(renderBlock(block, colDeps, field));
    box.appendChild(col);
    const next = band.columns[i + 1];
    if (band.separator && next) {
      const rule = document.createElement('div');
      const x = (c.xPx + c.widthPx + next.xPx) / 2;
      rule.style.cssText = `position:absolute;top:0;bottom:0;left:${x}px;border-left:0.75px solid currentColor;`;
      box.appendChild(rule);
    }
  });
  return box;
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
  field: FieldContext,
): HTMLDivElement {
  const band = document.createElement('div');
  const s = band.style;
  s.position = 'absolute';
  s.left = px(leftPx);
  s.width = px(contentW);
  if (pos.top) s.top = pos.top;
  if (pos.bottom) s.bottom = pos.bottom;
  s.color = '#000';
  stackBlocks(band);
  // Headers, footers and footnotes are other parts: never typed into here.
  if (deps.edit) markUneditable(band);
  for (const block of blocks) band.appendChild(renderBlock(block, deps, field));
  return band;
}

/** Render a single block to a DOM element (also used by the paginator to measure). */
export function renderBlock(block: DocxBlock, deps: RenderDeps, field?: FieldContext): HTMLElement {
  switch (block.kind) {
    case 'paragraph':
      return renderParagraph(block, deps, field);
    case 'table':
      return renderTable(block, deps, field);
    case 'image': {
      const el = renderImage(block, deps);
      return el;
    }
  }
}

// ─── Paragraphs ──────────────────────────────────────────────────────────────

/** What tab layout needs to know about a paragraph, keyed by its element. */
interface TabInfo {
  tabs: DocxTabStop[];
  defaultTabPx: number;
  indentLeftPx: number;
  /** The first line starts left of the indent (a hanging indent), which is itself a stop. */
  hanging: boolean;
}

const TAB_INFO = new WeakMap<HTMLElement, TabInfo>();

/** Attribute on every tab span, content tab or list-marker tab, for {@link layoutTabs}. */
const TAB_STOP_ATTR = 'data-cbx-tabstop';

/** Line-spacing rule in force for a paragraph, which runs of other sizes/fonts must follow. */
interface LineRule {
  /** Multiple of each font's own single line (auto spacing). */
  pct?: number;
  /** Minimum line height in px (atLeast). */
  atLeastPx?: number;
}

function renderParagraph(p: DocxParagraph, deps: RenderDeps, field?: FieldContext): HTMLElement {
  const el = document.createElement('div');
  const s = el.style;
  s.position = 'relative';
  // Word keeps every space it is given and breaks an over-long word rather
  // than letting it run out of the column.
  s.whiteSpace = 'pre-wrap';
  s.overflowWrap = 'break-word';

  if (deps.edit?.editable(p)) {
    el.setAttribute(EDIT_ATTR.para, deps.edit.key(p));
    // Stated, not inherited: inside the editable page stack this changes
    // nothing, but it makes a text box's paragraphs editable inside the locked
    // drawing that holds them.
    el.setAttribute('contenteditable', 'true');
    el.spellcheck = false;
  } else if (deps.edit) {
    // Header, footer, note and comment text belongs to another part.
    markUneditable(el);
  }

  // A right-to-left paragraph lays its text out from the right. In such a
  // paragraph "left"/"right" alignment and indents mean start/end, so they
  // mirror.
  if (p.rtl) s.direction = 'rtl';
  const align = p.rtl && (p.align === 'l' || p.align === 'r') ? (p.align === 'l' ? 'r' : 'l') : p.align;
  if (align) s.textAlign = ALIGN[align];
  if (p.mathDisplay) s.textAlign = p.mathDisplay.align;
  const indentLeft = p.rtl ? (p.indentRightPx ?? 0) : (p.indentLeftPx ?? 0);
  const indentRight = p.rtl ? (p.indentLeftPx ?? 0) : (p.indentRightPx ?? 0);
  if (p.indentFirstLinePx !== undefined) s.textIndent = px(p.indentFirstLinePx);
  // Logical (block-axis) margins: in a vertical text box paragraph spacing runs
  // across the lines, not along them.
  if (p.spaceBeforePt !== undefined) s.marginBlockStart = px(ptToPx(p.spaceBeforePt));
  if (p.spaceAfterPt !== undefined) s.marginBlockEnd = px(ptToPx(p.spaceAfterPt));

  const family = p.baseFontFamily ?? DEFAULT_FONT;
  const sizePx = ptToPx(p.baseFontSizePt ?? DEFAULT_SIZE_PT);
  s.fontFamily = fontStack(family);
  s.fontSize = px(sizePx);
  if (p.baseBold) s.fontWeight = 'bold';
  if (p.baseItalic) s.fontStyle = 'italic';
  if (p.baseColorHex) s.color = `#${p.baseColorHex}`;
  if (p.shadingHex) s.background = `#${p.shadingHex}`;

  const rule: LineRule = {};
  if (p.lineSpacingPt !== undefined && p.lineSpacingRule === 'exact') {
    s.lineHeight = px(ptToPx(p.lineSpacingPt));
  } else if (p.lineSpacingPt !== undefined) {
    rule.atLeastPx = ptToPx(p.lineSpacingPt);
    s.lineHeight = px(Math.max(rule.atLeastPx, lineHeightRatio(family) * sizePx));
  } else {
    rule.pct = p.lineSpacingPct ?? 1;
    s.lineHeight = String(rule.pct * lineHeightRatio(family));
  }

  // Borders sit outside the text: the indent is where the text starts, so a
  // side border and its gap are drawn into the indent.
  const b = p.paraBorders ?? {};
  const gap = p.paraBorderSpacePx ?? {};
  const leftExtra = b.left ? (gap.left ?? 0) + b.left.width : 0;
  const rightExtra = b.right ? (gap.right ?? 0) + b.right.width : 0;
  if (indentLeft - leftExtra) s.marginLeft = px(indentLeft - leftExtra);
  if (indentRight - rightExtra) s.marginRight = px(indentRight - rightExtra);
  if (b.top) {
    s.borderTop = strokeCss(b.top);
    if (gap.top) s.paddingTop = px(gap.top);
  }
  if (b.bottom) {
    s.borderBottom = strokeCss(b.bottom);
    if (gap.bottom) s.paddingBottom = px(gap.bottom);
  }
  if (b.left) {
    s.borderLeft = strokeCss(b.left);
    if (gap.left) s.paddingLeft = px(gap.left);
  }
  if (b.right) {
    s.borderRight = strokeCss(b.right);
    if (gap.right) s.paddingRight = px(gap.right);
  }

  const hasTabs = p.listSuffix !== 'space' && p.listSuffix !== 'nothing' && !!p.listMarker;
  // Tab stops are measured from the left edge, so they are laid out for
  // left-to-right paragraphs only.
  if (!p.rtl && (hasTabs || p.runs.some((r) => r.tabBefore))) {
    TAB_INFO.set(el, {
      tabs: p.tabs ?? [],
      defaultTabPx: p.defaultTabPx ?? 48,
      indentLeftPx: indentLeft,
      hanging: (p.indentFirstLinePx ?? 0) < 0,
    });
  }

  // List marker: drawn at the first-line position (the hanging indent pulls it
  // into the gutter), followed by its suffix — by default a tab to the text
  // indent, so wrapped lines and the first line's text align.
  if (p.listMarker) {
    const marker = document.createElement('span');
    marker.textContent = p.listMarker + (p.listSuffix === 'space' ? ' ' : '');
    const f = p.listMarkerFormat ?? {};
    const ms = marker.style;
    if (f.font) ms.fontFamily = fontStack(f.font);
    if (f.sizePt !== undefined) ms.fontSize = px(ptToPx(f.sizePt));
    ms.fontWeight = f.bold ? 'bold' : 'normal';
    ms.fontStyle = f.italic ? 'italic' : 'normal';
    if (f.colorHex) ms.color = `#${f.colorHex}`;
    // The marker sizes its line by its own font — or the symbol font its glyph
    // came from — under the paragraph's spacing rule.
    const markerFont = f.metricsFont ?? f.font;
    if (rule.pct !== undefined && markerFont) ms.lineHeight = String(rule.pct * lineHeightRatio(markerFont));
    // Generated from numbering.xml, not content — it must not be typed into or
    // deleted, and reconciliation skips it (tab included, since it lives inside).
    markUneditable(marker);
    if (p.listSuffix !== 'space' && p.listSuffix !== 'nothing') marker.appendChild(tabSpan());
    el.appendChild(marker);
  }

  const view = deps.revisions ?? 'markup';
  const visible = p.runs.filter((r) => !r.hidden && revisionShown(r, view));
  // A changed paragraph gets Word's change bar in the left margin.
  if (view === 'markup' && p.runs.some((r) => r.revision || r.formatChange)) {
    const bar = document.createElement('span');
    markUneditable(bar);
    bar.style.cssText = `position:absolute;top:0;bottom:0;left:${-(indentLeft - leftExtra) - 24}px;width:0;border-left:1px solid #555;`;
    el.appendChild(bar);
  }
  if (visible.length === 0 && !p.listMarker) {
    // Preserve the height of an empty line.
    el.appendChild(document.createTextNode('\u200b'));
  }
  let fieldRun: string | undefined;
  for (const run of p.runs) {
    // A dynamic field's result may be split over several runs; the computed
    // value replaces the first and the rest render empty.
    const dynamic = run.fieldCode !== undefined && isDynamicField(run.fieldCode);
    const continuing = dynamic && fieldRun === run.fieldCode;
    fieldRun = dynamic ? run.fieldCode : undefined;
    appendRun(el, run, deps, { family, sizePx, rule, rtl: !!p.rtl, display: !!p.mathDisplay }, field, continuing);
  }
  if (visible.length === 0 && p.listMarker) el.appendChild(document.createTextNode('\u200b'));

  // Flow-framed anchors (relativeFrom column/paragraph) position within this
  // paragraph — vertically from its top, horizontally from the column edge.
  // One text wraps around becomes a float at the paragraph's start, so the
  // lines it overlaps (here and in following paragraphs) flow past it.
  // Page-framed anchors are hoisted to the sheet by renderPage.
  if (p.anchors) {
    const firstFlow = el.firstChild;
    for (const a of p.anchors) {
      if (isPageAnchored(a)) continue;
      const f = anchorFloat(a, deps.column, indentLeft - leftExtra, indentRight - rightExtra);
      if (f) {
        el.insertBefore(renderFloat(a, f, deps), firstFlow);
        continue;
      }
      el.appendChild(renderAnchor(a, deps, indentLeft - leftExtra, indentRight - rightExtra));
    }
  }
  return el;
}

/** Mark generated decoration as not content, as an attribute *and* property (jsdom reflects neither). */
function markUneditable(el: HTMLElement): void {
  el.setAttribute('contenteditable', 'false');
  el.contentEditable = 'false';
}

/** An empty inline box that {@link layoutTabs} later widens to reach the next tab stop. */
function tabSpan(ptab?: DocxRun['ptab']): HTMLSpanElement {
  const tab = document.createElement('span');
  tab.setAttribute(TAB_STOP_ATTR, ptab ?? '');
  tab.style.display = 'inline-block';
  tab.style.width = '0px';
  tab.style.textIndent = '0';
  return tab;
}

interface RunContext {
  family: string;
  sizePx: number;
  rule: LineRule;
  /** The paragraph is right-to-left. */
  rtl: boolean;
  /** The paragraph is a display equation. */
  display: boolean;
}

function appendRun(
  parent: HTMLElement,
  run: DocxRun,
  deps: RenderDeps,
  para: RunContext,
  field: FieldContext | undefined,
  continuingField: boolean,
): void {
  /** Stamp an element with the run it draws, so the caret can map through it. */
  const stamp = (el: Element, part?: 'br' | 'tab' | 'obj') => {
    if (!deps.edit) return;
    el.setAttribute(EDIT_ATTR.run, deps.edit.key(run));
    if (part) el.setAttribute(EDIT_ATTR.part, part);
  };
  if (run.breakBefore) {
    const br = document.createElement('br');
    if (run.pageBreakBefore) br.setAttribute('data-cbx-pagebreak', '');
    stamp(br, 'br');
    parent.appendChild(br);
  }
  if (run.tabBefore) {
    const tab = tabSpan(run.ptab);
    // A content tab: reconciliation turns it back into <w:tab/>.
    tab.setAttribute(EDIT_ATTR.tab, '');
    markUneditable(tab);
    stamp(tab, 'tab');
    parent.appendChild(tab);
  }

  // An equation, drawn by the browser from MathML.
  if (run.math) {
    const math = renderMath(run.math, para.display, run.sizePt !== undefined ? ptToPx(run.sizePt) : para.sizePx);
    stamp(math, 'obj');
    parent.appendChild(math);
    return;
  }

  // An inline drawing (picture / shape) occupying this run's slot, with the
  // room Word keeps around it for its outline and shadow.
  if (run.drawing) {
    const d = renderDrawing(run.drawing, deps);
    const e = run.drawing.effectPx;
    if (e) d.style.margin = `${e.top}px ${e.right}px ${e.bottom}px ${e.left}px`;
    if (deps.edit) {
      // An object in the text: the caret goes round it, never into it.
      markUneditable(d);
      stamp(d, 'obj');
    }
    parent.appendChild(d);
    return;
  }

  const view = deps.revisions ?? 'markup';
  if (!revisionShown(run, view)) return;

  // A comment's reference: an invisible anchor the comment rail lines its card up with.
  if (run.comment) {
    const anchor = document.createElement('span');
    anchor.setAttribute('data-cbx-comment', run.comment.id);
    markUneditable(anchor);
    parent.appendChild(anchor);
    return;
  }

  if (run.separator) {
    // The footnote separator: a short rule, two inches long, on the line.
    const rule = document.createElement('span');
    rule.style.cssText = 'display:inline-block;width:192px;vertical-align:middle;border-top:0.75px solid currentColor;';
    parent.appendChild(rule);
    return;
  }

  const displayText = continuingField ? '' : fieldText(run, field);
  if (!displayText && !run.hidden) return;

  const span = document.createElement('span');
  span.textContent = displayText;
  const deleted = run.revision?.kind === 'del';
  if (run.note || deleted) {
    // A note's reference mark is numbered by Word, not typed; deleted text is
    // gone from the document. Both are shown, never edited.
    markUneditable(span);
    stamp(span, run.note ? 'obj' : undefined);
  } else if (deps.edit) {
    stamp(span);
    // A field's text is regenerated by Word, so it is shown but locked.
    if (deps.edit.isField(run)) span.contentEditable = 'false';
  }
  const s = span.style;
  // Hidden text stays in the DOM (so editing keeps it) but takes no space.
  if (run.hidden) s.display = 'none';
  // Stated either way: a run may switch off what its paragraph's style turns on.
  s.fontWeight = run.bold ? 'bold' : 'normal';
  s.fontStyle = run.italic ? 'italic' : 'normal';
  const deco: string[] = [];
  if (run.underline) deco.push('underline');
  if (run.strike) deco.push('line-through');
  if (deco.length) s.textDecoration = deco.join(' ');
  const sizePx = run.sizePt !== undefined ? ptToPx(run.sizePt) : para.sizePx;
  if (run.sizePt !== undefined) s.fontSize = px(sizePx);
  if (run.colorHex) s.color = `#${run.colorHex}`;
  if (run.highlightHex) s.background = `#${run.highlightHex}`;
  if (run.font || run.eastAsiaFont) s.fontFamily = fontStack(run.font ?? para.family, run.eastAsiaFont);
  // A right-to-left run inside a left-to-right paragraph is its own embedding.
  if (run.rtl && !para.rtl) {
    span.dir = 'rtl';
    s.unicodeBidi = 'embed';
  }
  if (run.charSpacingPt) s.letterSpacing = px(ptToPx(run.charSpacingPt));
  if (run.kern) s.fontKerning = 'normal';
  if (run.caps === 'all') s.textTransform = 'uppercase';
  else if (run.caps === 'small') s.fontVariant = 'small-caps';
  if (run.vertAlign === 'super') {
    s.verticalAlign = 'super';
    s.fontSize = px(sizePx * 0.65);
  } else if (run.vertAlign === 'sub') {
    s.verticalAlign = 'sub';
    s.fontSize = px(sizePx * 0.65);
  }

  // A run in another font or size gets that font's own line height under the
  // paragraph's rule, as Word sizes each line to its tallest run — except a
  // run of nothing but spaces, which Word never lets grow a line. East Asian
  // characters are set in the run's East Asian face, whose (much taller) line
  // then counts too.
  const family = run.font ?? para.family;
  if (EAST_ASIAN_TEXT.test(displayText) && para.rule.pct !== undefined) {
    const ratio = Math.max(lineHeightRatio(family), eastAsianLineHeightRatio(run.eastAsiaFont));
    s.lineHeight = String(para.rule.pct * ratio);
  } else if (/^[ \u00a0\t]+$/.test(displayText) && (sizePx > para.sizePx || family !== para.family)) {
    s.lineHeight = '0';
  } else if (family !== para.family || sizePx !== para.sizePx) {
    if (para.rule.pct !== undefined && family !== para.family) {
      s.lineHeight = String(para.rule.pct * lineHeightRatio(family));
    } else if (para.rule.atLeastPx !== undefined) {
      s.lineHeight = px(Math.max(para.rule.atLeastPx, lineHeightRatio(family) * sizePx));
    }
  }

  if (run.revision && view === 'markup') {
    // Word's markup: insertions underlined, deletions struck through, both in
    // the reviewer's colour.
    s.color = authorColor(run.revision.author);
    s.textDecoration = `${deco.join(' ')} ${deleted ? 'line-through' : 'underline'}`.trim();
  }
  if (run.commentIds?.length && deps.comments !== false) {
    s.background = 'rgba(255, 200, 40, 0.28)';
    span.setAttribute('data-cbx-commented', run.commentIds.join(' '));
  }

  if (run.hyperlink) {
    const a = document.createElement('a');
    a.href = run.hyperlink;
    a.target = '_blank';
    a.rel = 'noopener noreferrer';
    a.style.color = run.colorHex ? `#${run.colorHex}` : 'inherit';
    a.style.textDecoration = 'inherit';
    a.appendChild(span);
    parent.appendChild(a);
  } else {
    parent.appendChild(span);
  }
}

/** Fields whose result depends on the page being drawn. */
function isDynamicField(code: string): boolean {
  return /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.test(code);
}

/** The text a run shows: its own, or a field's value computed for this page. */
function fieldText(run: DocxRun, field: FieldContext | undefined): string {
  if (!run.fieldCode || !field) return run.text;
  const code = run.fieldCode.trim();
  const name = /^\S+/.exec(code)?.[0]?.toUpperCase();
  const format = /\\\*\s*(\w+)/.exec(code)?.[1];
  switch (name) {
    case 'PAGE':
      return field.pageNumber !== undefined ? formatField(field.pageNumber, format) : run.text;
    case 'NUMPAGES':
      return field.pageCount !== undefined ? formatField(String(field.pageCount), format) : run.text;
    case 'SECTIONPAGES':
      return field.sectionPageCount !== undefined ? formatField(String(field.sectionPageCount), format) : run.text;
    case 'STYLEREF': {
      const match = /^\s*styleref\s+(?:"([^"]+)"|'([^']+)'|([^\s]+))/i.exec(code);
      const styleName = (match?.[1] || match?.[2] || match?.[3] || '').toLowerCase();
      return field.resolvedStyles?.[styleName] ?? run.text;
    }
    default:
      return run.text;
  }
}

/** Apply a field's `\* format` switch to a number already formatted for its section. */
function formatField(value: string, format: string | undefined): string {
  const n = Number(value);
  if (!format || !Number.isFinite(n)) return value;
  switch (format) {
    case 'roman':
      return toRoman(n).toLowerCase();
    case 'Roman':
    case 'ROMAN':
      return toRoman(n);
    case 'alphabetic':
      return letters(n).toLowerCase();
    case 'ALPHABETIC':
      return letters(n);
    default:
      return value;
  }
}

function toRoman(n: number): string {
  const table: [number, string][] = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  let x = n;
  let out = '';
  for (const [v, sym] of table) {
    while (x >= v) {
      out += sym;
      x -= v;
    }
  }
  return out;
}

function letters(n: number): string {
  if (n <= 0) return '';
  return String.fromCharCode(65 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);
}

// ─── Comments ────────────────────────────────────────────────────────────────

/** Width of the rail of comment cards beside a page, px. */
export const COMMENT_RAIL_PX = 240;

/** Every comment referenced from a page, in reading order. */
export function pageComments(page: DocxPage): DocxComment[] {
  const out: DocxComment[] = [];
  const visit = (blocks: DocxBlock[]) => {
    for (const b of blocks) {
      if (b.kind === 'paragraph') for (const r of b.runs) if (r.comment) out.push(r.comment);
      if (b.kind === 'table') for (const row of b.rows) for (const c of row) if (c) visit(c.content);
    }
  };
  visit(page.elements);
  return out;
}

/**
 * The comment cards for a page, as Word's markup pane draws them: a column
 * beside the page, each card level with its comment's anchor. Call
 * {@link layoutComments} once the page and rail are in the document.
 */
export function renderCommentRail(page: DocxPage, deps: RenderDeps): HTMLDivElement | undefined {
  const comments = pageComments(page);
  if (!comments.length) return undefined;
  const rail = document.createElement('div');
  rail.setAttribute('data-cbx-comment-rail', '');
  if (deps.edit) markUneditable(rail);
  rail.style.cssText = `position:relative;flex:none;width:${COMMENT_RAIL_PX}px;margin-left:12px;`;
  rail.style.height = px(page.size.hPx);
  for (const c of comments) {
    const card = document.createElement('div');
    card.setAttribute('data-cbx-comment-card', c.id);
    const color = authorColor(c.author);
    card.style.cssText =
      `position:absolute;left:0;right:0;box-sizing:border-box;padding:6px 8px;background:#fff;` +
      `border-left:3px solid ${color};border-radius:3px;box-shadow:0 1px 3px rgba(0,0,0,0.3);` +
      'font:12px/1.35 "Segoe UI", system-ui, sans-serif;color:#222;';
    const head = document.createElement('div');
    head.style.cssText = `font-weight:600;color:${color};margin-bottom:2px;`;
    head.textContent = c.author || c.initials || 'Comment';
    card.appendChild(head);
    const body = document.createElement('div');
    body.style.fontSize = '11px';
    for (const b of c.blocks) {
      const el = renderBlock(b, { ...deps, comments: false });
      // Card text reads at the card's size, not the document's.
      el.style.fontSize = '';
      el.style.lineHeight = '';
      el.style.margin = '0';
      body.appendChild(el);
    }
    card.appendChild(body);
    rail.appendChild(card);
  }
  return rail;
}

/** Line each comment card up with its anchor on the page, pushing overlapping cards down. */
export function layoutComments(frame: HTMLElement): void {
  const rail = frame.querySelector<HTMLElement>('[data-cbx-comment-rail]');
  if (!rail) return;
  const railRect = rail.getBoundingClientRect();
  const scale = rail.offsetHeight > 0 && railRect.height > 0 ? railRect.height / rail.offsetHeight : 1;
  let next = 0;
  for (const card of Array.from(rail.querySelectorAll<HTMLElement>('[data-cbx-comment-card]'))) {
    const id = card.getAttribute('data-cbx-comment-card');
    const anchor =
      frame.querySelector(`[data-cbx-commented~="${id}"]`) ?? frame.querySelector(`[data-cbx-comment="${id}"]`);
    const want = anchor ? (anchor.getBoundingClientRect().top - railRect.top) / scale : next;
    const top = Math.max(want, next);
    card.style.top = px(top);
    next = top + card.offsetHeight + 6;
  }
}

// ─── Tab layout ──────────────────────────────────────────────────────────────

/**
 * Size every tab under `root` so the text after it starts (or, for a right /
 * centre / decimal stop, ends or centres) at the tab stop, as Word does.
 *
 * Must run with `root` attached to a document: a tab's width depends on where
 * the preceding text ended, which only layout knows. Tabs are processed in
 * order because each one moves everything after it.
 */
export function layoutTabs(root: HTMLElement): void {
  const tabs = root.querySelectorAll<HTMLElement>(`[${TAB_STOP_ATTR}]`);
  if (tabs.length === 0) return;
  for (const tab of Array.from(tabs)) {
    const para = owningParagraph(tab);
    if (!para) continue;
    const info = TAB_INFO.get(para)!;
    layoutTab(tab, para, info, Array.from(para.querySelectorAll<HTMLElement>(`[${TAB_STOP_ATTR}]`)));
  }
}

function owningParagraph(tab: HTMLElement): HTMLElement | undefined {
  for (let n = tab.parentElement; n; n = n.parentElement) {
    if (TAB_INFO.has(n)) return n;
  }
  return undefined;
}

function layoutTab(tab: HTMLElement, para: HTMLElement, info: TabInfo, siblings: HTMLElement[]): void {
  tab.style.width = '0px';
  tab.style.background = '';
  tab.style.borderBottom = '';
  tab.style.height = '';
  tab.style.verticalAlign = '';

  const cs = getComputedStyle(para);
  const padLeft = parseFloat(cs.paddingLeft) || 0;
  const textW = para.clientWidth - padLeft - (parseFloat(cs.paddingRight) || 0);
  // x: where the tab starts, relative to the paragraph's text indent.
  const x = offsetWithin(tab, para) - padLeft;
  const toIndent = (colPos: number) => colPos - info.indentLeftPx;

  const ptab = tab.getAttribute(TAB_STOP_ATTR);
  let stop: { pos: number; align: DocxTabStop['align']; leader?: DocxTabStop['leader'] } | undefined;
  if (ptab) {
    // An absolute-position tab aligns to the text column itself.
    const pos = ptab === 'right' ? textW : ptab === 'center' ? textW / 2 : 0;
    stop = { pos, align: ptab === 'right' ? 'right' : ptab === 'center' ? 'center' : 'left' };
  } else {
    const custom = info.tabs.filter((t) => t.align !== 'bar').map((t) => ({ ...t, pos: toIndent(t.posPx) }));
    const next = custom.find((t) => t.pos > x + 0.5);
    if (next) stop = { pos: next.pos, align: next.align, ...(next.leader ? { leader: next.leader } : {}) };
    // A hanging indent is itself a stop for the first line's tab.
    if (info.hanging && x < -0.5 && (!stop || stop.pos > 0)) stop = { pos: 0, align: 'left' };
    if (!stop) {
      // Default stops, every defaultTabPx from the column edge — but none to
      // the left of the last custom stop.
      const last = custom.length ? custom[custom.length - 1]!.pos : -Infinity;
      const from = Math.max(x + 0.5, last);
      const step = info.defaultTabPx;
      const k = Math.floor((from + info.indentLeftPx) / step) + 1;
      stop = { pos: toIndent(k * step), align: 'left' };
    }
  }

  let width = stop.pos - x;
  if (stop.align === 'right' || stop.align === 'center' || stop.align === 'decimal') {
    const next = siblings[siblings.indexOf(tab) + 1];
    const seg = segmentWidth(tab, next, para, stop.align === 'decimal');
    width -= stop.align === 'center' ? seg / 2 : seg;
  }
  // Never wider than the room left on the line, or it would wrap onto its own.
  width = Math.max(0, Math.min(width, textW - x));
  tab.style.width = px(width);
  if (stop.leader && width > 0) applyLeader(tab, stop.leader);
}

/** Horizontal offset of `el` within `ancestor`'s padding box, in unscaled px. */
function offsetWithin(el: HTMLElement, ancestor: HTMLElement): number {
  let x = 0;
  let n: HTMLElement | null = el;
  while (n && n !== ancestor) {
    x += n.offsetLeft;
    const parent: Element | null = n.offsetParent;
    if (parent === ancestor || !(parent instanceof HTMLElement)) break;
    n = parent;
  }
  return x;
}

/**
 * Width of the content between a tab and the next one (or the paragraph end):
 * what a right/centre stop aligns. For a decimal stop, only up to the first
 * decimal point.
 */
function segmentWidth(tab: HTMLElement, next: HTMLElement | undefined, para: HTMLElement, decimal: boolean): number {
  const doc = para.ownerDocument;
  const rect = para.getBoundingClientRect();
  const scale = para.offsetWidth > 0 && rect.width > 0 ? rect.width / para.offsetWidth : 1;
  const walker = doc.createTreeWalker(para, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  walker.currentNode = tab;
  let width = 0;
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n === next) break;
    if (n instanceof HTMLElement) {
      // Inline pictures and shapes count; floating ones do not.
      if ((n.tagName === 'IMG' || n.dataset.cbxDrawing !== undefined) && getComputedStyle(n).position !== 'absolute') {
        width += n.getBoundingClientRect().width;
      }
      continue;
    }
    const text = n as Text;
    let end = text.data.length;
    if (decimal) {
      const dot = text.data.indexOf('.');
      if (dot >= 0) end = dot;
    }
    const range = doc.createRange();
    range.setStart(text, 0);
    range.setEnd(text, end);
    width += range.getBoundingClientRect().width;
    if (decimal && end < text.data.length) break;
  }
  return width / scale;
}

/** Draw a tab leader across the tab's width, sitting on the text baseline. */
function applyLeader(tab: HTMLElement, leader: NonNullable<DocxTabStop['leader']>): void {
  const s = tab.style;
  switch (leader) {
    case 'dot':
    case 'middleDot':
      s.height = '3px';
      s.backgroundImage = 'radial-gradient(circle, currentColor 0.9px, transparent 1.2px)';
      s.backgroundSize = '0.3em 3px';
      s.backgroundRepeat = 'repeat-x';
      s.backgroundPosition = 'left bottom';
      if (leader === 'middleDot') s.verticalAlign = '0.3em';
      break;
    case 'hyphen':
      s.height = '1px';
      s.backgroundImage = 'repeating-linear-gradient(to right, currentColor 0 0.25em, transparent 0.25em 0.45em)';
      s.verticalAlign = '0.3em';
      break;
    case 'underscore':
      s.borderBottom = '1px solid currentColor';
      break;
    case 'heavy':
      s.borderBottom = '2px solid currentColor';
      break;
  }
}

// ─── Tables ──────────────────────────────────────────────────────────────────

function renderTable(table: DocxTable, deps: RenderDeps, field?: FieldContext): HTMLElement {
  const t = document.createElement('table');
  t.style.borderCollapse = 'collapse';
  t.style.tableLayout = 'fixed';
  if (table.widthPx) t.style.width = px(table.widthPx);
  if (table.align === 'center') {
    t.style.marginLeft = 'auto';
    t.style.marginRight = 'auto';
  } else if (table.align === 'right') {
    t.style.marginLeft = 'auto';
  } else if (table.indentPx) {
    t.style.marginLeft = px(table.indentPx);
  }

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
  table.rows.forEach((row, r) => {
    const tr = document.createElement('tr');
    const height = table.rowHeights?.[r];
    // A table row's CSS height is a minimum, which is Word's "at least"; an
    // exact row also clips its cells' content (see renderCell). Word's row
    // height excludes the row's borders, while a collapsed CSS row contains
    // half of each — so the borders are added back.
    if (height) tr.style.height = px(height.px + rowBorderShare(row));
    for (const cell of row) {
      if (cell === null) continue; // covered by a span
      tr.appendChild(renderCell(cell, deps, field, height?.rule === 'exact' ? height.px : undefined));
    }
    tbody.appendChild(tr);
  });
  t.appendChild(tbody);
  return t;
}

/** Half the widest top border plus half the widest bottom border of a row's cells. */
function rowBorderShare(row: (DocxTableCell | null)[]): number {
  let top = 0;
  let bottom = 0;
  for (const cell of row) {
    if (!cell?.borders) continue;
    top = Math.max(top, cell.borders.t?.width ?? 0);
    bottom = Math.max(bottom, cell.borders.b?.width ?? 0);
  }
  return (top + bottom) / 2;
}

function renderCell(cell: DocxTableCell, deps: RenderDeps, field?: FieldContext, exactPx?: number): HTMLElement {
  const td = document.createElement('td');
  if (cell.colSpan > 1) td.colSpan = cell.colSpan;
  if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
  const s = td.style;
  if (cell.empty) {
    // A grid column the row skips: nothing drawn, not even a border.
    s.padding = '0';
    s.border = 'none';
    return td;
  }
  // Positioning context for the cell's floating anchors (hoisted off their
  // paragraphs so a centered cell doesn't shift them).
  s.position = 'relative';
  s.verticalAlign = cell.vAlign === 'center' ? 'middle' : (cell.vAlign ?? 'top');
  if (cell.fillHex) {
    s.background = `#${cell.fillHex}`;
    // Word's "automatic" font color flips to white on a dark fill for contrast.
    // Set it as the cell default so explicit run/paragraph colors still win.
    if (isDarkFill(cell.fillHex)) s.color = '#fff';
  }
  const p = cell.cellPaddingPx ?? { top: 0, right: 7.2, bottom: 0, left: 7.2 };
  s.padding = `${p.top}px ${p.right}px ${p.bottom}px ${p.left}px`;
  if (cell.borders) {
    if (cell.borders.l) s.borderLeft = strokeCss(cell.borders.l);
    if (cell.borders.t) s.borderTop = strokeCss(cell.borders.t);
    if (cell.borders.r) s.borderRight = strokeCss(cell.borders.r);
    if (cell.borders.b) s.borderBottom = strokeCss(cell.borders.b);
  }

  const inner = document.createElement('div');
  stackBlocks(inner);
  if (exactPx !== undefined) {
    inner.style.maxHeight = px(Math.max(0, exactPx - p.top - p.bottom));
    inner.style.overflow = 'hidden';
  }
  for (const block of cell.content) inner.appendChild(renderBlock(block, deps, field));
  if (!cell.content.length) inner.appendChild(document.createTextNode('\u200b'));
  td.appendChild(inner);
  return td;
}

// ─── Drawings ────────────────────────────────────────────────────────────────

/**
 * Render a floating (anchored) drawing absolutely inside its containing,
 * position:relative paragraph/cell. `left`/`top` resolve the anchor's
 * relativeFrom + offset/align against the paragraph's content box using CSS
 * calc — so an image anchored in a right-hand cell lands on the right without
 * the renderer needing to know the box width.
 */
function renderAnchor(anchor: DocxAnchor, deps: RenderDeps, indentLeft = 0, indentRight = 0): HTMLElement {
  const el = renderDrawing(anchor.drawing, deps);
  if (deps.edit) markUneditable(el);
  const s = el.style;
  s.position = 'absolute';
  s.margin = '0';
  s.width = px(anchor.wPx);
  s.height = px(anchor.hPx);
  s.maxWidth = 'none';
  s.left = anchorLeft(anchor, indentLeft, indentRight);
  s.top = anchorTop(anchor);
  s.zIndex = anchorZ(anchor);
  return el;
}

/**
 * A wrapped drawing as a CSS float, which the paragraph's lines flow around.
 *
 * A drawing offset down into its paragraph must leave the lines above it full
 * width. A float's top margin (or a spacer above it) still narrows them, so
 * the float instead spans from the paragraph's top and `shape-outside` cuts
 * the offset off the area text avoids; the drawing sits at the float's foot.
 */
function renderFloat(a: DocxAnchor, f: FloatBox, deps: RenderDeps): HTMLElement {
  const wrap = document.createElement('div');
  markUneditable(wrap);
  const ws = wrap.style;
  const top = Math.max(0, f.top);
  ws.float = f.side;
  ws.position = 'relative';
  ws.width = px(a.wPx);
  ws.height = px(top + a.hPx);
  ws.marginBottom = px(f.bottom);
  if (f.top < 0) ws.marginTop = px(f.top);
  // (Word lets a drawing's top edge graze the line above by a hair without
  // wrapping it, hence the pixel of tolerance.)
  if (top > 0) ws.shapeOutside = `inset(${top + 1}px 0 0 0)`;
  if (f.side === 'left') {
    ws.marginLeft = px(f.inset);
    // Top-and-bottom: fill the rest of the line so no text fits beside it.
    ws.marginRight = f.fullWidth ? `calc(100% - ${f.inset + a.wPx}px)` : px(f.gap);
  } else {
    ws.marginRight = px(f.inset);
    ws.marginLeft = px(f.gap);
  }
  ws.textIndent = '0';

  const el = renderDrawing(a.drawing, deps);
  const s = el.style;
  s.position = 'absolute';
  s.left = '0';
  s.top = px(top);
  s.width = px(a.wPx);
  s.height = px(a.hPx);
  s.maxWidth = 'none';
  s.margin = '0';
  s.zIndex = anchorZ(a);
  wrap.appendChild(el);
  return wrap;
}

/**
 * Render a page-framed anchor (relativeFrom page/margins) in absolute page
 * coordinates on the sheet, so full-bleed banners ignore the text margins and
 * relativeFrom="page" offsets are measured from the page edge.
 */
function renderSheetAnchor(anchor: DocxAnchor, deps: RenderDeps, page: DocxPage): HTMLElement {
  const el = renderDrawing(anchor.drawing, deps);
  if (deps.edit) markUneditable(el);
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
    else if (anchor.relH === 'leftMargin' || anchor.relH === 'insideMargin') left = off;
    else left = margins.leftPx + off; // margin / column / character
  }

  // Vertical.
  let top: number;
  if (anchor.vAlign === 'center') {
    top = anchor.relV === 'page' ? (size.hPx - anchor.hPx) / 2 : margins.topPx + (size.hPx - margins.topPx - margins.bottomPx - anchor.hPx) / 2;
  } else if (anchor.vAlign === 'bottom') {
    top = anchor.relV === 'page' || anchor.relV === 'bottomMargin' ? size.hPx - anchor.hPx : size.hPx - margins.bottomPx - anchor.hPx;
  } else if (anchor.vAlign === 'top') {
    top = anchor.relV === 'page' || anchor.relV === 'topMargin' ? 0 : margins.topPx;
  } else {
    const off = anchor.vOffsetPx ?? 0;
    if (anchor.relV === 'page' || anchor.relV === 'topMargin') top = off;
    else if (anchor.relV === 'bottomMargin' || anchor.relV === 'outsideMargin') top = size.hPx - margins.bottomPx + off;
    else top = margins.topPx + off; // margin / text
  }

  s.left = px(left);
  s.top = px(top);
  s.zIndex = anchorZ(anchor);
  return el;
}

/**
 * Stacking order: Word layers floating drawings by relativeHeight, behind-text
 * ones below the text (negative) and the rest above it — all in the sheet's
 * one stacking context, so drawings from different paragraphs interleave.
 */
function anchorZ(a: DocxAnchor): string {
  const order = Math.min(Math.max(a.zOrder ?? 0, 0), 0x7ffffffe);
  return String(a.behindDoc ? order - 0x7fffffff : order + 1);
}

/**
 * CSS `left` for an anchor inside its paragraph. Positions are stated against
 * the text column, which extends `indentLeft`/`indentRight` past the
 * paragraph's own box.
 */
function anchorLeft(a: DocxAnchor, indentLeft: number, indentRight: number): string {
  const w = a.wPx;
  const column = `(100% + ${indentLeft + indentRight}px)`;
  if (a.hAlign === 'center') return `calc((${column} - ${w}px) / 2 - ${indentLeft}px)`;
  if (a.hAlign === 'right' || a.hAlign === 'outside') return `calc(${column} - ${w}px - ${indentLeft}px)`;
  if (a.hAlign === 'left' || a.hAlign === 'inside') return px(-indentLeft);
  const off = a.hOffsetPx ?? 0;
  // rightMargin/outsideMargin measure from the right column edge.
  if (a.relH === 'rightMargin' || a.relH === 'outsideMargin') return `calc(${column} + ${off - indentLeft}px)`;
  return px(off - indentLeft);
}

/** CSS `top` for an anchor, relative to the containing paragraph's top. */
function anchorTop(a: DocxAnchor): string {
  if (a.vAlign === 'center') return `calc((100% - ${a.hPx}px) / 2)`;
  if (a.vAlign === 'bottom') return `calc(100% - ${a.hPx}px)`;
  if (a.vAlign === 'top') return '0px';
  return px(a.vOffsetPx ?? 0);
}

/** A drawing (picture, shape, chart or group) as an inline-block element. */
function renderDrawing(drawing: DocxDrawing, deps: RenderDeps): HTMLElement {
  let el: HTMLElement;
  switch (drawing.kind) {
    case 'shape':
      el = renderShape(drawing, deps);
      break;
    case 'image':
      el = renderImage(drawing, deps);
      break;
    case 'chart':
      el = renderChartBox(drawing);
      break;
    case 'group':
      el = renderGroup(drawing, deps);
      break;
  }
  el.dataset.cbxDrawing = '';
  return el;
}

/** A chart, drawn by the shared DrawingML chart renderer into its box. */
function renderChartBox(c: DocxChart): HTMLElement {
  const el = document.createElement('div');
  el.style.cssText = `display:inline-block;position:relative;vertical-align:baseline;width:${c.widthPx}px;height:${c.heightPx}px;text-indent:0;`;
  el.appendChild(renderChart(c.chart, c.widthPx, c.heightPx));
  return el;
}

/** A group (or SmartArt diagram): each drawing placed and sized in the group's box. */
function renderGroup(g: DocxGroup, deps: RenderDeps): HTMLElement {
  const el = document.createElement('div');
  el.style.cssText = `display:inline-block;position:relative;vertical-align:baseline;width:${g.widthPx}px;height:${g.heightPx}px;text-indent:0;white-space:normal;`;
  for (const item of g.items) {
    const child = renderDrawing(item.drawing, deps);
    const s = child.style;
    s.position = 'absolute';
    s.left = px(item.xPx);
    s.top = px(item.yPx);
    s.width = px(item.wPx);
    s.height = px(item.hPx);
    s.margin = '0';
    s.maxWidth = 'none';
    if (item.rot) s.transform = `rotate(${item.rot}deg)`;
    el.appendChild(child);
  }
  return el;
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
  // An inline picture sits on the text baseline, as in Word.
  el.style.verticalAlign = 'baseline';
  if (img.duotone) el.style.filter = `url(#${duotoneFilter(el.ownerDocument, img.duotone)})`;
  return el;
}

const SVG_NS = 'http://www.w3.org/2000/svg';

/**
 * Register (once per document and colour pair) an SVG filter recolouring an
 * image as a duotone — luminance mapped linearly from the dark colour to the
 * light one — and return its id. CSS `filter: url()` only reaches filters in
 * the same document, so the definition lives in a hidden <svg> in the body.
 */
function duotoneFilter(doc: Document, [dark, light]: [string, string]): string {
  const id = `cbx-duotone-${dark}-${light}`.toLowerCase();
  if (doc.getElementById(id)) return id;
  const make = (name: string, attrs: Record<string, string>) => {
    const node = doc.createElementNS(SVG_NS, name);
    for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, v);
    return node;
  };
  const ch = (hex: string, i: number) => parseInt(hex.slice(i * 2, i * 2 + 2), 16) / 255;
  const table = (i: number) => `${ch(dark, i).toFixed(3)} ${ch(light, i).toFixed(3)}`;
  const svg = make('svg', { width: '0', height: '0', 'aria-hidden': 'true' });
  svg.style.position = 'absolute';
  const filter = make('filter', { id, 'color-interpolation-filters': 'sRGB' });
  filter.appendChild(
    make('feColorMatrix', {
      type: 'matrix',
      values: '0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0 0 0 1 0',
    }),
  );
  const transfer = make('feComponentTransfer', {});
  ['feFuncR', 'feFuncG', 'feFuncB'].forEach((f, i) => transfer.appendChild(make(f, { type: 'table', tableValues: table(i) })));
  filter.appendChild(transfer);
  svg.appendChild(filter);
  (doc.body ?? doc.documentElement).appendChild(svg);
  return id;
}

/** A DrawingML shape (filled/outlined box, optionally a text box). */
function renderShape(shape: DocxShape, deps: RenderDeps): HTMLElement {
  const el = document.createElement('div');
  const s = el.style;
  s.display = 'inline-block';
  s.verticalAlign = 'baseline';
  s.boxSizing = 'border-box';
  s.whiteSpace = 'normal';
  s.textIndent = '0';
  if (shape.widthPx) s.width = px(shape.widthPx);
  if (shape.heightPx) s.height = px(shape.heightPx);
  s.maxWidth = '100%';

  if (shape.rot) s.transform = `rotate(${shape.rot}deg)`;
  if (shape.preset || shape.paths) return renderPathShape(shape, deps, el);

  if (shape.geom === 'line') {
    // A connector/line: render as a rule using the outline color.
    s.borderTop = `${shape.lineWidthPx ? Math.max(1, shape.lineWidthPx) : 1}px solid ${shape.lineHex ? `#${shape.lineHex}` : '#000'}`;
    s.height = px(shape.heightPx || 0);
    return el;
  }

  if (shape.fillHex) s.background = `#${shape.fillHex}`;
  const fillUrl = shape.fillImagePart ? deps.imageUrl(shape.fillImagePart) : undefined;
  if (fillUrl) {
    s.backgroundImage = `url("${fillUrl}")`;
    s.backgroundSize = '100% 100%';
    s.backgroundRepeat = 'no-repeat';
  }
  // Text colour: the shape style's, else white on a dark fill (Word's automatic
  // colour); explicit run/paragraph colours still win.
  if (shape.textColorHex) s.color = `#${shape.textColorHex}`;
  else if (shape.fillHex && isDarkFill(shape.fillHex)) s.color = '#fff';
  if (shape.lineHex) {
    // An outline straddles the shape's edge: half inside as a border, half
    // outside as a spread shadow, which follows the corner radius.
    const half = Math.max(0.5, (shape.lineWidthPx ?? 1) / 2);
    s.border = `${half}px solid #${shape.lineHex}`;
    s.boxShadow = `0 0 0 ${half}px #${shape.lineHex}`;
  }
  if (shape.geom === 'roundRect') s.borderRadius = px(Math.min(shape.widthPx, shape.heightPx) * 0.16);
  else if (shape.geom === 'ellipse') s.borderRadius = '50%';

  // Vertically anchor the text box content.
  s.display = 'inline-flex';
  s.flexDirection = 'column';
  s.justifyContent = shape.vAnchor === 'bottom' ? 'flex-end' : shape.vAnchor === 'ctr' ? 'center' : 'flex-start';
  const inner = document.createElement('div');
  stackBlocks(inner);
  inner.style.width = '100%';
  inner.style.boxSizing = 'border-box';
  if (shape.vert) {
    // Vertical text: lines run top to bottom, stacking from the right; vert
    // turns every character, eaVert keeps East Asian ones upright, and
    // vert270 is the same turned the other way up (bottom to top).
    s.flexDirection = 'row';
    s.justifyContent = shape.vert === 'vert270' ? 'flex-start' : 'flex-end';
    inner.style.width = '';
    inner.style.height = '100%';
    inner.style.writingMode = 'vertical-rl';
    if (shape.vert !== 'eaVert') inner.style.textOrientation = 'sideways';
    if (shape.vert === 'vert270') inner.style.transform = 'rotate(180deg)';
  }
  const ins = shape.insetsPx ?? { top: 4.8, right: 9.6, bottom: 4.8, left: 9.6 };
  inner.style.padding = `${ins.top}px ${ins.right}px ${ins.bottom}px ${ins.left}px`;
  for (const block of shape.content) inner.appendChild(renderBlock(block, deps));
  el.appendChild(inner);
  return el;
}

/**
 * A shape whose outline CSS cannot draw (an arrow, a chevron, custom
 * geometry): the fill and outline are an SVG path from the shared DrawingML
 * geometry, with the text box laid over it.
 */
function renderPathShape(shape: DocxShape, deps: RenderDeps, el: HTMLElement): HTMLElement {
  const s = el.style;
  s.position = 'relative';
  s.display = 'inline-flex';
  s.flexDirection = 'column';
  s.justifyContent = shape.vAnchor === 'bottom' ? 'flex-end' : shape.vAnchor === 'ctr' ? 'center' : 'flex-start';
  if (shape.textColorHex) s.color = `#${shape.textColorHex}`;
  else if (shape.fillHex && isDarkFill(shape.fillHex)) s.color = '#fff';

  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  const w = shape.widthPx;
  const h = shape.heightPx;
  svg.setAttribute('width', String(w));
  svg.setAttribute('height', String(h));
  svg.setAttribute('viewBox', `0 0 ${w} ${h}`);
  svg.style.cssText = 'position:absolute;left:0;top:0;overflow:visible;';
  const open = shape.preset ? OPEN_PRESETS.has(shape.preset) : false;
  // A preset is generated in the shape's px box; a custom path is in its own
  // coordinate space, stretched to the box by a nested viewBox.
  const paths: { d: string; host: SVGElement }[] = [];
  if (shape.preset) paths.push({ d: presetPath(shape.preset, w, h, shape.adjust ?? {}), host: svg });
  for (const cp of shape.paths ?? []) {
    const box = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    box.setAttribute('width', String(w));
    box.setAttribute('height', String(h));
    box.setAttribute('viewBox', `0 0 ${cp.w || w} ${cp.h || h}`);
    box.setAttribute('preserveAspectRatio', 'none');
    box.setAttribute('overflow', 'visible');
    svg.appendChild(box);
    paths.push({ d: cp.d, host: box });
  }
  for (const { d, host } of paths) {
    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', d);
    // Keep the outline its stated width however the path is stretched.
    path.setAttribute('vector-effect', 'non-scaling-stroke');
    path.setAttribute('fill', !open && shape.fillHex ? `#${shape.fillHex}` : 'none');
    if (shape.preset && EVENODD_PRESETS.has(shape.preset)) path.setAttribute('fill-rule', 'evenodd');
    if (shape.lineHex) {
      path.setAttribute('stroke', `#${shape.lineHex}`);
      path.setAttribute('stroke-width', String(shape.lineWidthPx ?? 1));
    }
    host.appendChild(path);
  }
  el.appendChild(svg);
  if (shape.fillImagePart) {
    const url = deps.imageUrl(shape.fillImagePart);
    if (url) {
      s.backgroundImage = `url("${url}")`;
      s.backgroundSize = '100% 100%';
    }
  }

  if (shape.content.length) {
    const inner = document.createElement('div');
    stackBlocks(inner);
    inner.style.position = 'relative';
    inner.style.boxSizing = 'border-box';
    const ins = shape.insetsPx ?? { top: 4.8, right: 9.6, bottom: 4.8, left: 9.6 };
    inner.style.padding = `${ins.top}px ${ins.right}px ${ins.bottom}px ${ins.left}px`;
    for (const block of shape.content) inner.appendChild(renderBlock(block, deps));
    el.appendChild(inner);
  }
  return el;
}

// ─── Equations ───────────────────────────────────────────────────────────────

const MATHML_NS = 'http://www.w3.org/1998/Math/MathML';

/** An equation as a MathML element, uneditable (Word regenerates its layout). */
function renderMath(node: MathNode, display: boolean, sizePx: number): Element {
  const math = document.createElementNS(MATHML_NS, 'math');
  if (display) math.setAttribute('displaystyle', 'true');
  math.setAttribute('contenteditable', 'false');
  math.setAttribute('style', `font-family:"Cambria Math","STIX Two Math","Latin Modern Math",math;font-size:${px(sizePx)};`);
  math.appendChild(toMathML(node));
  return math;
}

function toMathML(n: MathNode): Element {
  const mk = (name: string, kids: Element[] = []) => {
    const el = document.createElementNS(MATHML_NS, name);
    for (const k of kids) el.appendChild(k);
    return el;
  };
  const mo = (v: string, attrs: Record<string, string> = {}) => {
    const el = mk('mo');
    el.textContent = v;
    for (const [k, val] of Object.entries(attrs)) el.setAttribute(k, val);
    return el;
  };
  switch (n.t) {
    case 'row':
      return mk('mrow', n.c.map(toMathML));
    case 'tok': {
      const el = mk(n.kind);
      el.textContent = n.v;
      if (n.kind === 'mi' && n.normal) el.setAttribute('mathvariant', 'normal');
      return el;
    }
    case 'frac': {
      if (n.linear) return mk('mrow', [toMathML(n.num), mo('/'), toMathML(n.den)]);
      const f = mk('mfrac', [toMathML(n.num), toMathML(n.den)]);
      if (n.noBar) f.setAttribute('linethickness', '0');
      return f;
    }
    case 'script': {
      const base = toMathML(n.base);
      if (n.pre) {
        return mk('mmultiscripts', [base, mk('mprescripts'), n.sub ? toMathML(n.sub) : mk('none'), n.sup ? toMathML(n.sup) : mk('none')]);
      }
      if (n.sub && n.sup) return mk('msubsup', [base, toMathML(n.sub), toMathML(n.sup)]);
      if (n.sub) return mk('msub', [base, toMathML(n.sub)]);
      return mk('msup', [base, toMathML(n.sup!)]);
    }
    case 'root':
      return n.deg ? mk('mroot', [toMathML(n.e), toMathML(n.deg)]) : mk('msqrt', [toMathML(n.e)]);
    case 'fence': {
      const kids: Element[] = [];
      if (n.open) kids.push(mo(n.open, { fence: 'true', stretchy: 'true' }));
      n.items.forEach((item, i) => {
        if (i > 0) kids.push(mo(n.sep, { separator: 'true' }));
        kids.push(toMathML(item));
      });
      if (n.close) kids.push(mo(n.close, { fence: 'true', stretchy: 'true' }));
      return mk('mrow', kids);
    }
    case 'nary': {
      // Limits go above/below in a display equation and beside the operator
      // inline, as Word sets them.
      const op = mo(n.op, { largeop: 'true', movablelimits: 'true' });
      const limits = n.sub && n.sup ? [toMathML(n.sub), toMathML(n.sup)] : n.sub ? [toMathML(n.sub)] : n.sup ? [toMathML(n.sup)] : [];
      let head: Element = op;
      if (limits.length) {
        const name = n.limits
          ? n.sub && n.sup ? 'munderover' : n.sub ? 'munder' : 'mover'
          : n.sub && n.sup ? 'msubsup' : n.sub ? 'msub' : 'msup';
        head = mk(name, [op, ...limits]);
      }
      return mk('mrow', [head, toMathML(n.e)]);
    }
    case 'under':
      return mk('munder', [toMathML(n.base), toMathML(n.under)]);
    case 'over': {
      const o = mk('mover', [toMathML(n.base), toMathML(n.over)]);
      if (n.accent) o.setAttribute('accent', 'true');
      return o;
    }
    case 'table': {
      const t = mk('mtable', n.rows.map((r) => mk('mtr', r.map((c) => mk('mtd', [toMathML(c)])))));
      if (n.align === 'left') t.setAttribute('columnalign', 'left');
      return t;
    }
    case 'box':
      return mk('mrow', [toMathML(n.e)]);
  }
}

// ─── Helpers ───────────────────────────────────────────────────────────────────

function strokeCss(stroke: DocxStroke): string {
  const style = stroke.style ?? 'solid';
  // Exact widths — a 1px floor adds a fraction of a point to every bordered
  // row, which over a long table moves the page break. A CSS double border
  // needs at least 3px to show both lines.
  const width = style === 'double' ? Math.max(3, stroke.width * 3) : stroke.width;
  return `${width}px ${style} #${stroke.color.hex}`;
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
