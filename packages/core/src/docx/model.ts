/**
 * Render-agnostic intermediate representation (IR) for DOCX documents.
 *
 * Mirrors the pptx/model.ts patterns: all lengths are in CSS pixels (96 DPI),
 * colors are resolved hex strings. The DOM renderer consumes only this model —
 * it never touches XML.
 *
 * Shared DrawingML primitives (Color, Fill, Stroke, TextAlign, EmbeddedFont)
 * are re-exported as *types* from the PPTX model since Word uses the same
 * primitives. This is a type-only import — it creates no runtime dependency on
 * the pptx/ slice.
 */
import type { Stroke } from '../pptx/model.js';

export type {
  Color,
  Fill,
  GradientStop,
  Stroke,
  TextAlign,
  EmbeddedFont,
  EmbeddedFontFace,
} from '../pptx/model.js';

// ─── Inline content ──────────────────────────────────────────────────────────

/** A contiguous run of text with uniform character formatting (<w:r>). */
export interface DocxRun {
  text: string;
  fieldCode?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Resolved font size in points. */
  sizePt?: number;
  /** Resolved text color hex (no leading #). */
  colorHex?: string;
  /** Highlight color hex behind the run (<w:highlight> / <w:shd>). */
  highlightHex?: string;
  /** Resolved typeface name. */
  font?: string;
  /** Super/subscript: 'super' | 'sub'. */
  vertAlign?: 'super' | 'sub';
  /** Small caps / all caps. */
  caps?: 'all' | 'small';
  /** Hyperlink target (resolved external URL or internal anchor). */
  hyperlink?: string;
  /** A hard line break inside the paragraph (<w:br/>) precedes this run's text. */
  breakBefore?: boolean;
  /** A tab stop (<w:tab/>) precedes this run's text. */
  tabBefore?: boolean;
  /**
   * An inline drawing (<w:drawing wp:inline> / inline VML) occupying this run's
   * slot in the text flow. When set, {@link text} is empty and the run renders
   * the drawing inline-block, so paragraph alignment (e.g. centered) applies.
   */
  drawing?: DocxDrawing;
}

// ─── Drawings (pictures + DrawingML shapes) ──────────────────────────────────

/** Preset shape geometry (<a:prstGeom prst>), reduced to what CSS can express. */
export type DocxGeom = 'rect' | 'roundRect' | 'ellipse' | 'line' | 'other';

/**
 * A DrawingML shape (<wps:wsp>) or VML shape: a filled/outlined box that may
 * contain its own block flow (a text box). Used for cover-page banners, callout
 * boxes, etc. Sized in px; the payload text is a nested block flow.
 */
export interface DocxShape {
  kind: 'shape';
  geom: DocxGeom;
  widthPx: number;
  heightPx: number;
  /** Solid fill hex (no leading #). */
  fillHex?: string;
  /** Outline color hex (no leading #). */
  lineHex?: string;
  lineWidthPx?: number;
  /** Text-box content (<w:txbxContent>), rendered inside the shape. */
  content: DocxBlock[];
  /** Vertical anchoring of the text box: 'ctr' | 'top' | 'bottom'. */
  vAnchor?: 'top' | 'ctr' | 'bottom';
  alt?: string;
}

/** Either a raster/vector picture or a DrawingML shape. */
export type DocxDrawing = DocxInlineImage | DocxShape;

/** Text-wrap mode of an anchored (floating) drawing. */
export type DocxWrap = 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';

/**
 * A floating (anchored, <wp:anchor>) drawing attached to its containing
 * paragraph. Positioned absolutely relative to the paragraph's content box; the
 * `relativeFrom` tokens + offset/align are resolved to CSS at render time
 * (where the page content width and margins are known). Keeping the anchor on
 * its paragraph — rather than hoisting it to page coordinates — preserves its
 * flow position, so e.g. an image anchored in a right-hand table cell renders on
 * the right.
 */
export interface DocxAnchor {
  drawing: DocxDrawing;
  wPx: number;
  hPx: number;
  /** Drawn behind the text layer (<wp:anchor behindDoc>). */
  behindDoc: boolean;
  wrap: DocxWrap;
  /** <wp:positionH relativeFrom>. */
  relH: string;
  /** <wp:positionV relativeFrom>. */
  relV: string;
  /** Horizontal offset in px (<wp:posOffset>), if absolute-positioned. */
  hOffsetPx?: number;
  /** Horizontal alignment (<wp:align>), if aligned. */
  hAlign?: 'left' | 'center' | 'right' | 'inside' | 'outside';
  /** Vertical offset in px. */
  vOffsetPx?: number;
  /** Vertical alignment. */
  vAlign?: 'top' | 'center' | 'bottom' | 'inside' | 'outside';
}

// ─── Page geometry ────────────────────────────────────────────────────────────

export interface DocxPageSize {
  wPx: number;
  hPx: number;
}

export interface DocxPageMargins {
  topPx: number;
  rightPx: number;
  bottomPx: number;
  leftPx: number;
  headerPx: number;
  footerPx: number;
}

// ─── Block elements ────────────────────────────────────────────────────────────

export type DocxBlock = DocxParagraph | DocxTable | DocxInlineImage;

export interface DocxParagraph {
  kind: 'paragraph';
  runs: DocxRun[];
  /** Resolved style name, e.g. 'Normal', 'Heading1'. */
  styleName: string;
  /** Paragraph default font family from the style chain (not hardcoded). */
  baseFontFamily?: string;
  /** Paragraph default font size in pt from the style chain. */
  baseFontSizePt?: number;
  /** Paragraph default bold from the style chain. */
  baseBold?: boolean;
  /** Paragraph default italic from the style chain. */
  baseItalic?: boolean;
  /** Paragraph default color hex (no #) from the style chain. */
  baseColorHex?: string;
  align?: import('../pptx/model.js').TextAlign;
  /** Left indent in px (includes list indent). */
  indentLeftPx?: number;
  /** First-line indent in px (negative = hanging indent for bullets). */
  indentFirstLinePx?: number;
  /** Right indent in px. */
  indentRightPx?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  /** Line spacing as a multiple of font size (1 = single, 1.5, 2, etc.). */
  lineSpacingPct?: number;
  /** Absolute line spacing in points. */
  lineSpacingPt?: number;
  /** Bullet / number marker text, pre-resolved from numbering.xml. */
  listMarker?: string;
  /** List indentation level, 0-based. */
  level?: number;
  keepTogether?: boolean;
  pageBreakBefore?: boolean;
  shadingHex?: string;
  /** Suppress spacing between consecutive paragraphs with the same style. */
  contextualSpacing?: boolean;
  /** Paragraph-level borders (from <w:pBdr>). */
  paraBorders?: Partial<{ top: Stroke; bottom: Stroke; left: Stroke; right: Stroke }>;
  /** Floating (anchored) drawings whose anchor sits in this paragraph. */
  anchors?: DocxAnchor[];
}

export interface DocxTable {
  kind: 'table';
  /** Total table width in px, if specified. */
  widthPx?: number;
  /** Column widths in px. */
  colWidths: number[];
  /** rows[r][c] — null for cells covered by a span. */
  rows: (DocxTableCell | null)[][];
  /** Table indent in px from <w:tblInd> (negative pulls left into the margin). */
  indentPx?: number;
}

export interface DocxTableCell {
  content: DocxBlock[];
  fillHex?: string;
  rowSpan: number;
  colSpan: number;
  borders?: Partial<{
    l: Stroke;
    t: Stroke;
    r: Stroke;
    b: Stroke;
  }>;
  vAlign?: 'top' | 'center' | 'bottom';
  /** Cell padding in px from <w:tcMar> or table-level <w:tblCellMar>. */
  cellPaddingPx?: { top: number; right: number; bottom: number; left: number };
}

export interface DocxInlineImage {
  kind: 'image';
  /** Resolved media part path inside the OPC package. */
  part: string;
  widthPx: number;
  heightPx: number;
  alt?: string;
}

// ─── Section (parse output) ──────────────────────────────────────────────────

/**
 * One document section (delimited by <w:sectPr>): its page geometry plus the
 * full block flow and the header/footer content for its pages. The paginator
 * flows a section's blocks into one or more fixed-size {@link DocxPage}s.
 */
export interface DocxSection {
  index: number;
  size: DocxPageSize;
  margins: DocxPageMargins;
  blocks: DocxBlock[];
  /** Default header/footer (<w:type="default">), applied to every page. */
  header?: DocxBlock[];
  footer?: DocxBlock[];
  /**
   * First-page header/footer (<w:type="first">), applied to the section's first
   * page when <w:titlePg> is set. Undefined means the first page has none (a
   * blank title-page header, which is how Word renders titlePg with no first ref).
   */
  firstHeader?: DocxBlock[];
  firstFooter?: DocxBlock[];
  /** <w:titlePg>: the first page uses the first-page header/footer set. */
  titlePg?: boolean;
}

// ─── Page (paginated output) ───────────────────────────────────────────────────

/**
 * A single fixed-size page produced by the paginator: the section page size,
 * the slice of blocks that fit within the content area (between the top and
 * bottom margins), and the section's header/footer drawn in the margin bands.
 */
export interface DocxPage {
  index: number;
  size: DocxPageSize;
  margins: DocxPageMargins;
  elements: DocxBlock[];
  header?: DocxBlock[];
  footer?: DocxBlock[];
  /**
   * Effective top/bottom padding in px for the sheet's content box. Grown past
   * the raw page margins when a tall header/footer would otherwise overlap the
   * body (Word reserves space for the header/footer). Falls back to margins.
   */
  contentTopPx?: number;
  contentBottomPx?: number;
  /** Resolved text for style references, keyed by style name (lowercased). */
  resolvedStyles?: Record<string, string>;
}

// ─── Document ──────────────────────────────────────────────────────────────────

export interface DocxDocumentData {
  sections: DocxSection[];
  embeddedFonts: import('../pptx/model.js').EmbeddedFont[];
}
