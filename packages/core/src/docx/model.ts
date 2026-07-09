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
}

export interface DocxTable {
  kind: 'table';
  /** Total table width in px, if specified. */
  widthPx?: number;
  /** Column widths in px. */
  colWidths: number[];
  /** rows[r][c] — null for cells covered by a span. */
  rows: (DocxTableCell | null)[][];
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

// ─── Page ────────────────────────────────────────────────────────────────────

/**
 * A single rendered page. Width is fixed to the section page size; in the
 * continuous-flow pass height is auto (content-driven) so nothing is clipped.
 * The viewer scales pages to fit the container width and stacks them vertically.
 */
export interface DocxPage {
  index: number;
  size: DocxPageSize;
  margins: DocxPageMargins;
  elements: DocxBlock[];
  header?: DocxBlock[];
  footer?: DocxBlock[];
}

// ─── Document ──────────────────────────────────────────────────────────────────

export interface DocxDocumentData {
  pages: DocxPage[];
  embeddedFonts: import('../pptx/model.js').EmbeddedFont[];
}
