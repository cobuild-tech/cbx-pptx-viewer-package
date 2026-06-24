/**
 * Render-agnostic intermediate representation (IR) for DOCX documents.
 *
 * Mirrors the pptx/model.ts patterns: all lengths are in CSS pixels (96 DPI),
 * colors are resolved hex strings. The DOM renderer consumes only this model —
 * it never touches XML.
 *
 * Shared types (TextRun, Bullet, Color, Fill, Stroke) are re-exported from the
 * PPTX model since Word uses the same DrawingML primitives.
 */
import type { Stroke } from '../pptx/model.js';

export type {
  Color,
  Fill,
  GradientStop,
  Stroke,
  TextRun,
  Bullet,
  TextAlign,
  EmbeddedFont,
  EmbeddedFontFace,
} from '../pptx/model.js';

// ─── Page geometry ──────────────────────────────────────────────────────────

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

// ─── Block elements ──────────────────────────────────────────────────────────

export type DocxBlock = DocxParagraph | DocxTable | DocxInlineImage;

export interface DocxParagraph {
  kind: 'paragraph';
  runs: import('../pptx/model.js').TextRun[];
  /** Resolved style name, e.g. 'Normal', 'Heading1'. */
  styleName: string;
  /** Paragraph default font family from the style chain (not hardcoded). */
  baseFontFamily?: string;
  /** Paragraph default font size in pt from the style chain (not hardcoded). */
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
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  /** Line spacing as a multiple of font size (1 = single, 1.5, 2, etc.). */
  lineSpacingPct?: number;
  /** Absolute line spacing in points. */
  lineSpacingPt?: number;
  bullet?: import('../pptx/model.js').Bullet;
  /** List indentation level, 0-based. */
  level?: number;
  keepTogether?: boolean;
  pageBreakBefore?: boolean;
  shadingHex?: string;
  /** Right indent in px. */
  indentRightPx?: number;
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
  content: DocxParagraph[];
  fill: import('../pptx/model.js').Fill;
  rowSpan: number;
  colSpan: number;
  borders?: Partial<{
    l: import('../pptx/model.js').Stroke;
    t: import('../pptx/model.js').Stroke;
    r: import('../pptx/model.js').Stroke;
    b: import('../pptx/model.js').Stroke;
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
 * A single rendered page. Width is fixed to the section page size; height is
 * auto (content-driven) so nothing is clipped. The viewer navigates pages like
 * PPTX slides, scaling by width to fit the container.
 */
export interface DocxPage {
  index: number;
  size: DocxPageSize;
  margins: DocxPageMargins;
  elements: DocxBlock[];
  header?: DocxBlock[];
  footer?: DocxBlock[];
}

// ─── Document ────────────────────────────────────────────────────────────────

export interface DocxDocumentData {
  pages: DocxPage[];
  embeddedFonts: import('../pptx/model.js').EmbeddedFont[];
}
