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
import type { Chart } from '../oxml/drawingml/chart/model.js';
import type { CustomPath } from '../oxml/drawingml/geometry/custom.js';

export type {
  Color,
  Fill,
  GradientStop,
  Stroke,
  TextAlign,
  EmbeddedFont,
  EmbeddedFontFace,
} from '../pptx/model.js';

/** A Word border: a stroke plus its line style (<w:… w:val>). */
export interface DocxStroke extends Stroke {
  style?: 'solid' | 'double' | 'dashed' | 'dotted';
}

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
   * The tab before this run is an absolute-position tab (<w:ptab>), aligned to
   * the margins rather than to the paragraph's tab stops.
   */
  ptab?: 'left' | 'center' | 'right';
  /** A page or column break (<w:br w:type="page|column"/>) precedes this run. */
  pageBreakBefore?: boolean;
  /** The break before this run is a column break (it ends the column, not the page). */
  columnBreak?: boolean;
  /** Hidden text (<w:vanish/>): kept in the model so editing preserves it, never drawn. */
  hidden?: boolean;
  /** Extra space between characters, in points (<w:spacing w:val>; negative condenses). */
  charSpacingPt?: number;
  /** Pair kerning applies (<w:kern> and the run is at least that size). Word kerns nothing otherwise. */
  kern?: boolean;
  /**
   * A footnote/endnote reference mark (<w:footnoteReference>): {@link text} is
   * its number, and the note's content travels with it so the paginator can
   * place it at the foot of the page the reference lands on.
   */
  note?: { kind: 'footnote' | 'endnote'; blocks: DocxBlock[] };
  /** The short rule above a page's footnotes (<w:separator/>). */
  separator?: boolean;
  /** Right-to-left run (<w:rtl>): complex-script text laid out right to left. */
  rtl?: boolean;
  /** Font for East Asian characters in this run (<w:rFonts w:eastAsia>). */
  eastAsiaFont?: string;
  /** An equation (<m:oMath>) occupying this run's slot; {@link text} is empty. */
  math?: MathNode;
  /** A tracked change this run belongs to (<w:ins>/<w:del>, or a move). */
  revision?: DocxRevision;
  /** A tracked formatting change on this run (<w:rPrChange>). */
  formatChange?: { author: string; date?: string };
  /** Ids of the comments whose ranges cover this run. */
  commentIds?: string[];
  /** A comment's reference mark (<w:commentReference>), carrying the comment. */
  comment?: DocxComment;
  /**
   * An inline drawing (<w:drawing wp:inline> / inline VML) occupying this run's
   * slot in the text flow. When set, {@link text} is empty and the run renders
   * the drawing inline-block, so paragraph alignment (e.g. centered) applies.
   */
  drawing?: DocxDrawing;
}

/** A custom tab stop (<w:tab> in <w:tabs>). */
export interface DocxTabStop {
  /** Position in px from the paragraph's leading text-column edge (not its indent). */
  posPx: number;
  align: 'left' | 'center' | 'right' | 'decimal' | 'bar';
  /** Fill drawn across the tab's width. */
  leader?: 'dot' | 'hyphen' | 'underscore' | 'middleDot' | 'heavy';
}

/** Character formatting of a generated list marker (numbering level rPr over the mark's). */
export interface DocxMarkerFormat {
  font?: string;
  /** Font whose line height the marker imposes (a symbol font its glyph was mapped from). */
  metricsFont?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  colorHex?: string;
}

/** One tracked change: inserted or deleted text and who made it. */
export interface DocxRevision {
  kind: 'ins' | 'del';
  author: string;
  date?: string;
}

/** A comment (comments.xml): who wrote it, and its content as blocks. */
export interface DocxComment {
  id: string;
  author: string;
  initials?: string;
  date?: string;
  blocks: DocxBlock[];
}

// ─── Equations ────────────────────────────────────────────────────────────────

/**
 * An equation, converted from Office Math (OMML) into a small layout tree
 * that maps onto MathML (which browsers draw natively).
 */
export type MathNode =
  | { t: 'row'; c: MathNode[] }
  /** A token: identifier (italic by default), number, operator, or plain text. */
  | { t: 'tok'; kind: 'mi' | 'mn' | 'mo' | 'mtext'; v: string; normal?: boolean }
  | { t: 'frac'; num: MathNode; den: MathNode; noBar?: boolean; linear?: boolean }
  | { t: 'script'; base: MathNode; sub?: MathNode; sup?: MathNode; pre?: boolean }
  | { t: 'root'; e: MathNode; deg?: MathNode }
  | { t: 'fence'; open: string; close: string; sep: string; items: MathNode[] }
  | { t: 'nary'; op: string; sub?: MathNode; sup?: MathNode; e: MathNode; limits: boolean }
  | { t: 'under'; base: MathNode; under: MathNode }
  | { t: 'over'; base: MathNode; over: MathNode; accent?: boolean }
  | { t: 'table'; rows: MathNode[][]; align?: 'left' | 'center' }
  | { t: 'box'; e: MathNode };

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
  /** Picture fill (<a:blipFill>): the media part, stretched over the shape. */
  fillImagePart?: string;
  /** Preset geometry name (<a:prstGeom prst>) and its adjust values, for an SVG outline. */
  preset?: string;
  adjust?: Record<string, number>;
  /** Custom geometry (<a:custGeom>) paths. */
  paths?: CustomPath[];
  /** Rotation in degrees (<a:xfrm rot>). */
  rot?: number;
  /** Outline color hex (no leading #). */
  lineHex?: string;
  lineWidthPx?: number;
  /** Text-box content (<w:txbxContent>), rendered inside the shape. */
  content: DocxBlock[];
  /** Vertical anchoring of the text box: 'ctr' | 'top' | 'bottom'. */
  vAnchor?: 'top' | 'ctr' | 'bottom';
  /** Default text colour from the shape style (<wps:style><a:fontRef>). */
  textColorHex?: string;
  /** Extra layout room around an inline shape for its effects (<wp:effectExtent>), px. */
  effectPx?: { top: number; right: number; bottom: number; left: number };
  /** Text direction in the box (<wps:bodyPr vert>): rotated or stacked East Asian. */
  vert?: 'vert' | 'vert270' | 'eaVert';
  /** Text-box insets in px (<wps:bodyPr> lIns/tIns/rIns/bIns). */
  insetsPx?: { top: number; right: number; bottom: number; left: number };
  alt?: string;
}

/** A chart (DrawingML chart part), drawn from its cached values. */
export interface DocxChart {
  kind: 'chart';
  chart: Chart;
  widthPx: number;
  heightPx: number;
  /** Extra layout room around an inline chart (<wp:effectExtent>), px. */
  effectPx?: { top: number; right: number; bottom: number; left: number };
}

/**
 * Drawings positioned inside a box: a Word shape group (<wpg:wgp>) or a
 * SmartArt diagram's pre-drawn shapes, each placed in the group's px space.
 */
export interface DocxGroup {
  kind: 'group';
  widthPx: number;
  heightPx: number;
  items: { xPx: number; yPx: number; wPx: number; hPx: number; rot?: number; drawing: DocxDrawing }[];
  effectPx?: { top: number; right: number; bottom: number; left: number };
}

/** A picture, a DrawingML shape, a chart, or a group of drawings. */
export type DocxDrawing = DocxInlineImage | DocxShape | DocxChart | DocxGroup;

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
  /** Stacking order among floating drawings (<wp:anchor relativeHeight>), higher on top. */
  zOrder?: number;
  wrap: DocxWrap;
  /** Which side(s) text may flow past the drawing (<wp:wrapSquare wrapText>). */
  wrapText?: 'bothSides' | 'left' | 'right' | 'largest';
  /** Gap kept between the drawing and wrapped text (<wp:anchor distT/B/L/R>), px. */
  distPx?: { top: number; right: number; bottom: number; left: number };
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
  /** Right-to-left paragraph (<w:bidi>): text starts at the right; left/right indents mirror. */
  rtl?: boolean;
  /** A display equation paragraph (<m:oMathPara>): equations centred on their own lines. */
  mathDisplay?: { align: 'left' | 'center' | 'right' };
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
  /** Absolute line spacing in points (see {@link lineSpacingRule}). */
  lineSpacingPt?: number;
  /** How {@link lineSpacingPt} applies: exactly, or as a minimum. */
  lineSpacingRule?: 'exact' | 'atLeast';
  /** Bullet / number marker text, pre-resolved from numbering.xml. */
  listMarker?: string;
  /** Formatting of the list marker. */
  listMarkerFormat?: DocxMarkerFormat;
  /** What follows the marker (<w:suff>): a tab to the text indent (default), a space, or nothing. */
  listSuffix?: 'tab' | 'space' | 'nothing';
  /** List indentation level, 0-based. */
  level?: number;
  /** Custom tab stops, sorted by position. */
  tabs?: DocxTabStop[];
  /** Interval of the default tab stops past the last custom one, in px. */
  defaultTabPx?: number;
  /** Keep on the same page as the next paragraph (<w:keepNext>). */
  keepNext?: boolean;
  /** Never split across pages (<w:keepLines>). */
  keepLines?: boolean;
  /** Avoid a single first/last line alone on a page (<w:widowControl>). */
  widowControl?: boolean;
  pageBreakBefore?: boolean;
  /** With {@link pageBreakBefore}: the break is a column break (next column, not next page). */
  columnBreakBefore?: boolean;
  shadingHex?: string;
  /** Suppress spacing between consecutive paragraphs with the same style. */
  contextualSpacing?: boolean;
  /** Paragraph-level borders (from <w:pBdr>). */
  paraBorders?: Partial<{ top: DocxStroke; bottom: DocxStroke; left: DocxStroke; right: DocxStroke }>;
  /** Gap between each paragraph border and the text, in px (<w:pBdr> w:space). */
  paraBorderSpacePx?: Partial<{ top: number; bottom: number; left: number; right: number }>;
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
  /** <w:tblW w:type="pct">: width as a percentage of the text column. */
  widthPct?: number;
  /** Horizontal alignment of the table within the text column (<w:jc>). */
  align?: 'left' | 'center' | 'right';
  /** Per-row height rules (<w:trHeight>), parallel to {@link rows}. */
  rowHeights?: (DocxRowHeight | undefined)[];
  /** Rows repeated at the top of each page the table continues onto (<w:tblHeader>). */
  headerRows?: number;
  /** Rows that must not split across a page (<w:cantSplit>), by row index. */
  cantSplit?: boolean[];
}

export interface DocxRowHeight {
  px: number;
  /** 'exact' clips the row to this height; 'atLeast' lets content grow it. */
  rule: 'exact' | 'atLeast';
}

export interface DocxTableCell {
  content: DocxBlock[];
  fillHex?: string;
  rowSpan: number;
  colSpan: number;
  borders?: Partial<{
    l: DocxStroke;
    t: DocxStroke;
    r: DocxStroke;
    b: DocxStroke;
  }>;
  vAlign?: 'top' | 'center' | 'bottom';
  /** A filler for grid columns a row skips (<w:gridBefore>/<w:gridAfter>): no borders, no content. */
  empty?: boolean;
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
  /** Extra layout room around an inline picture for its effects (<wp:effectExtent>), px. */
  effectPx?: { top: number; right: number; bottom: number; left: number };
  /**
   * <a:duotone>: the picture recoloured as a gradient from the first colour
   * (its darks) to the second (its lights). Hex, no leading #.
   */
  duotone?: [string, string];
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
  /**
   * Even-page header/footer (<w:type="even">), used on even-numbered pages when
   * the document sets <w:evenAndOddHeaders>; undefined there means a blank one.
   */
  evenHeader?: DocxBlock[];
  evenFooter?: DocxBlock[];
  /** The document distinguishes odd and even pages' headers and footers. */
  evenAndOddHeaders?: boolean;
  /** <w:titlePg>: the first page uses the first-page header/footer set. */
  titlePg?: boolean;
  /**
   * How the section begins (<w:type>): 'continuous' flows on from the previous
   * section's last page instead of starting a new one.
   */
  start?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';
  /** <w:pgNumType w:start>: the number the section's first page restarts at. */
  pageNumberStart?: number;
  /** <w:pgNumType w:fmt>: decimal, lowerRoman, upperLetter, … */
  pageNumberFormat?: string;
  /** The rule drawn above a page's footnotes (the footnote separator's content). */
  footnoteSeparator?: DocxBlock[];
  /** Text columns (<w:cols>), when the section has more than one. */
  columns?: DocxColumns;
}

/** A section's text columns: each one's offset and width within the text area. */
export interface DocxColumns {
  cols: { xPx: number; widthPx: number }[];
  /** <w:cols w:sep>: a rule between columns. */
  separator: boolean;
}

/** A horizontal band of a page laid out in columns (one section's run on that page). */
export interface DocxBand {
  columns: { xPx: number; widthPx: number }[];
  /** Blocks of each column, in order. */
  blocks: DocxBlock[][];
  heightPx: number;
  separator?: boolean;
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
  /** The page number as displayed (PAGE field), after section restarts and format. */
  pageNumber?: string;
  /** Total pages in the document (NUMPAGES field). */
  pageCount?: number;
  /** Footnotes referenced on this page, separator first, drawn above the bottom margin. */
  footnotes?: DocxBlock[];
  /**
   * The page's content as bands of columns, top to bottom — present when any
   * part of the page is set in more than one column. {@link elements} then
   * still lists every block, in reading order.
   */
  bands?: DocxBand[];
  /** Pages in this page's section (SECTIONPAGES field). */
  sectionPageCount?: number;
}

// ─── Document ──────────────────────────────────────────────────────────────────

export interface DocxDocumentData {
  sections: DocxSection[];
  embeddedFonts: import('../pptx/model.js').EmbeddedFont[];
}
