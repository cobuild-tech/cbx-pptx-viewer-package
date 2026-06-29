/**
 * Render-agnostic intermediate representation (IR).
 *
 * Parsers turn XML into this model and the resolver fills in everything a slide
 * inherits (colors, fonts, placeholder geometry, text styles). The DOM renderer
 * consumes only this model — it never touches XML. All lengths are in CSS pixels
 * in the slide's base coordinate space (see {@link units}); the viewer scales the
 * whole slide to fit the viewport.
 */

/** A resolved color: sRGB hex (no leading #) plus optional alpha 0..1. */
export interface Color {
  hex: string;
  alpha?: number;
}

export interface GradientStop {
  /** Position along the gradient, 0..1. */
  pos: number;
  color: Color;
}

export type Fill =
  | { type: 'none' }
  | { type: 'solid'; color: Color }
  | {
      type: 'gradient';
      stops: GradientStop[];
      /** Linear angle in degrees (CSS convention). Omitted for radial. */
      angle?: number;
      radial: boolean;
    }
  | {
      type: 'image';
      /** Resolved media part path; resolve to a URL via the Deck. */
      part: string;
      /** Source-rectangle crop as fractions {l,t,r,b} of the image. */
      crop?: { l: number; t: number; r: number; b: number };
    };

/** Arrowhead/marker on a line end (`<a:headEnd>` / `<a:tailEnd>`). */
export type LineEndType = 'triangle' | 'arrow' | 'stealth' | 'diamond' | 'oval';
export interface LineEnd {
  type: LineEndType;
  /** Marker width/length category across/along the line. */
  w: 'sm' | 'med' | 'lg';
  len: 'sm' | 'med' | 'lg';
}

export interface Stroke {
  color: Color;
  /** Width in px. */
  width: number;
  /** Dash pattern in px (CSS stroke-dasharray), if dashed. */
  dash?: number[];
  cap?: 'butt' | 'round' | 'square';
  /** Arrowhead at the start of the line, if any. */
  headEnd?: LineEnd;
  /** Arrowhead at the end of the line, if any. */
  tailEnd?: LineEnd;
}

/**
 * A visual effect from `<a:effectLst>` (or a theme `effectRef`). Offsets/radii
 * are in px; the renderer maps these to CSS filters / box-reflect.
 */
export type Effect =
  | { type: 'outerShadow'; dx: number; dy: number; blur: number; color: Color }
  | { type: 'innerShadow'; dx: number; dy: number; blur: number; color: Color }
  | { type: 'glow'; radius: number; color: Color }
  | { type: 'softEdge'; radius: number }
  | { type: 'reflection'; blur: number; dist: number; startAlpha: number; endAlpha: number };

export type TextAlign = 'l' | 'ctr' | 'r' | 'just';
export type VerticalAnchor = 'top' | 'ctr' | 'bottom';

export interface TextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Resolved font size in points. */
  sizePt?: number;
  color?: Color;
  /** Resolved typeface name. */
  font?: string;
  /** Super/subscript as a percentage (positive = super, negative = sub). */
  baseline?: number;
  /** Highlighter color behind the run text (`<a:highlight>`). */
  highlight?: Color;
  /** Character spacing in points (`spc`, can be negative). */
  letterSpacingPt?: number;
  /** Small caps / all caps (`cap` attribute). */
  caps?: 'all' | 'small';
  hyperlink?: string;
}

export type Bullet =
  | { type: 'none' }
  | { type: 'char'; char: string; font?: string; color?: Color; sizePct?: number }
  | { type: 'number'; scheme: string; startAt?: number; color?: Color; sizePct?: number };

export interface Paragraph {
  runs: TextRun[];
  align?: TextAlign;
  /** Indentation level 0..8. */
  level: number;
  bullet?: Bullet;
  /** Left margin in px. */
  marginLeftPx?: number;
  /** First-line indent in px (can be negative for hanging bullets). */
  indentPx?: number;
  /** Line spacing: either a multiple (pct, 1 = single) or an absolute pt value. */
  lineSpacingPct?: number;
  lineSpacingPt?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
}

export interface TextBody {
  paragraphs: Paragraph[];
  anchor: VerticalAnchor;
  wrap: boolean;
  /** Internal insets in px: left, top, right, bottom. */
  insets: { l: number; t: number; r: number; b: number };
  /** normAutofit shrink factors, if the box shrinks text to fit. */
  fontScale?: number;
  lnSpcReductionPct?: number;
}

export interface Transform {
  x: number;
  y: number;
  w: number;
  h: number;
  /** Rotation in degrees, clockwise. */
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
}

export type Geometry =
  | { type: 'preset'; preset: string; adjust: Record<string, number> }
  | { type: 'custom'; paths: CustomPath[] };

export interface CustomPath {
  /** SVG path data in this path's own coordinate space. */
  d: string;
  /** Path coordinate-space width/height (from <a:path w= h=>), in EMU. */
  w: number;
  h: number;
}

interface ShapeBase {
  id?: string;
  name?: string;
  transform?: Transform;
  /** Visual effects (shadow/glow/reflection/soft edge), in document order. */
  effects?: Effect[];
}

export interface PresetShape extends ShapeBase {
  kind: 'shape';
  geom: Geometry;
  fill: Fill;
  stroke?: Stroke;
  text?: TextBody;
  /**
   * Explicit text rectangle in the same (absolute) coordinate space as
   * {@link transform}, from a diagram drawing's `<dsp:txXfrm>`. When present the
   * text is laid out in this sub-box instead of filling the shape — SmartArt
   * uses it to place a label inside one wedge/segment of a larger shape.
   */
  textBox?: Transform;
  /** Placeholder type if this shape is a placeholder (title, body, etc.). */
  placeholder?: { type: string; idx?: number };
}

export interface PictureShape extends ShapeBase {
  kind: 'picture';
  /** Resolved media part path. */
  part: string;
  fill: Fill;
  stroke?: Stroke;
  crop?: { l: number; t: number; r: number; b: number };
  /** Shape the image is clipped to (e.g. a circle); rectangle if omitted. */
  geom?: Geometry;
  placeholder?: { type: string; idx?: number };
}

export interface GroupShape extends ShapeBase {
  kind: 'group';
  children: Shape[];
  /** Child coordinate origin/extent (chOff/chExt) in px, for nested mapping. */
  childOffset: { x: number; y: number; w: number; h: number };
}

export interface ConnectorShape extends ShapeBase {
  kind: 'connector';
  geom: Geometry;
  stroke?: Stroke;
}

export interface FrameShape extends ShapeBase {
  kind: 'frame';
  frameType: 'table' | 'chart' | 'diagram' | 'unknown';
  table?: Table;
  /** SmartArt: the diagram's pre-laid-out shapes, positioned in the frame box. */
  diagram?: Shape[];
  /** Parsed chart (bar/line/pie/…) from the referenced chart part. */
  chart?: Chart;
}

export type ChartKind = 'bar' | 'line' | 'pie' | 'doughnut' | 'area' | 'scatter';

/** How multi-series plots stack/cluster (`<c:grouping>`). */
export type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';

export interface ChartSeries {
  /** Series display name (legend / tooltip). */
  name?: string;
  /** Y values (bar height, line/area point, pie slice), ordered by category. */
  values: number[];
  /** Scatter/bubble X values, parallel to {@link values}. */
  xValues?: number[];
  /** Series-level fill color (bars/area/line), resolved against the theme. */
  color?: Color;
  /** Per-point colors (pie/doughnut slices, or individual bar `<c:dPt>`). */
  pointColors?: (Color | undefined)[];
}

/**
 * A chart parsed from a `chartN.xml` part. Only the cached values
 * (`numCache`/`strCache`) are read — we render the snapshot, not the live
 * spreadsheet link. Lengths/positions are computed by the renderer in the
 * frame's pixel box.
 */
export interface Chart {
  kind: ChartKind;
  /** Bar charts: true = horizontal bars (`barDir=bar`), false = columns (`col`). */
  barHorizontal?: boolean;
  grouping?: ChartGrouping;
  /** Shared category-axis labels. */
  categories: string[];
  series: ChartSeries[];
  title?: string;
  /** Legend position, or omitted when there is no legend. */
  legend?: 'r' | 'l' | 't' | 'b' | 'tr';
  /** Doughnut hole radius as a fraction 0..1 of the outer radius. */
  holeSize?: number;
  /** Whether the chart shows numeric data labels on points. */
  showValueLabels?: boolean;
}

export type Shape =
  | PresetShape
  | PictureShape
  | GroupShape
  | ConnectorShape
  | FrameShape;

export interface TableCell {
  text?: TextBody;
  fill: Fill;
  /** Column/row span (1 = no span). Cells covered by a span are omitted. */
  colSpan: number;
  rowSpan: number;
  borders: { l?: Stroke; t?: Stroke; r?: Stroke; b?: Stroke };
}

export interface Table {
  /** Column widths in px. */
  colWidths: number[];
  /** Row heights in px. */
  rowHeights: number[];
  /** rows[r][c] — cells covered by a span from another cell are null. */
  rows: (TableCell | null)[][];
}

export interface Slide {
  index: number;
  background: Fill;
  shapes: Shape[];
  /** Source slide part path, for round-tripping / editing later. */
  part: string;
}

export interface SlideSize {
  wPx: number;
  hPx: number;
}

/** One variant (weight/style) of an embedded font, pointing at its font part. */
export interface EmbeddedFontFace {
  weight: number;
  style: 'normal' | 'italic';
  /** Font part path inside the package (e.g. ppt/fonts/font1.fntdata). */
  part: string;
}

/** An embedded typeface and its available variants (from <p:embeddedFontLst>). */
export interface EmbeddedFont {
  typeface: string;
  faces: EmbeddedFontFace[];
}
