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

export interface Stroke {
  color: Color;
  /** Width in px. */
  width: number;
  /** Dash pattern in px (CSS stroke-dasharray), if dashed. */
  dash?: number[];
  cap?: 'butt' | 'round' | 'square';
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
