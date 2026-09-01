/**
 * Shared DOM-rendering primitives used across every feature renderer.
 *
 * Keeps the cross-cutting pieces — the render dependency contract, positioning
 * a shape from its transform, and painting a {@link Fill} onto an element's CSS
 * background — in one place so feature slices (shapes/text/tables/pictures/…)
 * don't duplicate them.
 */
import type { Fill, Shape, TextBody, TextRun, Transform } from '../model.js';
import { colorToCss } from '../color.js';
import { EDIT_ATTR } from '../../oxml/edit/attrs.js';

export const SVG_NS = 'http://www.w3.org/2000/svg';

/**
 * How a coordinate space lands on the slide.
 *
 * A shape's transform is stated in the space of whatever contains it: the slide
 * for a top-level shape, but a group's own child space (`chOff`/`chExt`) for
 * anything inside a group. `renderGroup` realises that space as a CSS
 * scale+translate, so one shape's box can be mapped to the slide with
 * `slideX = ox + sx * x`. The selection overlay lives in slide space, so it
 * needs this to draw handles on a shape nested in a group — and to take a drag
 * measured on screen back into the space the shape's XML is written in.
 */
export interface ShapeFrame {
  /** Slide-space position of this space's origin. */
  ox: number;
  oy: number;
  /** Scale from this space to slide space. */
  sx: number;
  sy: number;
  /**
   * True when an ancestor group is rotated or mirrored, which the two numbers
   * above cannot express. Manipulating a shape inside one would draw handles in
   * the wrong place, so the editor keeps the group itself as the unit instead.
   */
  turned: boolean;
}

/** The slide's own space: the identity mapping. */
export const SLIDE_FRAME: ShapeFrame = { ox: 0, oy: 0, sx: 1, sy: 1, turned: false };

/**
 * What the renderer needs in order to make text editable. Supplied by the edit
 * layer; absent for a read-only render, which is the default. Declared here so
 * the render slice owns its own dependency contract rather than importing from
 * `pptx/edit`.
 */
export interface EditRenderContext {
  /** Stable key identifying a model object for this render pass. */
  key(model: object): string;
  /** True if this text body belongs to the slide itself and may be edited. */
  editable(body: TextBody): boolean;
  /**
   * True for runs whose text PowerPoint generates (`<a:fld>` — slide number,
   * date). They render but must not be typed into.
   */
  isField(run: TextRun): boolean;
  /**
   * True if this shape is the slide's own and may be selected and moved. A
   * shape composited in from the layout or master is drawn but belongs to every
   * slide sharing it, so it is not the user's to reposition here.
   */
  selectable?(shape: Shape): boolean;
  /**
   * True if this text body is the one open for typing. Selection comes first —
   * a single click picks the shape — so only the body the user has explicitly
   * entered is made contentEditable.
   */
  textEditing?(body: TextBody): boolean;
  /**
   * Record which coordinate space a selectable shape was drawn in, so the
   * editor can map its box to the slide. Called only for shapes that are
   * marked selectable; absent for the slide's own space, which is the identity.
   */
  shapeFrame?(shape: Shape, frame: ShapeFrame): void;
}

export interface RenderDeps {
  /** Resolve a media part path to a displayable URL. */
  imageUrl(part: string): string | undefined;
  /** Present only when the deck is being rendered for editing. */
  edit?: EditRenderContext;
  /**
   * Cumulative non-uniform scale of all ancestor groups, as a product of each
   * group's (off/chExt) ratio. Geometry is squished by the ancestor CSS scale,
   * but PowerPoint never squishes glyphs: text keeps its aspect ratio and just
   * reflows in the scaled box (font follows the vertical scale). The text
   * renderer reads this to counter the horizontal squish. Absent ⇒ no scaling.
   */
  groupScale?: { sx: number; sy: number };
  /**
   * The coordinate space the shapes being rendered are stated in. Absent at the
   * top level, where it is the slide's own space; {@link renderGroup} composes a
   * new one for a group's children.
   */
  frame?: ShapeFrame;
}

/**
 * Mark a shape as selectable, if the edit layer says it is, and record the
 * space it was drawn in. Both the slide renderer and `renderGroup` go through
 * this, which is what makes a shape inside a group addressable at all.
 */
export function markSelectable(el: HTMLElement, shape: Shape, deps: RenderDeps): void {
  const edit = deps.edit;
  if (!edit?.selectable?.(shape)) return;
  el.setAttribute(EDIT_ATTR.shape, edit.key(shape));
  edit.shapeFrame?.(shape, deps.frame ?? SLIDE_FRAME);
}

/** Position + rotate/flip a shape container per its transform. */
export function positioned(transform: Transform | undefined): HTMLDivElement {
  const el = document.createElement('div');
  el.style.position = 'absolute';
  if (transform) {
    el.style.left = `${transform.x}px`;
    el.style.top = `${transform.y}px`;
    el.style.width = `${transform.w}px`;
    el.style.height = `${transform.h}px`;
    const parts: string[] = [];
    if (transform.rot) parts.push(`rotate(${transform.rot}deg)`);
    if (transform.flipH || transform.flipV) {
      parts.push(`scale(${transform.flipH ? -1 : 1}, ${transform.flipV ? -1 : 1})`);
    }
    if (parts.length) {
      el.style.transform = parts.join(' ');
      el.style.transformOrigin = 'center';
    }
  }
  return el;
}

/** Apply a {@link Fill} to an element's CSS background. */
export function applyFillBackground(el: HTMLElement, fill: Fill, deps: RenderDeps): void {
  switch (fill.type) {
    case 'solid':
      el.style.background = colorToCss(fill.color);
      break;
    case 'gradient': {
      const stops = fill.stops.map((s) => `${colorToCss(s.color)} ${(s.pos * 100).toFixed(1)}%`);
      if (fill.radial) {
        const at = fill.center
          ? ` at ${(fill.center.x * 100).toFixed(1)}% ${(fill.center.y * 100).toFixed(1)}%`
          : '';
        el.style.background = `radial-gradient(farthest-corner${at}, ${stops.join(',')})`;
      } else {
        el.style.background = `linear-gradient(${fill.angle ?? 0}deg, ${stops.join(',')})`;
      }
      break;
    }
    case 'image': {
      const url = deps.imageUrl(fill.part);
      if (url) {
        el.style.backgroundImage = `url("${url}")`;
        el.style.backgroundSize = 'cover';
        el.style.backgroundPosition = 'center';
      }
      break;
    }
    case 'none':
      break;
  }
}
