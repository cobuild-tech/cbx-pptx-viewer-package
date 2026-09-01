/**
 * Shape selection and direct manipulation: the handles, the drag, the marquee.
 *
 * The overlay is a **sibling of the slide, not a child of it**. The slide is
 * drawn in its own pixel space and CSS-scaled to fit the stage, so handles
 * living inside it would scale too — 8px grips turning into 3px specks on a
 * zoomed-out deck. Drawing them in stage space instead keeps them a constant
 * size at any zoom, at the cost of mapping coordinates in both directions,
 * which is what {@link toSlide} and {@link place} do.
 *
 * The overlay is transparent to the pointer except on the grips themselves, so
 * a press inside a selected shape still reaches the slide underneath and this
 * class can hit-test it the same way as any other press.
 *
 * Drags are previewed by mutating the rendered element's own inline geometry
 * and committed only on release: one XML write and one re-render per drag,
 * rather than per pointer event. Nothing here writes XML — it reports the
 * finished transform and lets the viewer's edit session own the change.
 */
import { installStyleSheet } from '../../oxml/stylesheet.js';
import { EDIT_ATTR } from '../../oxml/edit/attrs.js';
import type { Shape, Transform } from '../model.js';
import { SLIDE_FRAME, type ShapeFrame } from '../render/primitives.js';
import {
  boundsOf,
  frameRect,
  moveBox,
  resizeBox,
  rotateBox,
  unframeDelta,
  unframePoint,
  type Handle,
} from '../edit/geometry.js';

/** How the slide is currently placed inside the host, for coordinate mapping. */
export interface SlideView {
  slideEl: HTMLElement | null;
  /** Uniform scale applied to the slide element. */
  scale: number;
  /** Slide origin within the host, in host px. */
  left: number;
  top: number;
}

export interface ShapeSelectionOptions {
  /** Element the overlay mounts into and listens on — the viewer's holder. */
  host: HTMLElement;
  /** The current placement of the slide. Asked for on demand, never cached. */
  view(): SlideView;
  /** Resolve a render key stamped on the DOM back to its shape. */
  resolve(key: string): Shape | undefined;
  /**
   * The coordinate space a shape was drawn in. A shape inside a group states
   * its box in the group's child space, not the slide's. Omit for a caller with
   * no groups to worry about.
   */
  frameOf?(shape: Shape): ShapeFrame;
  /** Selection changed (by click, marquee, Escape or programmatically). */
  onChange(shapes: Shape[]): void;
  /** A drag finished: write these transforms. Never called for a no-op drag. */
  onCommit(edits: Array<{ shape: Shape; transform: Transform }>): void;
  /**
   * The user asked to type in this shape — a click on a shape that was already
   * selected, which is how PowerPoint opens a text box. `at` is the click in
   * client coordinates, so the caller can drop the caret where it landed.
   */
  onActivate(shape: Shape, at: { x: number; y: number }): void;
}

type Drag =
  | { kind: 'move' }
  | { kind: 'resize'; handle: Handle }
  | { kind: 'rotate' }
  | { kind: 'marquee' };

interface DragState {
  drag: Drag;
  /**
   * The shape was already the whole selection when the press began. Releasing
   * without dragging then means "let me type in this", the same second click
   * PowerPoint uses to enter a text box.
   */
  activates: boolean;
  /** Pointer down position, in slide coordinates. */
  from: { x: number; y: number };
  /** Each dragged shape with the transform (and space) it started from. */
  subjects: Array<{ shape: Shape; el: HTMLElement; start: Transform; frame: ShapeFrame }>;
  /** True once the pointer has moved past the click threshold. */
  live: boolean;
  /**
   * The shapes under the press, outermost first — a group, then the child of it
   * that was hit, and so on. Releasing without travel on an already-selected
   * group steps one level down this chain, which is how PowerPoint reaches
   * inside a group.
   */
  chain: Shape[];
  /** Index in `chain` the press actually picked. */
  picked: number;
}

const STYLE_ID = 'cbx-shape-select';
/** Pointer travel (host px) before a press becomes a drag rather than a click. */
const DRAG_THRESHOLD = 3;
/** Smallest shape a resize may produce, in slide px. */
const MIN_SIZE = 2;

export class ShapeSelection {
  private readonly opts: ShapeSelectionOptions;
  private readonly root: HTMLDivElement;
  private readonly frame: HTMLDivElement;
  private readonly marquee: HTMLDivElement;
  private readonly outlines: HTMLDivElement;
  private readonly disposeStyles: () => void;
  private shapes: Shape[] = [];
  /**
   * Groups the user has stepped inside, outermost first. A press picks the
   * shallowest shape that is not on this path, so an untouched selection picks
   * the group and an entered one picks its child.
   */
  private entered: Shape[] = [];
  private state: DragState | null = null;
  private enabled = true;
  private readonly onPointerDown: (e: PointerEvent) => void;
  private readonly onPointerMove: (e: PointerEvent) => void;
  private readonly onPointerUp: (e: PointerEvent) => void;

  constructor(options: ShapeSelectionOptions) {
    this.opts = options;
    const doc = options.host.ownerDocument;
    this.disposeStyles = installStyleSheet(doc, STYLE_ID, CSS);

    this.root = doc.createElement('div');
    this.root.className = 'cbx-sel';

    this.outlines = doc.createElement('div');
    this.outlines.className = 'cbx-sel-outlines';
    this.root.appendChild(this.outlines);

    this.frame = doc.createElement('div');
    this.frame.className = 'cbx-sel-frame';
    this.frame.hidden = true;
    for (const h of HANDLES) {
      const grip = doc.createElement('div');
      grip.className = 'cbx-sel-handle';
      grip.dataset['cbxHandle'] = h;
      this.frame.appendChild(grip);
    }
    const knob = doc.createElement('div');
    knob.className = 'cbx-sel-rotate';
    knob.dataset['cbxHandle'] = 'rotate';
    this.frame.appendChild(knob);
    this.root.appendChild(this.frame);

    this.marquee = doc.createElement('div');
    this.marquee.className = 'cbx-sel-marquee';
    this.marquee.hidden = true;
    this.root.appendChild(this.marquee);

    options.host.appendChild(this.root);

    this.onPointerDown = (e) => this.pointerDown(e);
    this.onPointerMove = (e) => this.pointerMove(e);
    this.onPointerUp = (e) => this.pointerUp(e);
    options.host.addEventListener('pointerdown', this.onPointerDown);
    options.host.addEventListener('pointermove', this.onPointerMove);
    options.host.addEventListener('pointerup', this.onPointerUp);
    options.host.addEventListener('pointercancel', this.onPointerUp);
  }

  /** The shapes currently selected, in the order they were picked. */
  get selected(): readonly Shape[] {
    return this.shapes;
  }

  /**
   * Turn manipulation off while the user is typing: the overlay hides and every
   * press falls through to the contentEditable text underneath.
   */
  setEnabled(on: boolean): void {
    this.enabled = on;
    if (!on) this.cancelDrag();
    this.root.style.display = on ? '' : 'none';
  }

  /** Replace the selection. Pass `[]` to clear. */
  select(shapes: Shape[], notify = true): void {
    // Clearing the selection is also how you step back out of a group.
    if (shapes.length === 0) this.entered = [];
    this.shapes = shapes;
    this.refresh();
    if (notify) this.opts.onChange(this.shapes);
  }

  /**
   * Step out of the group the selection is inside, selecting the group itself —
   * what Escape does in PowerPoint. False when the selection is not in a group,
   * so the caller can clear it instead.
   */
  leaveGroup(): boolean {
    const parent = this.entered[this.entered.length - 1];
    if (!parent) return false;
    this.entered = this.entered.slice(0, -1);
    this.select([parent]);
    return true;
  }

  /** Redraw the handles — after a re-render, a scale change or a scroll. */
  refresh(): void {
    const view = this.opts.view();
    this.outlines.replaceChildren();
    if (!view.slideEl || this.shapes.length === 0) {
      this.frame.hidden = true;
      return;
    }

    if (this.shapes.length === 1) {
      const t = this.shapes[0]!.transform;
      if (!t) {
        this.frame.hidden = true;
        return;
      }
      this.frame.hidden = false;
      const box = frameRect(t, this.frameFor(this.shapes[0]!));
      place(this.frame, box.x, box.y, box.w, box.h, view);
      // The frame rotates with the shape, so a grip stays on the edge it grips.
      this.frame.style.transform = t.rot ? `rotate(${t.rot}deg)` : '';
      for (const grip of this.frame.querySelectorAll<HTMLElement>('[data-cbx-handle]')) {
        const h = grip.dataset['cbxHandle'];
        if (h && h !== 'rotate') grip.style.cursor = cursorFor(h as Handle, t.rot ?? 0);
      }
      return;
    }

    // Several shapes: one outline each, no grips. PowerPoint resizes a multiple
    // selection as a block; we do not, and showing grips would promise that.
    this.frame.hidden = true;
    for (const shape of this.shapes) {
      if (!shape.transform) continue;
      const b = frameRect(boundsOf(shape.transform), this.frameFor(shape));
      const box = this.opts.host.ownerDocument.createElement('div');
      box.className = 'cbx-sel-outline';
      place(box, b.x, b.y, b.w, b.h, view);
      this.outlines.appendChild(box);
    }
  }

  destroy(): void {
    const host = this.opts.host;
    host.removeEventListener('pointerdown', this.onPointerDown);
    host.removeEventListener('pointermove', this.onPointerMove);
    host.removeEventListener('pointerup', this.onPointerUp);
    host.removeEventListener('pointercancel', this.onPointerUp);
    this.root.remove();
    this.disposeStyles();
  }

  // ─── Pointer handling ──────────────────────────────────────────────────────

  private pointerDown(e: PointerEvent): void {
    if (!this.enabled || e.button !== 0) return;
    const view = this.opts.view();
    if (!view.slideEl) return;
    const target = e.target instanceof Element ? e.target : null;
    const from = this.toSlide(e, view);
    if (!from) return;

    const handle = target?.closest<HTMLElement>('[data-cbx-handle]')?.dataset['cbxHandle'];
    if (handle && this.shapes.length === 1) {
      this.begin(
        e,
        handle === 'rotate' ? { kind: 'rotate' } : { kind: 'resize', handle: handle as Handle },
        from,
        false,
      );
      return;
    }

    const chain = this.chainAt(target);
    // The shallowest shape the user is not already inside. With nothing
    // entered that is the outermost one, so a click on a shape within a group
    // selects the group — PowerPoint's rule.
    const picked = chain.findIndex((s) => !this.entered.includes(s));
    const shape = picked === -1 ? undefined : chain[picked];
    if (!shape || picked === -1) {
      // Empty slide: clear (which also leaves any group) and rubber-band.
      if (this.shapes.length) this.select([]);
      else this.entered = [];
      this.begin(e, { kind: 'marquee' }, from, false, [], -1);
      return;
    }
    // Picking a shape means we are inside exactly its ancestors — which also
    // steps back out of a group when the press lands outside it.
    this.entered = chain.slice(0, picked);

    const additive = e.shiftKey || e.metaKey || e.ctrlKey;
    if (additive) {
      const next = this.shapes.includes(shape)
        ? this.shapes.filter((s) => s !== shape)
        : [...this.shapes, shape];
      this.select(next);
      // Toggling membership is not a grab; a drag would move a shape the user
      // was only trying to add.
      return;
    }
    // Clicking a shape that is already the entire selection is the gesture that
    // opens its text — or steps into it, if it is a group; clicking an
    // unselected one only selects it.
    const activates = this.shapes.length === 1 && this.shapes[0] === shape;
    if (!activates) this.select([shape]);
    this.begin(e, { kind: 'move' }, from, activates, chain, picked);
  }

  /**
   * The shapes under an element, outermost first. A group's children are marked
   * just like top-level shapes, so this is the ancestor chain of marked
   * elements — what makes reaching inside a group possible at all.
   */
  private chainAt(target: Element | null): Shape[] {
    const out: Shape[] = [];
    let el = target?.closest<HTMLElement>(`[${EDIT_ATTR.shape}]`) ?? null;
    while (el) {
      const shape = this.opts.resolve(el.getAttribute(EDIT_ATTR.shape) ?? '');
      if (shape) out.unshift(shape);
      el = el.parentElement?.closest<HTMLElement>(`[${EDIT_ATTR.shape}]`) ?? null;
    }
    return out;
  }

  /** The space a shape's box is stated in. */
  private frameFor(shape: Shape): ShapeFrame {
    return this.opts.frameOf?.(shape) ?? SLIDE_FRAME;
  }

  private begin(
    e: PointerEvent,
    drag: Drag,
    from: { x: number; y: number },
    activates: boolean,
    chain: Shape[] = [],
    picked = -1,
  ): void {
    const subjects: DragState['subjects'] = [];
    if (drag.kind !== 'marquee') {
      for (const shape of this.shapes) {
        const el = this.elementFor(shape);
        if (el && shape.transform) {
          subjects.push({ shape, el, start: { ...shape.transform }, frame: this.frameFor(shape) });
        }
      }
      if (subjects.length === 0) return;
    }
    this.state = { drag, from, subjects, live: false, activates, chain, picked };
    this.opts.host.setPointerCapture(e.pointerId);
    // Stops the browser starting its own text/image drag. It also suppresses
    // the compatibility mouse events — click and dblclick never fire — which is
    // why entering a text box is driven from pointerup here rather than from a
    // dblclick listener.
    e.preventDefault();
  }

  private pointerMove(e: PointerEvent): void {
    const state = this.state;
    if (!state) return;
    const view = this.opts.view();
    const at = this.toSlide(e, view);
    if (!at) return;

    const dx = at.x - state.from.x;
    const dy = at.y - state.from.y;
    if (!state.live) {
      if (Math.hypot(dx, dy) * view.scale < DRAG_THRESHOLD) return;
      state.live = true;
    }

    if (state.drag.kind === 'marquee') {
      this.marquee.hidden = false;
      place(
        this.marquee,
        Math.min(state.from.x, at.x),
        Math.min(state.from.y, at.y),
        Math.abs(dx),
        Math.abs(dy),
        view,
      );
      return;
    }

    for (const s of state.subjects) {
      applyPreview(s.el, this.transformFor(state, s, dx, dy, at, e));
    }
    // Keep the handles on the shape while it moves; a single subject is the
    // only case that shows them.
    const first = state.subjects[0];
    if (first && state.subjects.length === 1) {
      const t = this.transformFor(state, first, dx, dy, at, e);
      const box = frameRect(t, first.frame);
      place(this.frame, box.x, box.y, box.w, box.h, view);
      this.frame.style.transform = t.rot ? `rotate(${t.rot}deg)` : '';
    }
  }

  private pointerUp(e: PointerEvent): void {
    const state = this.state;
    this.state = null;
    if (!state) return;
    if (this.opts.host.hasPointerCapture(e.pointerId)) {
      this.opts.host.releasePointerCapture(e.pointerId);
    }
    this.marquee.hidden = true;
    if (!state.live) {
      // A click, not a drag.
      const shape = state.subjects[0]?.shape;
      if (!state.activates || !shape) return;
      // Clicking an already-selected group steps inside it and selects the
      // child that was under the pointer, the way a second click in PowerPoint
      // does. A rotated or mirrored group is kept as the unit instead: the
      // handles for a child inside one would not line up with what is drawn.
      const inner = state.chain[state.picked + 1];
      if (inner && !this.frameFor(inner).turned) {
        this.entered = state.chain.slice(0, state.picked + 1);
        this.select([inner]);
        return;
      }
      this.opts.onActivate(shape, { x: e.clientX, y: e.clientY });
      return;
    }

    const view = this.opts.view();
    const at = this.toSlide(e, view);
    if (!at) {
      this.refresh();
      return;
    }

    if (state.drag.kind === 'marquee') {
      this.select(this.shapesWithin(state.from, at));
      return;
    }

    const dx = at.x - state.from.x;
    const dy = at.y - state.from.y;
    const edits = state.subjects.map((s) => ({
      shape: s.shape,
      transform: this.transformFor(state, s, dx, dy, at, e),
    }));
    // Every subject's element is about to be replaced by a re-render, but a
    // caller that declines the edit must not be left with a preview stuck to
    // the old geometry.
    for (const s of state.subjects) applyPreview(s.el, s.start);
    this.opts.onCommit(edits);
  }

  private cancelDrag(): void {
    const state = this.state;
    this.state = null;
    this.marquee.hidden = true;
    if (!state) return;
    for (const s of state.subjects) applyPreview(s.el, s.start);
  }

  /**
   * The transform a drag has reached, for one subject.
   *
   * The drag is measured on the slide but the box is written in the shape's own
   * space, so both the delta and the pointer are taken back through the
   * subject's frame first — for a top-level shape that is the identity, for one
   * inside a group it undoes the group's scale.
   */
  private transformFor(
    state: DragState,
    subject: DragState['subjects'][number],
    dx: number,
    dy: number,
    at: { x: number; y: number },
    e: PointerEvent | MouseEvent,
  ): Transform {
    const start = subject.start;
    const d = unframeDelta(subject.frame, dx, dy);
    switch (state.drag.kind) {
      case 'move':
        // Shift constrains to the dominant axis, as it does in PowerPoint.
        return e.shiftKey
          ? Math.abs(dx) > Math.abs(dy)
            ? moveBox(start, d.x, 0)
            : moveBox(start, 0, d.y)
          : moveBox(start, d.x, d.y);
      case 'resize':
        return resizeBox(start, state.drag.handle, d.x, d.y, {
          aspect: e.shiftKey,
          fromCenter: e.altKey,
          min: MIN_SIZE,
        });
      case 'rotate': {
        const p = unframePoint(subject.frame, at.x, at.y);
        return rotateBox(start, p.x, p.y, e.shiftKey);
      }
      default:
        return start;
    }
  }

  /** Shapes wholly inside the rubber band — PowerPoint's enclosure rule. */
  private shapesWithin(a: { x: number; y: number }, b: { x: number; y: number }): Shape[] {
    const view = this.opts.view();
    if (!view.slideEl) return [];
    const l = Math.min(a.x, b.x);
    const t = Math.min(a.y, b.y);
    const r = Math.max(a.x, b.x);
    const bot = Math.max(a.y, b.y);
    const out: Shape[] = [];
    for (const el of view.slideEl.querySelectorAll<HTMLElement>(`[${EDIT_ATTR.shape}]`)) {
      const chain = this.chainAt(el);
      const shape = chain[chain.length - 1];
      if (!shape?.transform) continue;
      // Rubber-banding works at whatever level the user is at: the slide's own
      // shapes normally, a group's children once inside it. Otherwise a marquee
      // over a group would select the group and everything in it at once.
      if (chain.length !== this.entered.length + 1) continue;
      const box = frameRect(boundsOf(shape.transform), this.frameFor(shape));
      if (box.x >= l && box.y >= t && box.x + box.w <= r && box.y + box.h <= bot) out.push(shape);
    }
    return out;
  }

  /** The rendered element for a shape, found by the key the renderer stamped. */
  private elementFor(shape: Shape): HTMLElement | null {
    const slideEl = this.opts.view().slideEl;
    if (!slideEl) return null;
    for (const el of slideEl.querySelectorAll<HTMLElement>(`[${EDIT_ATTR.shape}]`)) {
      if (this.opts.resolve(el.getAttribute(EDIT_ATTR.shape) ?? '') === shape) return el;
    }
    return null;
  }

  /** Pointer position in the slide's own coordinate space. */
  private toSlide(e: PointerEvent | MouseEvent, view: SlideView): { x: number; y: number } | null {
    if (!view.slideEl || view.scale <= 0) return null;
    // The slide's client rect already has the scale in it, so dividing by the
    // scale is all that is needed — and it stays correct under page scroll.
    const rect = view.slideEl.getBoundingClientRect();
    return { x: (e.clientX - rect.left) / view.scale, y: (e.clientY - rect.top) / view.scale };
  }
}

/** The eight grips, in the order they are drawn. */
const HANDLES: Handle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];

/** Position an overlay element from a slide-space box. */
function place(
  el: HTMLElement,
  x: number,
  y: number,
  w: number,
  h: number,
  view: SlideView,
): void {
  el.style.left = `${view.left + x * view.scale}px`;
  el.style.top = `${view.top + y * view.scale}px`;
  el.style.width = `${w * view.scale}px`;
  el.style.height = `${h * view.scale}px`;
}

/**
 * Show a transform on the rendered shape itself, so a drag reads as moving the
 * shape rather than an empty rectangle. This mirrors what `positioned()` writes
 * at render time; the re-render after the commit overwrites all of it.
 */
function applyPreview(el: HTMLElement, t: Transform): void {
  el.style.left = `${t.x}px`;
  el.style.top = `${t.y}px`;
  el.style.width = `${t.w}px`;
  el.style.height = `${t.h}px`;
  const parts: string[] = [];
  if (t.rot) parts.push(`rotate(${t.rot}deg)`);
  if (t.flipH || t.flipV) parts.push(`scale(${t.flipH ? -1 : 1}, ${t.flipV ? -1 : 1})`);
  el.style.transform = parts.join(' ');
  el.style.transformOrigin = 'center';
}

/** Resize cursors, indexed by eighth-turn clockwise from north. */
const CURSORS = ['ns-resize', 'nesw-resize', 'ew-resize', 'nwse-resize'];
const HANDLE_ANGLE: Record<Handle, number> = {
  n: 0, ne: 45, e: 90, se: 135, s: 180, sw: 225, w: 270, nw: 315,
};

/** The cursor for a grip, turned with the shape so it points across the edge. */
function cursorFor(handle: Handle, rot: number): string {
  const eighth = Math.round((HANDLE_ANGLE[handle] + rot) / 45);
  return CURSORS[((eighth % 4) + 4) % 4]!;
}

const CSS = `
[data-cbx-shape]{cursor:move;}
.cbx-sel{position:absolute;inset:0;pointer-events:none;z-index:2;}
.cbx-sel-frame,.cbx-sel-outline,.cbx-sel-marquee{position:absolute;box-sizing:border-box;}
.cbx-sel-frame{outline:1px solid #0d6efd;transform-origin:center;}
.cbx-sel-outline{outline:1px solid rgba(13,110,253,.75);}
.cbx-sel-marquee{border:1px solid #0d6efd;background:rgba(13,110,253,.12);}
.cbx-sel-handle{position:absolute;width:8px;height:8px;margin:-4px;border:1px solid #0d6efd;
  border-radius:50%;background:#fff;pointer-events:auto;}
.cbx-sel-handle[data-cbx-handle="nw"]{left:0;top:0;}
.cbx-sel-handle[data-cbx-handle="n"]{left:50%;top:0;}
.cbx-sel-handle[data-cbx-handle="ne"]{left:100%;top:0;}
.cbx-sel-handle[data-cbx-handle="e"]{left:100%;top:50%;}
.cbx-sel-handle[data-cbx-handle="se"]{left:100%;top:100%;}
.cbx-sel-handle[data-cbx-handle="s"]{left:50%;top:100%;}
.cbx-sel-handle[data-cbx-handle="sw"]{left:0;top:100%;}
.cbx-sel-handle[data-cbx-handle="w"]{left:0;top:50%;}
.cbx-sel-rotate{position:absolute;left:50%;top:-22px;width:10px;height:10px;margin:-5px;
  border:1px solid #0d6efd;border-radius:50%;background:#fff;pointer-events:auto;cursor:grab;}
.cbx-sel-rotate::before{content:"";position:absolute;left:50%;top:9px;width:1px;height:13px;
  background:#0d6efd;}
`;
