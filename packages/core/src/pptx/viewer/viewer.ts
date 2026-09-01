/**
 * Viewer controller: mounts a {@link Deck} into a container and renders one
 * slide at a time.
 *
 * Layout: PowerPoint's — a scrolling thumbnail rail down the left, the current
 * slide on the stage to its right. Pass `filmstrip: false` for the bare stage.
 *
 * Sizing: by default the slide is fit to the stage like PowerPoint's
 * slideshow — uniformly scaled to the largest size that fits both the stage's
 * width and height, then centred (`fit: 'contain'`). Pass `fit: 'width'` for
 * the embedded-document style that fills the width and lets the page scroll
 * vertically.
 *
 * On load it installs any fonts embedded in the deck (FontFace) and re-renders
 * once they're ready, so text reflows in the genuine font.
 */
import type { Deck } from '../deck/deck.js';
import { renderSlide } from '../render/dom.js';
import { SLIDE_FRAME } from '../render/primitives.js';
import { installDeckFonts, type FontInstallation } from '../render/fonts.js';
import { installWebFonts, type WebFontOptions } from '../render/webfonts.js';
import { EditContext } from '../edit/context.js';
import { EditSession } from '../edit/session.js';
import { reconcileTextBody } from '../edit/reconcile.js';
import { applyFormatToSelection, bodyElementOf, formatAtSelection } from '../../oxml/edit/selection.js';
import { installEditStyles } from '../../oxml/edit/styles.js';
import { readRunFormat } from '../edit/format.js';
import type { RunFormat } from '../../oxml/edit/format.js';
import { EDIT_ATTR } from '../text/render.js';
import type { Shape, TextBody, Transform } from '../model.js';
import type { XmlNode } from '../../oxml/xml.js';
import type { ZOrderMove } from '../edit/shapeOps.js';
import { moveBox } from '../edit/geometry.js';
import { Filmstrip } from './filmstrip.js';
import { ShapeSelection } from './selection.js';

export interface ViewerOptions {
  startIndex?: number;
  /** Enable arrow-key / space navigation on the container. Default true. */
  keyboard?: boolean;
  /**
   * How the slide is scaled into the container. `'contain'` (default) fits the
   * whole slide to the container and centres it; `'width'` fills the width and
   * lets height follow the aspect ratio (the page scrolls).
   */
  fit?: 'contain' | 'width';
  /** Called whenever the current slide index changes. */
  onChange?: (index: number, count: number) => void;
  /**
   * Show the thumbnail rail down the left, PowerPoint-style. Default true; pass
   * `false` for a bare stage (an embed with its own navigation, say).
   */
  filmstrip?: boolean;
  /** Width of the thumbnail rail in CSS px. Default 200. */
  filmstripWidth?: number;
  /**
   * Fetch the deck's fonts (and metric-compatible substitutes for Office fonts)
   * from Google Fonts when they aren't available locally, so text wraps as it
   * does in PowerPoint. Pass `false` to disable network font loading, or an
   * object to customize the source.
   */
  webFonts?: boolean | WebFontOptions;
  /**
   * Make the slide's own text bodies editable in place. Text inherited from the
   * layout or master stays read-only. Off by default.
   */
  editable?: boolean;
  /** Called after a text edit is committed, or after undo/redo. */
  onEdit?: (slideIndex: number) => void;
  /** Called when the caret moves, with the formatting in effect there. */
  onSelectionChange?: (format: RunFormat) => void;
  /**
   * Select, move, resize, rotate, restack and delete the slide's own shapes.
   * Defaults to `editable`; pass `false` to keep text editing without direct
   * manipulation.
   */
  shapeEditing?: boolean;
  /** Called when the set of selected shapes changes. */
  onShapeSelectionChange?: (shapes: readonly Shape[]) => void;
}

export class Viewer {
  private readonly deck: Deck;
  private readonly container: HTMLElement;
  /**
   * The area the slide is scaled into. Its own element when the filmstrip is
   * showing (the rail is its flex sibling), otherwise the container itself —
   * so a viewer without a rail keeps exactly the DOM it always had.
   */
  private readonly stage: HTMLElement;
  private readonly holder: HTMLDivElement;
  private strip: Filmstrip | null = null;
  private readonly onChange: ViewerOptions['onChange'];
  private readonly fit: 'contain' | 'width';
  private index = 0;
  private slideEl: HTMLElement | null = null;
  private resizeObserver: ResizeObserver | null = null;
  private keyHandler: ((e: KeyboardEvent) => void) | null = null;
  private fonts: FontInstallation;
  private readonly editable: boolean;
  private editCtx: EditContext | null = null;
  private session: EditSession | null = null;
  private readonly onEdit: ViewerOptions['onEdit'];
  private readonly onSelectionChange: ViewerOptions['onSelectionChange'];
  private focusOutHandler: ((e: FocusEvent) => void) | null = null;
  private selectionHandler: (() => void) | null = null;
  private disposeStyles: (() => void) | null = null;
  private selection: ShapeSelection | null = null;
  private readonly onShapeSelectionChange: ViewerOptions['onShapeSelectionChange'];
  /** Current slide placement, mirrored for the selection overlay's mapping. */
  private scale = 1;
  private offsetX = 0;
  private offsetY = 0;

  constructor(deck: Deck, container: HTMLElement, options: ViewerOptions = {}) {
    this.deck = deck;
    this.container = container;
    this.fit = options.fit ?? 'contain';
    if (options.onChange) this.onChange = options.onChange;
    this.editable = options.editable ?? false;
    if (options.onEdit) this.onEdit = options.onEdit;
    if (options.onSelectionChange) this.onSelectionChange = options.onSelectionChange;
    if (options.onShapeSelectionChange) this.onShapeSelectionChange = options.onShapeSelectionChange;
    if (this.editable) {
      this.editCtx = new EditContext(deck, options.startIndex ?? 0);
      this.session = new EditSession(deck, { onChange: (i) => this.onEdit?.(i) });
      // PowerPoint outlines nothing until you select something, and neither do
      // we: selection handles say what is editable, so the only text-box
      // decoration left is the focus ring on the box being typed in.
      this.disposeStyles = installEditStyles(container.ownerDocument, 'none');
    }

    container.style.position = 'relative';

    if (options.filmstrip !== false) {
      container.style.display = 'flex';
      container.style.flexDirection = 'row';
      container.style.alignItems = 'stretch';
      this.strip = new Filmstrip(container.ownerDocument, {
        slides: () => this.deck.slides,
        size: deck.size,
        imageUrl: (p) => this.deck.imageUrl(p),
        onSelect: (i) => this.goTo(i),
        ...(options.filmstripWidth !== undefined ? { width: options.filmstripWidth } : {}),
      });
      container.appendChild(this.strip.el);

      const stage = document.createElement('div');
      stage.style.position = 'relative';
      stage.style.flex = '1 1 auto';
      // Without this a wide slide would push the flex item past the container
      // instead of scaling down into it, and the rail would be squeezed out.
      stage.style.minWidth = '0';
      stage.style.height = this.fit === 'contain' ? '100%' : 'auto';
      container.appendChild(stage);
      this.stage = stage;
    } else {
      this.stage = container;
    }

    // The holder fills the stage; the slide is positioned/scaled within it.
    this.holder = document.createElement('div');
    this.holder.style.position = 'relative';
    this.holder.style.width = '100%';
    this.holder.style.height = this.fit === 'contain' ? '100%' : 'auto';
    this.holder.style.margin = '0 auto';
    this.stage.appendChild(this.holder);

    if (options.keyboard !== false) this.enableKeyboard();
    if (this.editable) this.enableEditing();
    if (this.editable && (options.shapeEditing ?? true)) this.enableShapeEditing();
    this.resizeObserver = new ResizeObserver(() => this.applyScale());
    this.resizeObserver.observe(this.stage);

    // Install the deck's fonts — embedded faces plus, for fonts not available
    // locally, substitutes fetched from Google Fonts — then re-render so text
    // reflows with correct metrics.
    const embedded = installDeckFonts(deck);
    const web =
      options.webFonts === false
        ? { ready: Promise.resolve(), dispose() {} }
        : installWebFonts(deck, typeof options.webFonts === 'object' ? options.webFonts : {});
    this.fonts = {
      ready: Promise.all([embedded.ready, web.ready]).then(() => undefined),
      dispose() {
        embedded.dispose();
        web.dispose();
      },
    };
    this.fonts.ready.then(() => {
      if (!this.slideEl) return;
      // Text reflows once the real metrics land, so the thumbnails are as stale
      // as the stage is.
      this.strip?.rebuild();
      this.goTo(this.index);
    });

    this.goTo(options.startIndex ?? 0);
  }

  get count(): number {
    return this.deck.slides.length;
  }

  get currentIndex(): number {
    return this.index;
  }

  goTo(index: number): void {
    const clamped = Math.max(0, Math.min(this.deck.slides.length - 1, index));
    // Model objects belong to one slide, so a selection cannot survive leaving
    // it — and neither can a half-typed text box.
    if (clamped !== this.index) {
      this.commitActive();
      this.selection?.select([]);
    }
    this.index = clamped;
    const slide = this.deck.slides[clamped];
    if (!slide) return;

    // Keys identify model objects for this render only, so they are reissued
    // whenever the slide is (re)rendered.
    this.editCtx?.retarget(clamped);
    const el = renderSlide(slide, this.deck.size, {
      imageUrl: (p) => this.deck.imageUrl(p),
      ...(this.editCtx ? { edit: this.editCtx } : {}),
    });
    el.style.transformOrigin = 'top left';
    el.style.position = 'absolute';
    if (this.slideEl) this.slideEl.remove();
    this.slideEl = el;
    this.holder.appendChild(el);
    this.applyScale();
    this.selection?.refresh();
    this.strip?.setActive(this.index);
    this.onChange?.(this.index, this.count);
  }

  next(): void {
    this.goTo(this.index + 1);
  }

  prev(): void {
    this.goTo(this.index - 1);
  }

  /**
   * Scale the slide into the container. `'contain'` fits the whole slide to both
   * dimensions and centres it; `'width'` fills the width and lets height follow.
   */
  private applyScale(): void {
    if (!this.slideEl) return;
    const { wPx, hPx } = this.deck.size;
    const cw = this.stage.clientWidth || wPx;

    if (this.fit === 'width') {
      const scale = cw / wPx;
      this.slideEl.style.transform = `scale(${scale})`;
      this.slideEl.style.left = '0px';
      this.slideEl.style.top = '0px';
      this.holder.style.height = `${hPx * scale}px`;
      this.place(scale, 0, 0);
      return;
    }

    // contain: largest scale that fits both width and height, then centre.
    const ch = this.stage.clientHeight || hPx;
    const scale = Math.min(cw / wPx, ch / hPx);
    const left = Math.max(0, (cw - wPx * scale) / 2);
    const top = Math.max(0, (ch - hPx * scale) / 2);
    this.slideEl.style.transform = `scale(${scale})`;
    this.slideEl.style.left = `${left}px`;
    this.slideEl.style.top = `${top}px`;
    this.place(scale, left, top);
  }

  /** Record where the slide landed and move the handles with it. */
  private place(scale: number, left: number, top: number): void {
    this.scale = scale;
    this.offsetX = left;
    this.offsetY = top;
    this.selection?.refresh();
  }

  private enableKeyboard(): void {
    if (this.container.tabIndex < 0) this.container.tabIndex = 0;
    this.keyHandler = (e: KeyboardEvent) => {
      // Escape is the way out of a text box, so it has to be read before the
      // guard that hands every other key to the text.
      if (e.key === 'Escape' && this.isEditingText) {
        e.preventDefault();
        this.exitTextEditing();
        return;
      }
      // Arrows, space and Home/End all mean something else inside text. Without
      // this an editor would jump slides mid-word.
      if (this.isEditingTarget(e.target)) return;
      if (this.handleShapeKey(e)) return;
      // Delete removes a slide only with the rail focused — the same scoping
      // PowerPoint uses, so the key can't destroy a slide from the stage.
      if ((e.key === 'Delete' || e.key === 'Backspace') && this.isFilmstripTarget(e.target)) {
        e.preventDefault();
        this.deleteSlide();
        return;
      }
      switch (e.key) {
        case 'ArrowRight':
        case 'ArrowDown':
        case 'PageDown':
        case ' ':
          e.preventDefault();
          this.next();
          break;
        case 'ArrowLeft':
        case 'ArrowUp':
        case 'PageUp':
          e.preventDefault();
          this.prev();
          break;
        case 'Home':
          e.preventDefault();
          this.goTo(0);
          break;
        case 'End':
          e.preventDefault();
          this.goTo(this.count - 1);
          break;
      }
    };
    this.container.addEventListener('keydown', this.keyHandler);
  }


  /**
   * Keys that act on the selection. Returns true when the key was consumed, so
   * navigation never fires as well — an arrow key nudges a selected shape
   * rather than jumping to the next slide.
   */
  private handleShapeKey(e: KeyboardEvent): boolean {
    const sel = this.selection;
    if (!sel) return false;

    // Tab walks the slide's shapes whether or not anything is selected yet.
    if (e.key === 'Tab' && !this.isFilmstripTarget(e.target)) {
      e.preventDefault();
      this.cycleSelection(e.shiftKey ? -1 : 1);
      return true;
    }
    if (sel.selected.length === 0) return false;

    const step = e.shiftKey ? 10 : 1;
    switch (e.key) {
      case 'Escape':
        e.preventDefault();
        // Inside a group, Escape selects the group rather than deselecting —
        // one level per press, as in PowerPoint.
        if (!sel.leaveGroup()) sel.select([]);
        return true;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        this.deleteSelectedShapes();
        return true;
      case 'Enter':
      case 'F2': {
        const shape = sel.selected[0];
        if (!shape) return false;
        e.preventDefault();
        this.editText(shape);
        return true;
      }
      case 'ArrowLeft':
        e.preventDefault();
        this.nudgeSelection(-step, 0);
        return true;
      case 'ArrowRight':
        e.preventDefault();
        this.nudgeSelection(step, 0);
        return true;
      case 'ArrowUp':
        e.preventDefault();
        this.nudgeSelection(0, -step);
        return true;
      case 'ArrowDown':
        e.preventDefault();
        this.nudgeSelection(0, step);
        return true;
      case ']':
      case '[': {
        // PowerPoint's restack shortcuts: Ctrl+] forward, adding Shift to send
        // it the whole way.
        if (!e.ctrlKey && !e.metaKey) return false;
        e.preventDefault();
        const forward = e.key === ']';
        this.reorderSelectedShape(
          e.shiftKey ? (forward ? 'front' : 'back') : forward ? 'forward' : 'backward',
        );
        return true;
      }
      default:
        return false;
    }
  }

  /**
   * Every shape on the slide the user could select, groups and their children
   * alike, in document order. A group child is addressable in its own right
   * once the user has stepped inside the group, so anything that re-finds a
   * selection after a rebuild has to look at the whole tree, not just the top
   * level.
   */
  private ownShapes(): Shape[] {
    const out: Shape[] = [];
    const walk = (shapes: readonly Shape[]): void => {
      for (const s of shapes) {
        out.push(s);
        if (s.kind === 'group') walk(s.children);
      }
    };
    walk(this.deck.slides[this.index]?.shapes ?? []);
    return out;
  }

  /** Step the selection through the slide's top-level shapes in z-order. */
  private cycleSelection(delta: number): void {
    const shapes = (this.deck.slides[this.index]?.shapes ?? []).filter(
      (s) => this.deck.isEditable(this.index, s) && s.transform,
    );
    if (shapes.length === 0) return;
    const current = this.selectedShapes[0];
    const at = current ? shapes.indexOf(current) : -1;
    const next = shapes[(((at + delta) % shapes.length) + shapes.length) % shapes.length];
    if (next) this.selection?.select([next]);
  }

  /** True if the event target sits inside the thumbnail rail. */
  private isFilmstripTarget(target: EventTarget | null): boolean {
    return !!this.strip && target instanceof Node && this.strip.el.contains(target);
  }

  /** True if the event target sits inside an editable text body. */
  private isEditingTarget(target: EventTarget | null): boolean {
    if (!this.editable || !this.slideEl || !(target instanceof Node)) return false;
    return !!bodyElementOf(target, this.slideEl);
  }

  private enableEditing(): void {
    // Commit when focus leaves a text body — including when it moves to another
    // body on the same slide, which focusout reports and blur does not.
    this.focusOutHandler = (e: FocusEvent) => {
      if (!this.slideEl) return;
      const from = bodyElementOf(e.target as Node | null, this.slideEl);
      if (!from) return;
      const to = e.relatedTarget instanceof Node ? bodyElementOf(e.relatedTarget, this.slideEl) : null;
      if (to === from) return;
      // Drop the text-editing state *before* committing, so the re-render the
      // commit does is already the read-only one — one render, not two.
      const wasEditing = !!this.editCtx?.editingBody;
      this.editCtx?.setTextEditing(null);
      const committed = this.commitBody(from);
      if (wasEditing) {
        this.selection?.setEnabled(true);
        if (!committed) this.goTo(this.index);
      }
    };
    this.container.addEventListener('focusout', this.focusOutHandler, true);

    if (this.onSelectionChange) {
      this.selectionHandler = () => {
        if (!this.slideEl || !this.editCtx) return;
        const sel = this.container.ownerDocument.getSelection();
        if (!sel || sel.rangeCount === 0) return;
        if (!bodyElementOf(sel.getRangeAt(0).startContainer, this.slideEl)) return;
        this.onSelectionChange?.(formatAtSelection(this.slideEl, (k) => this.editCtx?.resolve(k), readRunFormat));
      };
      this.container.ownerDocument.addEventListener('selectionchange', this.selectionHandler);
    }
  }

  /** Read one edited text body back into the XML and re-render the slide. */
  private commitBody(bodyEl: HTMLElement): boolean {
    if (!this.editCtx || !this.session) return false;
    const body = this.editCtx.resolve(bodyEl.getAttribute(EDIT_ATTR.body)) as TextBody | undefined;
    if (!body) return false;

    const paras = reconcileTextBody(bodyEl, (k) => this.editCtx?.resolve(k));
    const slide = this.session.commitTextBody(this.index, body, paras);
    if (!slide) return false;
    // Re-render so the committed XML — not the browser's improvised markup — is
    // what the user is looking at.
    this.strip?.refresh(this.index);
    this.goTo(this.index);
    return true;
  }

  /** Commit whatever text body currently has focus, if any. */
  commitActive(): boolean {
    if (!this.editable || !this.slideEl) return false;
    const active = this.container.ownerDocument.activeElement;
    const bodyEl = active instanceof Node ? bodyElementOf(active, this.slideEl) : null;
    return bodyEl ? this.commitBody(bodyEl) : false;
  }

  /** Apply formatting to the current selection and commit it. */
  applyFormat(format: RunFormat): boolean {
    if (!this.editable || !this.slideEl) return false;
    const bodyEl = applyFormatToSelection(this.slideEl, format);
    return bodyEl ? this.commitBody(bodyEl) : false;
  }


  // ─── Shape selection and manipulation ──────────────────────────────────────

  private enableShapeEditing(): void {
    this.selection = new ShapeSelection({
      host: this.holder,
      view: () => ({
        slideEl: this.slideEl,
        scale: this.scale,
        left: this.offsetX,
        top: this.offsetY,
      }),
      resolve: (key) => this.editCtx?.resolve(key) as Shape | undefined,
      frameOf: (shape) => this.editCtx?.frameOf(shape) ?? SLIDE_FRAME,
      onChange: (shapes) => {
        // The press that made this selection called preventDefault (to stop the
        // browser's own drag), so focus never moved on its own — and the keys
        // that act on a selection are bound to the container. Without this,
        // Delete and the arrow keys would go to the page instead.
        if (shapes.length) this.container.focus();
        this.onShapeSelectionChange?.(shapes);
      },
      onCommit: (edits) => this.commitShapes(edits),
      onActivate: (shape, at) => this.editText(shape, at),
    });
  }

  /** The shapes currently selected on the stage. */
  get selectedShapes(): readonly Shape[] {
    return this.selection?.selected ?? [];
  }

  /** Select shapes on the current slide, or pass `[]` to clear. */
  selectShapes(shapes: Shape[]): void {
    this.selection?.select(shapes);
  }

  /** Write finished drags and re-render with the same shapes still selected. */
  private commitShapes(edits: Array<{ shape: Shape; transform: Transform }>): void {
    if (!this.session) return;
    this.withSelection(() => !!this.session?.commitShapeTransforms(this.index, edits));
  }

  /** Delete the selected shapes. Returns false if nothing was deleted. */
  deleteSelectedShapes(): boolean {
    const shapes = [...this.selectedShapes];
    if (!this.session || shapes.length === 0) return false;
    const slide = this.session.deleteShapes(this.index, shapes);
    if (!slide) return false;
    this.selection?.select([]);
    this.strip?.refresh(this.index);
    this.goTo(this.index);
    return true;
  }

  /** Move the selected shape through the z-order. */
  reorderSelectedShape(move: ZOrderMove): boolean {
    const shape = this.selectedShapes[0];
    if (!this.session || !shape || this.selectedShapes.length !== 1) return false;
    return this.withSelection(() => !!this.session?.reorderShape(this.index, shape, move));
  }

  /** Nudge the selection by a slide-space delta (arrow keys). */
  nudgeSelection(dx: number, dy: number): boolean {
    const shapes = this.selectedShapes.filter((s) => s.transform);
    if (shapes.length === 0) return false;
    this.commitShapes(
      shapes.map((shape) => ({ shape, transform: moveBox(shape.transform!, dx, dy) })),
    );
    return true;
  }

  /**
   * Run an edit that rebuilds the slide, then put the selection back.
   *
   * A commit re-parses the slide, so every model object the selection holds is
   * stale afterwards. The XML nodes behind them are not — a rebuild mutates the
   * cached tree rather than replacing it — so identity is carried across on the
   * source node, falling back to the shape id for the one case where the tree
   * really is new (an undo, which restores the part from a snapshot).
   */
  private withSelection(edit: () => boolean): boolean {
    const marks = this.selectedShapes.map((s) => this.markOf(s));
    if (!edit()) {
      // Nothing changed, but the preview geometry is still on screen.
      this.goTo(this.index);
      return false;
    }
    this.strip?.refresh(this.index);
    this.goTo(this.index);
    this.restoreSelection(marks);
    return true;
  }

  /** What identifies a selected shape across a rebuild of the slide. */
  private markOf(shape: Shape): { node: XmlNode | undefined; id: string | undefined } {
    return { node: this.deck.sourceOf(shape)?.node, id: shape.id };
  }

  private restoreSelection(marks: Array<{ node: XmlNode | undefined; id: string | undefined }>): void {
    if (!this.selection || marks.length === 0) return;
    const shapes = this.ownShapes();
    const found: Shape[] = [];
    for (const mark of marks) {
      const match =
        shapes.find((s) => mark.node && this.deck.sourceOf(s)?.node === mark.node) ??
        (mark.id !== undefined ? shapes.find((s) => s.id === mark.id) : undefined);
      if (match) found.push(match);
    }
    this.selection.select(found);
  }

  /**
   * Enter text editing for a shape: the body becomes contentEditable, the
   * handles go away, and the caret lands in the text. Does nothing for a shape
   * with no text of its own (a picture, or a title inherited from the layout).
   */
  editText(shape: Shape, at?: { x: number; y: number }): boolean {
    const body = this.bodyToEdit(shape, at);
    if (!this.editCtx || !body || !this.session?.isEditable(this.index, body)) return false;
    this.editCtx.setTextEditing(body);
    this.selection?.setEnabled(false);
    // Re-render so the box becomes contentEditable, then focus the box that now
    // is — the element the user clicked has just been replaced.
    this.goTo(this.index);
    const el = this.slideEl?.querySelector<HTMLElement>(
      `[${EDIT_ATTR.body}][contenteditable="true"]`,
    );
    if (!el) return false;
    el.focus();
    this.placeCaret(el, at);
    return true;
  }

  /**
   * Which text body a request to type in this shape means.
   *
   * A shape has one of its own. A table has one per cell, so the click decides
   * which — and with no click to go on (Enter/F2) it is the first cell that has
   * text. Anything else (a picture, a chart) has none.
   */
  private bodyToEdit(shape: Shape, at?: { x: number; y: number }): TextBody | undefined {
    if (shape.kind === 'shape') return shape.text;
    if (shape.kind !== 'frame' || !shape.table) return undefined;
    if (at) {
      const el = this.container.ownerDocument
        .elementFromPoint(at.x, at.y)
        ?.closest<HTMLElement>(`[${EDIT_ATTR.body}]`);
      const body = el
        ? (this.editCtx?.resolve(el.getAttribute(EDIT_ATTR.body)) as TextBody | undefined)
        : undefined;
      if (body) return body;
    }
    for (const row of shape.table.rows) {
      for (const cell of row) {
        if (cell?.text) return cell.text;
      }
    }
    return undefined;
  }

  /**
   * Drop the caret where the user clicked, falling back to selecting the whole
   * box when there is no click to go on (F2, Enter) or the browser cannot map
   * the point — landing mid-word is what makes this feel like typing in
   * PowerPoint rather than in a form field.
   */
  private placeCaret(el: HTMLElement, at?: { x: number; y: number }): void {
    const doc = this.container.ownerDocument;
    const sel = doc.getSelection();
    if (!sel) return;
    const range = at ? caretRangeAt(doc, at.x, at.y) : null;
    if (range && el.contains(range.startContainer)) {
      sel.removeAllRanges();
      sel.addRange(range);
      return;
    }
    sel.selectAllChildren(el);
    sel.collapseToEnd();
  }

  /**
   * Leave text editing and go back to selecting shapes, keeping whatever was
   * typed — Escape ends the edit in PowerPoint, it does not revert it.
   */
  exitTextEditing(): void {
    if (!this.editCtx?.editingBody) return;
    this.editCtx.setTextEditing(null);
    const committed = this.commitActive();
    this.selection?.setEnabled(true);
    if (!committed) this.goTo(this.index);
    this.container.focus();
  }

  /** True while a text body is open for typing. */
  get isEditingText(): boolean {
    return !!this.editCtx?.editingBody;
  }

  get canUndo(): boolean {
    return this.session?.canUndo ?? false;
  }

  get canRedo(): boolean {
    return this.session?.canRedo ?? false;
  }

  /**
   * True if the current slide (or `index`) can be deleted. A deck must keep at
   * least one slide.
   */
  canDeleteSlide(index = this.index): boolean {
    return this.session?.canDeleteSlide(index) ?? false;
  }

  /**
   * Delete a slide and show whichever slide takes its place (the last one, if
   * the deleted slide was last). Undoable. Returns false if nothing was deleted.
   */
  deleteSlide(index = this.index): boolean {
    if (!this.session) return false;
    // Anything half-typed in a text body belongs in the deck before the running
    // order changes underneath it.
    this.commitActive();
    if (!this.session.deleteSlide(index)) return false;
    this.selection?.select([]);
    this.editCtx?.reset();
    // Every slide after the deleted one shifted, so the rail is rebuilt whole
    // rather than patched.
    this.strip?.rebuild();
    this.goTo(Math.min(index, this.count - 1));
    return true;
  }

  undo(): void {
    if (!this.session?.canUndo) return;
    this.afterRestore(this.session.undo()?.index);
  }

  redo(): void {
    if (!this.session?.canRedo) return;
    this.afterRestore(this.session.redo()?.index);
  }

  /**
   * Land the viewer after an undo/redo. A change in slide count means the
   * running order moved, so the rail is rebuilt; otherwise one slide's content
   * changed and only its thumbnail is stale.
   */
  private afterRestore(landed: number | undefined): void {
    // Restoring a part re-parses it from a snapshot, so neither the model
    // objects nor the XML nodes the selection was holding still exist.
    this.selection?.select([]);
    const target = landed ?? Math.min(this.index, this.count - 1);
    if (this.strip && this.strip.count !== this.count) this.strip.rebuild();
    else this.strip?.refresh(target);
    this.goTo(target);
  }

  /** True if the deck has unsaved edits. */
  get hasEdits(): boolean {
    return this.session?.hasEdits ?? false;
  }

  /** Re-zip the deck, edits included, as a .pptx Blob. */
  exportBlob(): Blob {
    // Flush anything still being typed before packaging.
    this.commitActive();
    return this.deck.exportBlob();
  }

  destroy(): void {
    this.resizeObserver?.disconnect();
    this.fonts.dispose();
    if (this.keyHandler) this.container.removeEventListener('keydown', this.keyHandler);
    if (this.focusOutHandler) {
      this.container.removeEventListener('focusout', this.focusOutHandler, true);
    }
    if (this.selectionHandler) {
      this.container.ownerDocument.removeEventListener('selectionchange', this.selectionHandler);
    }
    this.disposeStyles?.();
    this.selection?.destroy();
    this.selection = null;
    if (this.slideEl) this.slideEl.remove();
    this.slideEl = null;
    this.strip?.destroy();
    this.strip = null;
    // The stage is ours only when a rail forced a two-column layout; otherwise
    // it is the caller's container and stays put.
    if (this.stage !== this.container) this.stage.remove();
  }
}

export function createViewer(
  deck: Deck,
  container: HTMLElement,
  options?: ViewerOptions,
): Viewer {
  return new Viewer(deck, container, options);
}

/**
 * The caret position under a client point. Browsers split on the spelling of
 * this — WebKit/Blink have `caretRangeFromPoint`, Gecko `caretPositionFromPoint`
 * — and jsdom has neither, so all three cases are handled here.
 */
function caretRangeAt(doc: Document, x: number, y: number): Range | null {
  const legacy = (doc as Document & { caretRangeFromPoint?: (x: number, y: number) => Range | null })
    .caretRangeFromPoint;
  if (typeof legacy === 'function') return legacy.call(doc, x, y);

  const standard = (
    doc as Document & {
      caretPositionFromPoint?: (
        x: number,
        y: number,
      ) => { offsetNode: Node; offset: number } | null;
    }
  ).caretPositionFromPoint;
  if (typeof standard !== 'function') return null;
  const pos = standard.call(doc, x, y);
  if (!pos) return null;
  const range = doc.createRange();
  range.setStart(pos.offsetNode, pos.offset);
  range.collapse(true);
  return range;
}
