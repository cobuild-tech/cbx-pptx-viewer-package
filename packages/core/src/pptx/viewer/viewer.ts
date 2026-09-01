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
import { installDeckFonts, type FontInstallation } from '../render/fonts.js';
import { installWebFonts, type WebFontOptions } from '../render/webfonts.js';
import { EditContext } from '../edit/context.js';
import { EditSession } from '../edit/session.js';
import { reconcileTextBody } from '../edit/reconcile.js';
import { applyFormatToSelection, bodyElementOf, formatAtSelection } from '../../oxml/edit/selection.js';
import { installEditStyles, type TextBoxOutline } from '../../oxml/edit/styles.js';
import { readRunFormat } from '../edit/format.js';
import type { RunFormat } from '../../oxml/edit/format.js';
import { EDIT_ATTR } from '../text/render.js';
import type { TextBody } from '../model.js';
import { Filmstrip } from './filmstrip.js';

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
   * How editable text boxes are outlined, so the user can see what can be
   * edited. `'hover'` (default) reveals a box under the pointer, `'always'`
   * shows every editable box at once, `'none'` shows only the focused one.
   * Ignored unless `editable`.
   */
  textBoxOutline?: TextBoxOutline;
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

  constructor(deck: Deck, container: HTMLElement, options: ViewerOptions = {}) {
    this.deck = deck;
    this.container = container;
    this.fit = options.fit ?? 'contain';
    if (options.onChange) this.onChange = options.onChange;
    this.editable = options.editable ?? false;
    if (options.onEdit) this.onEdit = options.onEdit;
    if (options.onSelectionChange) this.onSelectionChange = options.onSelectionChange;
    if (this.editable) {
      this.editCtx = new EditContext(deck, options.startIndex ?? 0);
      this.session = new EditSession(deck, { onChange: (i) => this.onEdit?.(i) });
      this.disposeStyles = installEditStyles(
        container.ownerDocument,
        options.textBoxOutline ?? 'hover',
      );
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
      return;
    }

    // contain: largest scale that fits both width and height, then centre.
    const ch = this.stage.clientHeight || hPx;
    const scale = Math.min(cw / wPx, ch / hPx);
    this.slideEl.style.transform = `scale(${scale})`;
    this.slideEl.style.left = `${Math.max(0, (cw - wPx * scale) / 2)}px`;
    this.slideEl.style.top = `${Math.max(0, (ch - hPx * scale) / 2)}px`;
  }

  private enableKeyboard(): void {
    if (this.container.tabIndex < 0) this.container.tabIndex = 0;
    this.keyHandler = (e: KeyboardEvent) => {
      // Arrows, space and Home/End all mean something else inside text. Without
      // this an editor would jump slides mid-word.
      if (this.isEditingTarget(e.target)) return;
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
      this.commitBody(from);
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

  /** Change how editable text boxes are outlined. */
  setTextBoxOutline(mode: TextBoxOutline): void {
    if (!this.editable) return;
    this.disposeStyles?.();
    this.disposeStyles = installEditStyles(this.container.ownerDocument, mode);
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
