/**
 * Viewer controller: mounts a {@link Deck} into a container and renders one
 * slide at a time.
 *
 * Sizing: by default the slide is fit to the container like PowerPoint's
 * slideshow — uniformly scaled to the largest size that fits both the
 * container's width and height, then centred (`fit: 'contain'`). Pass
 * `fit: 'width'` for the embedded-document style that fills the width and lets
 * the page scroll vertically.
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
  private readonly holder: HTMLDivElement;
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

    // The holder fills the container; the slide is positioned/scaled within it.
    this.holder = document.createElement('div');
    this.holder.style.position = 'relative';
    this.holder.style.width = '100%';
    this.holder.style.height = this.fit === 'contain' ? '100%' : 'auto';
    this.holder.style.margin = '0 auto';
    container.appendChild(this.holder);

    if (options.keyboard !== false) this.enableKeyboard();
    if (this.editable) this.enableEditing();
    this.resizeObserver = new ResizeObserver(() => this.applyScale());
    this.resizeObserver.observe(container);

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
      if (this.slideEl) this.goTo(this.index);
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
    const cw = this.container.clientWidth || wPx;

    if (this.fit === 'width') {
      const scale = cw / wPx;
      this.slideEl.style.transform = `scale(${scale})`;
      this.slideEl.style.left = '0px';
      this.slideEl.style.top = '0px';
      this.holder.style.height = `${hPx * scale}px`;
      return;
    }

    // contain: largest scale that fits both width and height, then centre.
    const ch = this.container.clientHeight || hPx;
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

  undo(): void {
    const slide = this.session?.undo();
    if (slide) this.goTo(slide.index);
  }

  redo(): void {
    const slide = this.session?.redo();
    if (slide) this.goTo(slide.index);
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
  }
}

export function createViewer(
  deck: Deck,
  container: HTMLElement,
  options?: ViewerOptions,
): Viewer {
  return new Viewer(deck, container, options);
}
