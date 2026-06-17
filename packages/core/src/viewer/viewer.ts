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
import type { Deck } from '../parse/deck.js';
import { renderSlide } from '../render/dom.js';
import { installDeckFonts, type FontInstallation } from '../render/fonts.js';

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

  constructor(deck: Deck, container: HTMLElement, options: ViewerOptions = {}) {
    this.deck = deck;
    this.container = container;
    this.fit = options.fit ?? 'contain';
    if (options.onChange) this.onChange = options.onChange;

    container.style.position = 'relative';

    // The holder fills the container; the slide is positioned/scaled within it.
    this.holder = document.createElement('div');
    this.holder.style.position = 'relative';
    this.holder.style.width = '100%';
    this.holder.style.height = this.fit === 'contain' ? '100%' : 'auto';
    this.holder.style.margin = '0 auto';
    container.appendChild(this.holder);

    if (options.keyboard !== false) this.enableKeyboard();
    this.resizeObserver = new ResizeObserver(() => this.applyScale());
    this.resizeObserver.observe(container);

    // Install embedded fonts, then re-render so text uses the real font.
    this.fonts = installDeckFonts(deck);
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

    const el = renderSlide(slide, this.deck.size, { imageUrl: (p) => this.deck.imageUrl(p) });
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

  destroy(): void {
    this.resizeObserver?.disconnect();
    this.fonts.dispose();
    if (this.keyHandler) this.container.removeEventListener('keydown', this.keyHandler);
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
