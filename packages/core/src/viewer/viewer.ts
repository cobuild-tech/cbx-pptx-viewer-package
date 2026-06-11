/**
 * Viewer controller: mounts a {@link Deck} into a container element, renders one
 * slide at a time, scales it to fit the viewport, and handles navigation and
 * keyboard shortcuts. Framework wrappers (React, etc.) build on top of this.
 */
import type { Deck } from '../parse/deck.js';
import { renderSlide } from '../render/dom.js';

export type FitMode = 'contain' | 'width';

export interface ViewerOptions {
  fit?: FitMode;
  startIndex?: number;
  /** Enable arrow-key / space navigation on the container. Default true. */
  keyboard?: boolean;
  /** Called whenever the current slide index changes. */
  onChange?: (index: number, count: number) => void;
}

export class Viewer {
  private readonly deck: Deck;
  private readonly container: HTMLElement;
  private readonly holder: HTMLDivElement;
  private readonly fit: FitMode;
  private readonly onChange: ViewerOptions['onChange'];
  private index = 0;
  private slideEl: HTMLElement | null = null;
  private resizeObserver: ResizeObserver | null = null;
  private keyHandler: ((e: KeyboardEvent) => void) | null = null;

  constructor(deck: Deck, container: HTMLElement, options: ViewerOptions = {}) {
    this.deck = deck;
    this.container = container;
    this.fit = options.fit ?? 'contain';
    if (options.onChange) this.onChange = options.onChange;

    container.style.display = 'flex';
    container.style.alignItems = 'center';
    container.style.justifyContent = 'center';
    container.style.overflow = 'hidden';

    this.holder = document.createElement('div');
    this.holder.style.position = 'relative';
    container.appendChild(this.holder);

    if (options.keyboard !== false) this.enableKeyboard();
    this.resizeObserver = new ResizeObserver(() => this.applyScale());
    this.resizeObserver.observe(container);

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

  private applyScale(): void {
    if (!this.slideEl) return;
    const { wPx, hPx } = this.deck.size;
    const cw = this.container.clientWidth || wPx;
    const ch = this.container.clientHeight || hPx;
    const scale =
      this.fit === 'width' ? cw / wPx : Math.min(cw / wPx, ch / hPx);
    this.slideEl.style.transform = `scale(${scale})`;
    this.holder.style.width = `${wPx * scale}px`;
    this.holder.style.height = `${hPx * scale}px`;
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
    if (this.keyHandler) this.container.removeEventListener('keydown', this.keyHandler);
    if (this.slideEl) this.slideEl.remove();
    this.slideEl = null;
  }
}

/** Convenience: load nothing — just mount an existing deck. */
export function createViewer(
  deck: Deck,
  container: HTMLElement,
  options?: ViewerOptions,
): Viewer {
  return new Viewer(deck, container, options);
}
