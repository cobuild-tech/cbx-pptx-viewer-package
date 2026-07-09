/**
 * Viewer controller: mounts a {@link DocxDocument} into a container and renders
 * its pages as a vertical, scrollable stack scaled to fit the container width —
 * the embedded-document style (vs PPTX's one-slide-at-a-time fit).
 *
 * Navigation (next/prev/goTo) scrolls to a page; the current page index is
 * reported via `onChange` as the user scrolls.
 */
import type { DocxDocument } from '../document/document.js';
import { renderPage } from '../render/dom.js';
import { paginate } from './paginate.js';
import type { DocxPage } from '../model.js';

export interface DocxViewerOptions {
  startIndex?: number;
  /** Enable PageUp/PageDown/Home/End navigation on the container. Default true. */
  keyboard?: boolean;
  /**
   * Initial zoom. A number is an explicit scale (1 = 100%); `'fit-width'`
   * (default) fits the page to the container width, capped at 100% so wide
   * panels don't upscale.
   */
  zoom?: number | 'fit-width';
  /** Called when the current (top-most visible) page changes. */
  onChange?: (index: number, count: number) => void;
  /** Called whenever the effective zoom scale changes (1 = 100%). */
  onScaleChange?: (scale: number) => void;
}

const PAGE_GAP = 16;
const MIN_ZOOM = 0.25;
const MAX_ZOOM = 5;

export class DocxViewer {
  private readonly doc: DocxDocument;
  private readonly container: HTMLElement;
  private readonly holder: HTMLDivElement;
  private readonly pagesEl: HTMLDivElement;
  private readonly pageEls: HTMLElement[] = [];
  private readonly pageModels: DocxPage[];
  private readonly onChange: DocxViewerOptions['onChange'];
  private readonly onScaleChange: DocxViewerOptions['onScaleChange'];
  /** 'fit-width' recomputes on resize; a number is a fixed user zoom. */
  private zoomMode: number | 'fit-width';
  private index = 0;
  private scale = 1;
  private resizeObserver: ResizeObserver | null = null;
  private keyHandler: ((e: KeyboardEvent) => void) | null = null;
  private scrollHandler: (() => void) | null = null;

  constructor(doc: DocxDocument, container: HTMLElement, options: DocxViewerOptions = {}) {
    this.doc = doc;
    this.container = container;
    this.zoomMode = options.zoom ?? 'fit-width';
    if (options.onChange) this.onChange = options.onChange;
    if (options.onScaleChange) this.onScaleChange = options.onScaleChange;

    container.style.overflow = 'auto';
    container.style.background = '#525659';

    this.holder = document.createElement('div');
    this.holder.style.margin = '0 auto';
    this.holder.style.padding = `${PAGE_GAP}px 0`;

    this.pagesEl = document.createElement('div');
    this.pagesEl.style.display = 'flex';
    this.pagesEl.style.flexDirection = 'column';
    this.pagesEl.style.alignItems = 'center';
    this.pagesEl.style.gap = `${PAGE_GAP}px`;
    this.pagesEl.style.transformOrigin = 'top left';
    this.holder.appendChild(this.pagesEl);
    container.appendChild(this.holder);

    const deps = { imageUrl: (p: string) => this.doc.imageUrl(p) };
    // Flow sections into fixed-size pages (needs the DOM for measurement).
    this.pageModels = paginate(this.doc.sections, deps);
    for (const page of this.pageModels) {
      const el = renderPage(page, deps);
      this.pageEls.push(el);
      this.pagesEl.appendChild(el);
    }

    if (options.keyboard !== false) this.enableKeyboard();
    this.scrollHandler = () => this.updateCurrentFromScroll();
    container.addEventListener('scroll', this.scrollHandler, { passive: true });
    this.resizeObserver = new ResizeObserver(() => this.applyScale());
    this.resizeObserver.observe(container);

    this.applyScale();
    if (options.startIndex) this.goTo(options.startIndex);
    else this.onChange?.(0, this.count);
  }

  get count(): number {
    return this.pages.length;
  }
  get currentIndex(): number {
    return this.index;
  }
  private get pages(): DocxPage[] {
    return this.pageModels;
  }

  goTo(index: number): void {
    const clamped = Math.max(0, Math.min(this.count - 1, index));
    this.index = clamped;
    const el = this.pageEls[clamped];
    if (el) this.container.scrollTop = el.offsetTop * this.scale;
    this.onChange?.(this.index, this.count);
  }

  next(): void {
    this.goTo(this.index + 1);
  }
  prev(): void {
    this.goTo(this.index - 1);
  }

  /** Current effective zoom scale (1 = 100%). */
  get currentScale(): number {
    return this.scale;
  }

  /** Set an explicit zoom (1 = 100%); switches out of fit-width mode. */
  setZoom(scale: number): void {
    this.zoomMode = clamp(scale, MIN_ZOOM, MAX_ZOOM);
    this.applyScale();
  }
  zoomIn(): void {
    this.setZoom(this.scale * 1.25);
  }
  zoomOut(): void {
    this.setZoom(this.scale / 1.25);
  }
  /** Re-fit each page to the container width (capped at 100%). */
  fitWidth(): void {
    this.zoomMode = 'fit-width';
    this.applyScale();
  }

  /** Scale the page stack and size the holder box so scrollbars/centering match. */
  private applyScale(): void {
    const maxPageW = this.pages.reduce((m, p) => Math.max(m, p.size.wPx), 1);
    const avail = (this.container.clientWidth || maxPageW) - PAGE_GAP * 2;

    const next =
      this.zoomMode === 'fit-width' ? Math.min(avail / maxPageW, 1) : this.zoomMode;
    this.scale = clamp(next, MIN_ZOOM, MAX_ZOOM);
    this.pagesEl.style.transform = this.scale === 1 ? '' : `scale(${this.scale})`;

    // Transform doesn't change the layout box; size the holder to the scaled
    // dimensions so the scrollbars and `margin:auto` centering are correct.
    const naturalH = this.pagesEl.offsetHeight;
    this.holder.style.width = `${maxPageW * this.scale}px`;
    this.holder.style.height = `${naturalH * this.scale + PAGE_GAP * 2}px`;
    this.onScaleChange?.(this.scale);
  }

  private updateCurrentFromScroll(): void {
    const y = this.container.scrollTop / (this.scale || 1);
    let idx = 0;
    for (let i = 0; i < this.pageEls.length; i++) {
      if (this.pageEls[i]!.offsetTop <= y + 4) idx = i;
      else break;
    }
    if (idx !== this.index) {
      this.index = idx;
      this.onChange?.(this.index, this.count);
    }
  }

  private enableKeyboard(): void {
    if (this.container.tabIndex < 0) this.container.tabIndex = 0;
    this.keyHandler = (e: KeyboardEvent) => {
      if (e.ctrlKey || e.metaKey) {
        if (e.key === '=' || e.key === '+') {
          e.preventDefault();
          this.zoomIn();
          return;
        }
        if (e.key === '-') {
          e.preventDefault();
          this.zoomOut();
          return;
        }
        if (e.key === '0') {
          e.preventDefault();
          this.fitWidth();
          return;
        }
      }
      switch (e.key) {
        case 'PageDown':
          e.preventDefault();
          this.next();
          break;
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
    if (this.scrollHandler) this.container.removeEventListener('scroll', this.scrollHandler);
    this.pagesEl.remove();
    this.holder.remove();
  }
}

function clamp(n: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, n));
}

export function createDocxViewer(
  doc: DocxDocument,
  container: HTMLElement,
  options?: DocxViewerOptions,
): DocxViewer {
  return new DocxViewer(doc, container, options);
}
