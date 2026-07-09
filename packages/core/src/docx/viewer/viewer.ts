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
   * Sizing. `'width'` (default) scales each page to fill the container width and
   * scrolls vertically. `'actual'` renders at 100% with horizontal scroll.
   */
  fit?: 'width' | 'actual';
  /** Called when the current (top-most visible) page changes. */
  onChange?: (index: number, count: number) => void;
}

const PAGE_GAP = 16;

export class DocxViewer {
  private readonly doc: DocxDocument;
  private readonly container: HTMLElement;
  private readonly holder: HTMLDivElement;
  private readonly pagesEl: HTMLDivElement;
  private readonly pageEls: HTMLElement[] = [];
  private readonly pageModels: DocxPage[];
  private readonly onChange: DocxViewerOptions['onChange'];
  private readonly fit: 'width' | 'actual';
  private index = 0;
  private scale = 1;
  private resizeObserver: ResizeObserver | null = null;
  private keyHandler: ((e: KeyboardEvent) => void) | null = null;
  private scrollHandler: (() => void) | null = null;

  constructor(doc: DocxDocument, container: HTMLElement, options: DocxViewerOptions = {}) {
    this.doc = doc;
    this.container = container;
    this.fit = options.fit ?? 'width';
    if (options.onChange) this.onChange = options.onChange;

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

  /** Scale the page stack to fit the container width and size the holder box. */
  private applyScale(): void {
    const maxPageW = this.pages.reduce((m, p) => Math.max(m, p.size.wPx), 1);
    const avail = this.container.clientWidth || maxPageW;

    this.scale = this.fit === 'width' ? avail / maxPageW : 1;
    this.pagesEl.style.transform = this.scale === 1 ? '' : `scale(${this.scale})`;

    // Transform doesn't change layout box size; size the holder so scrollbars
    // and centering reflect the scaled dimensions.
    const naturalH = this.pagesEl.offsetHeight;
    this.holder.style.width = `${maxPageW * this.scale}px`;
    this.holder.style.height = `${naturalH * this.scale + PAGE_GAP * 2}px`;
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

export function createDocxViewer(
  doc: DocxDocument,
  container: HTMLElement,
  options?: DocxViewerOptions,
): DocxViewer {
  return new DocxViewer(doc, container, options);
}
