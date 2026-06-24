/**
 * DocxViewer controller: mounts a {@link DocxDocument} into a container and
 * renders all pages in a vertically scrollable stack.
 *
 * Unlike the PPTX viewer (one slide at a time), DOCX pages flow top-to-bottom
 * like a real document. Each page is a white card scaled to fill the container
 * width via CSS `zoom` (which adjusts layout dimensions, unlike `transform:
 * scale`). An IntersectionObserver tracks the most-visible page and fires
 * onChange so the React toolbar can show the current page number.
 *
 * `goTo(index)` / `next()` / `prev()` scroll to the target page smoothly.
 */
import type { DocxDocument } from '../document/document.js';
import { renderPage } from '../render/dom.js';

export interface DocxViewerOptions {
  startIndex?: number;
  /** Enable keyboard navigation (PageUp/Down, arrow keys). Default true. */
  keyboard?: boolean;
  /** Called whenever the most-visible page changes. */
  onChange?: (index: number, count: number) => void;
}

export class DocxViewer {
  private readonly doc: DocxDocument;
  private readonly container: HTMLElement;
  private readonly scrollEl: HTMLDivElement;
  private readonly pageEls: HTMLElement[] = [];
  private readonly onChange: DocxViewerOptions['onChange'];
  private index = 0;
  private resizeObserver: ResizeObserver | null = null;
  private intersectionObserver: IntersectionObserver | null = null;
  private keyHandler: ((e: KeyboardEvent) => void) | null = null;

  constructor(doc: DocxDocument, container: HTMLElement, options: DocxViewerOptions = {}) {
    this.doc = doc;
    this.container = container;
    if (options.onChange) this.onChange = options.onChange;

    // Fill the container with an absolutely-positioned scroll wrapper so the
    // viewer owns the scroll context (the React wrapper just provides a div).
    container.style.position = 'relative';

    this.scrollEl = document.createElement('div');
    this.scrollEl.style.position = 'absolute';
    this.scrollEl.style.inset = '0';
    this.scrollEl.style.overflowY = 'auto';
    this.scrollEl.style.overflowX = 'hidden';
    this.scrollEl.style.paddingTop = '28px';
    this.scrollEl.style.paddingBottom = '28px';
    container.appendChild(this.scrollEl);

    // Render all pages into the scroll container.
    for (let i = 0; i < doc.pages.length; i++) {
      const page = doc.pages[i]!;
      const pageEl = renderPage(page, { imageUrl: (p) => doc.imageUrl(p) }, i === doc.pages.length - 1);
      pageEl.dataset.pageIndex = String(i);

      // Wrapper centres each page card horizontally.
      const wrapper = document.createElement('div');
      wrapper.style.display = 'flex';
      wrapper.style.justifyContent = 'center';
      wrapper.style.marginBottom = '28px';
      wrapper.appendChild(pageEl);
      this.scrollEl.appendChild(wrapper);
      this.pageEls.push(pageEl);
    }

    this.applyZoom();

    this.resizeObserver = new ResizeObserver(() => this.applyZoom());
    this.resizeObserver.observe(container);

    this.setupIntersection();

    if (options.keyboard !== false) this.enableKeyboard();

    // Jump to starting page after a frame so the DOM is laid out.
    const start = options.startIndex ?? 0;
    if (start > 0) requestAnimationFrame(() => this.goTo(start));

    // Fire initial onChange.
    this.onChange?.(0, this.count);
  }

  get count(): number {
    return this.doc.pages.length;
  }

  get currentIndex(): number {
    return this.index;
  }

  /** Scroll smoothly to the given page index. */
  goTo(index: number): void {
    const clamped = Math.max(0, Math.min(this.doc.pages.length - 1, index));
    const wrapper = this.pageEls[clamped]?.parentElement;
    if (wrapper) {
      wrapper.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }

  next(): void {
    this.goTo(this.index + 1);
  }

  prev(): void {
    this.goTo(this.index - 1);
  }

  /** Scale every page to fill the container width using CSS zoom. */
  private applyZoom(): void {
    const cw = this.container.clientWidth;
    if (!cw) return;
    for (let i = 0; i < this.pageEls.length; i++) {
      const pageEl = this.pageEls[i]!;
      const page = this.doc.pages[i]!;
      const zoom = Math.min(1, (cw - 48) / page.size.wPx);
      // CSS zoom scales layout dimensions (unlike transform: scale).
      (pageEl.style as CSSStyleDeclaration & { zoom: string }).zoom = String(zoom);
    }
  }

  /** Track which page is most visible and fire onChange. */
  private setupIntersection(): void {
    if (typeof IntersectionObserver === 'undefined') return;

    const ratios = new Map<number, number>();

    this.intersectionObserver = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          const el = entry.target as HTMLElement;
          const idx = Number(el.dataset.pageIndex ?? '-1');
          if (idx >= 0) ratios.set(idx, entry.intersectionRatio);
        }
        // Pick the page with the highest visible ratio.
        let best = this.index;
        let bestRatio = -1;
        for (const [idx, ratio] of ratios) {
          if (ratio > bestRatio) { bestRatio = ratio; best = idx; }
        }
        if (best !== this.index) {
          this.index = best;
          this.onChange?.(this.index, this.count);
        }
      },
      { root: this.scrollEl, threshold: [0, 0.1, 0.25, 0.5, 0.75, 1] },
    );

    for (const el of this.pageEls) this.intersectionObserver.observe(el);
  }

  private enableKeyboard(): void {
    if (this.container.tabIndex < 0) this.container.tabIndex = 0;
    this.keyHandler = (e: KeyboardEvent) => {
      switch (e.key) {
        case 'ArrowDown':
        case 'PageDown':
        case ' ':
          e.preventDefault();
          this.next();
          break;
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
    this.intersectionObserver?.disconnect();
    if (this.keyHandler) this.container.removeEventListener('keydown', this.keyHandler);
    this.scrollEl.remove();
  }
}

export function createDocxViewer(
  doc: DocxDocument,
  container: HTMLElement,
  options?: DocxViewerOptions,
): DocxViewer {
  return new DocxViewer(doc, container, options);
}
