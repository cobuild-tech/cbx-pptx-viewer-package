/**
 * DocxViewer controller: mounts a {@link DocxDocument} into a container and
 * renders all pages in a vertically scrollable stack with a right-side
 * thumbnail strip for quick navigation.
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

const SIDEBAR_W = 112; // px — thumbnail strip width

export class DocxViewer {
  private readonly doc: DocxDocument;
  private readonly container: HTMLElement;
  private readonly scrollEl: HTMLDivElement;
  private readonly sidebarEl: HTMLDivElement;
  private readonly pageEls: HTMLElement[] = [];
  private readonly thumbEls: HTMLElement[] = [];
  private readonly onChange: DocxViewerOptions['onChange'];
  private index = 0;
  private resizeObserver: ResizeObserver | null = null;
  private intersectionObserver: IntersectionObserver | null = null;
  private keyHandler: ((e: KeyboardEvent) => void) | null = null;

  constructor(doc: DocxDocument, container: HTMLElement, options: DocxViewerOptions = {}) {
    this.doc = doc;
    this.container = container;
    if (options.onChange) this.onChange = options.onChange;

    container.style.position = 'relative';

    // ── Outer flex shell (main scroll + sidebar) ──────────────────────────
    const shell = document.createElement('div');
    shell.style.position = 'absolute';
    shell.style.inset = '0';
    shell.style.display = 'flex';
    shell.style.overflow = 'hidden';
    container.appendChild(shell);

    // ── Main scroll area ──────────────────────────────────────────────────
    this.scrollEl = document.createElement('div');
    this.scrollEl.style.flex = '1';
    this.scrollEl.style.minWidth = '0';
    this.scrollEl.style.overflowY = 'auto';
    this.scrollEl.style.overflowX = 'hidden';
    this.scrollEl.style.paddingTop = '28px';
    this.scrollEl.style.paddingBottom = '28px';
    shell.appendChild(this.scrollEl);

    // ── Thumbnail sidebar ─────────────────────────────────────────────────
    this.sidebarEl = document.createElement('div');
    this.sidebarEl.style.width = `${SIDEBAR_W}px`;
    this.sidebarEl.style.flexShrink = '0';
    this.sidebarEl.style.overflowY = 'auto';
    this.sidebarEl.style.overflowX = 'hidden';
    this.sidebarEl.style.background = '#2e2e2e';
    this.sidebarEl.style.borderLeft = '1px solid #444';
    this.sidebarEl.style.paddingTop = '12px';
    this.sidebarEl.style.paddingBottom = '12px';
    this.sidebarEl.style.display = 'flex';
    this.sidebarEl.style.flexDirection = 'column';
    this.sidebarEl.style.alignItems = 'center';
    this.sidebarEl.style.gap = '8px';
    shell.appendChild(this.sidebarEl);

    // ── Render pages + thumbnails ─────────────────────────────────────────
    const thumbW = SIDEBAR_W - 20; // 92px usable thumbnail width

    for (let i = 0; i < doc.pages.length; i++) {
      const page = doc.pages[i]!;
      const pageEl = renderPage(page, { imageUrl: (p) => doc.imageUrl(p) });
      pageEl.dataset.pageIndex = String(i);

      const wrapper = document.createElement('div');
      wrapper.style.display = 'flex';
      wrapper.style.justifyContent = 'center';
      wrapper.style.marginBottom = '28px';
      wrapper.appendChild(pageEl);
      this.scrollEl.appendChild(wrapper);
      this.pageEls.push(pageEl);

      // Thumbnail card — renders actual page content scaled down
      const thumbScale = thumbW / page.size.wPx;
      const thumbH = Math.round(page.size.hPx * thumbScale);

      const thumb = document.createElement('div');
      thumb.style.width = `${thumbW}px`;
      thumb.style.height = `${thumbH}px`;
      thumb.style.flexShrink = '0';
      thumb.style.border = '2px solid transparent';
      thumb.style.borderRadius = '2px';
      thumb.style.cursor = 'pointer';
      thumb.style.position = 'relative';
      thumb.style.overflow = 'hidden';
      thumb.style.transition = 'border-color 0.15s';
      thumb.title = `Page ${i + 1}`;

      // Render page content at full size, then scale it down visually
      const thumbPageEl = renderPage(page, { imageUrl: (p) => doc.imageUrl(p) });
      thumbPageEl.style.boxShadow = 'none';
      thumbPageEl.style.transformOrigin = 'top left';
      thumbPageEl.style.transform = `scale(${thumbScale})`;
      // pointer-events off so clicks pass through to the thumb wrapper
      thumbPageEl.style.pointerEvents = 'none';
      thumb.appendChild(thumbPageEl);

      // Page number label overlaid at the bottom
      const label = document.createElement('div');
      label.textContent = String(i + 1);
      label.style.position = 'absolute';
      label.style.bottom = '0';
      label.style.left = '0';
      label.style.right = '0';
      label.style.textAlign = 'center';
      label.style.fontSize = '9px';
      label.style.color = '#555';
      label.style.background = 'rgba(255,255,255,0.8)';
      label.style.lineHeight = '14px';
      label.style.pointerEvents = 'none';
      thumb.appendChild(label);

      thumb.addEventListener('click', () => this.goTo(i));
      this.sidebarEl.appendChild(thumb);
      this.thumbEls.push(thumb);
    }

    this.applyZoom();
    this.updateActiveThumb(0);

    this.resizeObserver = new ResizeObserver(() => this.applyZoom());
    this.resizeObserver.observe(container);

    this.setupIntersection();

    if (options.keyboard !== false) this.enableKeyboard();

    const start = options.startIndex ?? 0;
    if (start > 0) requestAnimationFrame(() => this.goTo(start));

    this.onChange?.(0, this.count);
  }

  get count(): number {
    return this.doc.pages.length;
  }

  get currentIndex(): number {
    return this.index;
  }

  goTo(index: number): void {
    const clamped = Math.max(0, Math.min(this.doc.pages.length - 1, index));
    const wrapper = this.pageEls[clamped]?.parentElement;
    if (wrapper) {
      wrapper.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }

  next(): void { this.goTo(this.index + 1); }
  prev(): void { this.goTo(this.index - 1); }

  private applyZoom(): void {
    // Available width = shell width minus sidebar
    const cw = this.container.clientWidth - SIDEBAR_W;
    if (cw <= 0) return;
    for (let i = 0; i < this.pageEls.length; i++) {
      const pageEl = this.pageEls[i]!;
      const page = this.doc.pages[i]!;
      const zoom = Math.min(1, (cw - 48) / page.size.wPx);
      (pageEl.style as CSSStyleDeclaration & { zoom: string }).zoom = String(zoom);
    }
  }

  private updateActiveThumb(idx: number): void {
    for (let i = 0; i < this.thumbEls.length; i++) {
      const th = this.thumbEls[i]!;
      const active = i === idx;
      th.style.borderColor = active ? '#4a9eff' : 'transparent';
      th.style.boxShadow = active ? '0 0 0 1px #4a9eff' : 'none';
    }
    // Scroll the active thumbnail into view in the sidebar.
    this.thumbEls[idx]?.scrollIntoView({ block: 'nearest' });
  }

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
        let best = this.index;
        let bestRatio = -1;
        for (const [idx, ratio] of ratios) {
          if (ratio > bestRatio) { bestRatio = ratio; best = idx; }
        }
        if (best !== this.index) {
          this.index = best;
          this.updateActiveThumb(this.index);
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
        case 'ArrowDown': case 'PageDown': case ' ':
          e.preventDefault(); this.next(); break;
        case 'ArrowUp': case 'PageUp':
          e.preventDefault(); this.prev(); break;
        case 'Home': e.preventDefault(); this.goTo(0); break;
        case 'End':  e.preventDefault(); this.goTo(this.count - 1); break;
      }
    };
    this.container.addEventListener('keydown', this.keyHandler);
  }

  destroy(): void {
    this.resizeObserver?.disconnect();
    this.intersectionObserver?.disconnect();
    if (this.keyHandler) this.container.removeEventListener('keydown', this.keyHandler);
    this.container.querySelector('div')?.remove(); // remove shell
  }
}

export function createDocxViewer(
  doc: DocxDocument,
  container: HTMLElement,
  options?: DocxViewerOptions,
): DocxViewer {
  return new DocxViewer(doc, container, options);
}
