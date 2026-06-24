/**
 * DocxViewer controller: mounts a {@link DocxDocument} into a container,
 * measures every block off-screen for accurate pagination, then renders
 * all pages in a vertically scrollable stack with a right-side thumbnail strip.
 */
import type { DocxDocument } from '../document/document.js';
import { renderPage, renderBlock } from '../render/dom.js';
import type { DocxPage } from '../model.js';

export interface DocxViewerOptions {
  startIndex?: number;
  /** Enable keyboard navigation (PageUp/Down, arrow keys). Default true. */
  keyboard?: boolean;
  /** Called whenever the most-visible page changes. */
  onChange?: (index: number, count: number) => void;
}

const SIDEBAR_W = 112; // px — thumbnail strip width
const ZOOM_STEP = 0.1;
const ZOOM_MIN  = 0.25;
const ZOOM_MAX  = 3.0;

export class DocxViewer {
  private readonly doc: DocxDocument;
  private readonly container: HTMLElement;
  private readonly scrollEl: HTMLDivElement;
  private readonly sidebarEl: HTMLDivElement;
  private readonly pageEls: HTMLElement[] = [];
  private readonly thumbEls: HTMLElement[] = [];
  private readonly zoomLabel: HTMLSpanElement;
  private readonly onChange: DocxViewerOptions['onChange'];
  private index = 0;
  private userZoom: number | null = 1;
  private resizeObserver: ResizeObserver | null = null;
  private intersectionObserver: IntersectionObserver | null = null;
  private keyHandler: ((e: KeyboardEvent) => void) | null = null;

  constructor(doc: DocxDocument, container: HTMLElement, options: DocxViewerOptions = {}) {
    this.doc = doc;
    this.container = container;
    if (options.onChange) this.onChange = options.onChange;

    container.style.position = 'relative';

    // ── Outer flex shell ──────────────────────────────────────────────────
    const shell = document.createElement('div');
    shell.style.position = 'absolute';
    shell.style.inset = '0';
    shell.style.display = 'flex';
    shell.style.overflow = 'hidden';
    container.appendChild(shell);

    // ── Main column (toolbar + scroll) ────────────────────────────────────
    const mainCol = document.createElement('div');
    mainCol.style.flex = '1';
    mainCol.style.minWidth = '0';
    mainCol.style.display = 'flex';
    mainCol.style.flexDirection = 'column';
    mainCol.style.overflow = 'hidden';
    shell.appendChild(mainCol);

    // ── Zoom toolbar ──────────────────────────────────────────────────────
    const toolbar = document.createElement('div');
    toolbar.style.display = 'flex';
    toolbar.style.alignItems = 'center';
    toolbar.style.justifyContent = 'center';
    toolbar.style.gap = '4px';
    toolbar.style.padding = '5px 12px';
    toolbar.style.background = '#1e1e1e';
    toolbar.style.borderBottom = '1px solid #333';
    toolbar.style.flexShrink = '0';
    toolbar.style.userSelect = 'none';
    mainCol.appendChild(toolbar);

    const btnStyle = (el: HTMLButtonElement) => {
      el.style.background = '#2e2e2e';
      el.style.border = '1px solid #444';
      el.style.borderRadius = '4px';
      el.style.color = '#ddd';
      el.style.cursor = 'pointer';
      el.style.fontSize = '14px';
      el.style.lineHeight = '1';
      el.style.padding = '3px 9px';
      el.style.transition = 'background 0.1s';
      el.addEventListener('mouseover', () => { el.style.background = '#3a3a3a'; });
      el.addEventListener('mouseout',  () => { el.style.background = '#2e2e2e'; });
    };

    const zoomOutBtn = document.createElement('button');
    zoomOutBtn.textContent = '−';
    zoomOutBtn.title = 'Zoom out';
    btnStyle(zoomOutBtn);
    zoomOutBtn.addEventListener('click', () => this.zoomOut());

    this.zoomLabel = document.createElement('span');
    this.zoomLabel.style.color = '#bbb';
    this.zoomLabel.style.fontSize = '12px';
    this.zoomLabel.style.minWidth = '46px';
    this.zoomLabel.style.textAlign = 'center';
    this.zoomLabel.style.cursor = 'pointer';
    this.zoomLabel.style.padding = '2px 4px';
    this.zoomLabel.style.borderRadius = '3px';
    this.zoomLabel.title = 'Reset to fit width';
    this.zoomLabel.addEventListener('click', () => this.zoomFit());
    this.zoomLabel.addEventListener('mouseover', () => { this.zoomLabel.style.background = '#2e2e2e'; });
    this.zoomLabel.addEventListener('mouseout',  () => { this.zoomLabel.style.background = 'transparent'; });

    const zoomInBtn = document.createElement('button');
    zoomInBtn.textContent = '+';
    zoomInBtn.title = 'Zoom in';
    btnStyle(zoomInBtn);
    zoomInBtn.addEventListener('click', () => this.zoomIn());

    toolbar.appendChild(zoomOutBtn);
    toolbar.appendChild(this.zoomLabel);
    toolbar.appendChild(zoomInBtn);

    // ── Main scroll area ──────────────────────────────────────────────────
    this.scrollEl = document.createElement('div');
    this.scrollEl.style.flex = '1';
    this.scrollEl.style.minWidth = '0';
    this.scrollEl.style.overflowY = 'auto';
    this.scrollEl.style.overflowX = 'auto';
    this.scrollEl.style.paddingTop = '28px';
    this.scrollEl.style.paddingBottom = '28px';
    mainCol.appendChild(this.scrollEl);

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

    // ── DOM-measured pagination then render ───────────────────────────────
    const pages = this.measureAndPaginate();
    this.renderPages(pages);

    this.applyZoom();
    this.updateActiveThumb(0);

    this.resizeObserver = new ResizeObserver(() => {
      if (this.userZoom === null) this.applyZoom();
    });
    this.resizeObserver.observe(container);

    this.setupIntersection();

    if (options.keyboard !== false) this.enableKeyboard();

    const start = options.startIndex ?? 0;
    if (start > 0) requestAnimationFrame(() => this.goTo(start));

    this.onChange?.(0, this.count);
  }

  // ── DOM measurement → repagination ────────────────────────────────────────

  /**
   * Render every block into a hidden off-screen container, measure its actual
   * rendered height, then call doc.repaginate() with those real heights.
   * Falls back to the heuristic pages (doc.pages) if the DOM is unavailable.
   */
  private measureAndPaginate(): DocxPage[] {
    if (typeof document === 'undefined') return this.doc.pages;

    const deps = { imageUrl: (p: string) => this.doc.imageUrl(p) };

    // Scratch container: off-screen but still participates in layout so
    // getBoundingClientRect returns real dimensions.
    const scratch = document.createElement('div');
    scratch.style.cssText = [
      'position:fixed',
      'top:-99999px',
      'left:0',
      'visibility:hidden',
      'pointer-events:none',
      'box-sizing:border-box',
    ].join(';');
    document.body.appendChild(scratch);

    // Use the first page's geometry for the scratch width.
    // For multi-section documents the width rarely changes, and even if it
    // does the difference is small — this is vastly more accurate than the
    // heuristic formula.
    const firstPage = this.doc.pages[0];
    if (!firstPage) {
      document.body.removeChild(scratch);
      return this.doc.pages;
    }

    const pages = this.doc.repaginate((block, contentWidthPx) => {
      // Images carry explicit dimensions — trust them directly.
      if (block.kind === 'image') return block.heightPx + 4;

      const el = renderBlock(block, deps);
      if (!el) return 10;

      scratch.style.width = `${contentWidthPx}px`;
      scratch.appendChild(el);

      if (block.kind === 'table') {
        // Return per-row heights so the paginator can split the table across pages.
        const trEls = el.querySelectorAll('tr');
        const rowHeights: number[] = [];
        for (const tr of trEls) {
          rowHeights.push(Math.max(0, tr.getBoundingClientRect().height));
        }
        scratch.removeChild(el);
        return rowHeights.length > 0 ? rowHeights : [el.getBoundingClientRect().height];
      }

      const h = el.getBoundingClientRect().height;
      scratch.removeChild(el);
      // Add a small margin buffer (paragraph spacing, etc.) to avoid tight cuts.
      return Math.max(4, h);
    });

    document.body.removeChild(scratch);
    return pages;
  }

  // ── Page + thumbnail rendering ────────────────────────────────────────────

  private renderPages(pages: DocxPage[]): void {
    const deps = { imageUrl: (p: string) => this.doc.imageUrl(p) };
    const thumbW = SIDEBAR_W - 20;

    for (let i = 0; i < pages.length; i++) {
      const page = pages[i]!;
      const pageEl = renderPage(page, deps);
      pageEl.dataset.pageIndex = String(i);

      const wrapper = document.createElement('div');
      wrapper.style.display = 'flex';
      wrapper.style.justifyContent = 'center';
      wrapper.style.marginBottom = '28px';
      wrapper.appendChild(pageEl);
      this.scrollEl.appendChild(wrapper);
      this.pageEls.push(pageEl);

      // Thumbnail
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

      const thumbPageEl = renderPage(page, deps);
      thumbPageEl.style.boxShadow = 'none';
      thumbPageEl.style.transformOrigin = 'top left';
      thumbPageEl.style.transform = `scale(${thumbScale})`;
      thumbPageEl.style.pointerEvents = 'none';
      thumb.appendChild(thumbPageEl);

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
  }

  // ── Public navigation / zoom API ──────────────────────────────────────────

  get count(): number { return this.pageEls.length; }
  get currentIndex(): number { return this.index; }

  goTo(index: number): void {
    const clamped = Math.max(0, Math.min(this.pageEls.length - 1, index));
    const wrapper = this.pageEls[clamped]?.parentElement;
    if (wrapper) wrapper.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  next(): void { this.goTo(this.index + 1); }
  prev(): void { this.goTo(this.index - 1); }

  zoomIn(): void {
    const next = Math.min(ZOOM_MAX, Math.round((this.currentZoomValue() + ZOOM_STEP) * 100) / 100);
    this.setZoom(next);
  }

  zoomOut(): void {
    const next = Math.max(ZOOM_MIN, Math.round((this.currentZoomValue() - ZOOM_STEP) * 100) / 100);
    this.setZoom(next);
  }

  setZoom(level: number): void {
    this.userZoom = Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, level));
    this.applyZoom();
  }

  zoomFit(): void {
    this.userZoom = null;
    this.applyZoom();
  }

  // ── Internal helpers ──────────────────────────────────────────────────────

  private currentZoomValue(): number {
    if (this.userZoom !== null) return this.userZoom;
    const cw = this.container.clientWidth - SIDEBAR_W;
    const pageEl = this.pageEls[0];
    // Read the natural page width from the element's inline style.
    const pageW = pageEl ? parseFloat(pageEl.style.width) || 794 : 794;
    return cw > 0 ? (cw - 48) / pageW : 1;
  }

  private applyZoom(): void {
    const cw = this.container.clientWidth - SIDEBAR_W;
    if (cw <= 0) return;

    for (let i = 0; i < this.pageEls.length; i++) {
      const pageEl = this.pageEls[i]!;
      const pageW = parseFloat(pageEl.style.width) || 794;
      const zoom = this.userZoom !== null ? this.userZoom : (cw - 48) / pageW;
      (pageEl.style as CSSStyleDeclaration & { zoom: string }).zoom = String(zoom);
    }

    const displayZoom = this.currentZoomValue();
    this.zoomLabel.textContent = `${Math.round(displayZoom * 100)}%`;
  }

  private updateActiveThumb(idx: number): void {
    for (let i = 0; i < this.thumbEls.length; i++) {
      const th = this.thumbEls[i]!;
      th.style.borderColor = i === idx ? '#4a9eff' : 'transparent';
      th.style.boxShadow   = i === idx ? '0 0 0 1px #4a9eff' : 'none';
    }
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
        let best = this.index, bestRatio = -1;
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
        case '+': case '=': e.preventDefault(); this.zoomIn(); break;
        case '-': e.preventDefault(); this.zoomOut(); break;
        case '0': e.preventDefault(); this.zoomFit(); break;
      }
    };
    this.container.addEventListener('keydown', this.keyHandler);
  }

  destroy(): void {
    this.resizeObserver?.disconnect();
    this.intersectionObserver?.disconnect();
    if (this.keyHandler) this.container.removeEventListener('keydown', this.keyHandler);
    this.container.querySelector('div')?.remove();
  }
}

export function createDocxViewer(
  doc: DocxDocument,
  container: HTMLElement,
  options?: DocxViewerOptions,
): DocxViewer {
  return new DocxViewer(doc, container, options);
}
