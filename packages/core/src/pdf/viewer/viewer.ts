/**
 * PdfViewer controller: mounts a {@link PdfDocument} into a container,
 * renders all pages to canvas elements, and presents a vertically scrollable
 * stack with a right-side thumbnail strip and zoom controls.
 *
 * Mirrors the DocxViewer API exactly: goTo / next / prev / zoomIn / zoomOut /
 * setZoom / zoomFit, keyboard navigation, IntersectionObserver page tracking,
 * ResizeObserver for fit-to-width re-scale.
 */
import type { PdfDocument } from '../document/document.js';
import type { PdfPage } from '../model.js';

export interface PdfViewerOptions {
  startIndex?: number;
  /** Enable keyboard navigation (PageUp/Down, arrow keys, +/-). Default true. */
  keyboard?: boolean;
  /** Called whenever the most-visible page changes. */
  onChange?: (index: number, count: number) => void;
}

const SIDEBAR_W = 112;
const THUMB_W = SIDEBAR_W - 20;
const RENDER_SCALE = 2.0;
const ZOOM_STEP = 0.1;
const ZOOM_MIN = 0.25;
const ZOOM_MAX = 3.0;

export class PdfViewer {
  private readonly doc: PdfDocument;
  private readonly container: HTMLElement;
  private readonly scrollEl: HTMLDivElement;
  private readonly sidebarEl: HTMLDivElement;
  private readonly pageContainers: HTMLElement[] = [];
  private readonly thumbContainers: HTMLElement[] = [];
  private readonly zoomLabel: HTMLSpanElement;
  private readonly onChange: PdfViewerOptions['onChange'];
  private index = 0;
  private userZoom: number | null = null; // null = fit-to-width
  private resizeObserver: ResizeObserver | null = null;
  private intersectionObserver: IntersectionObserver | null = null;
  private keyHandler: ((e: KeyboardEvent) => void) | null = null;

  constructor(doc: PdfDocument, container: HTMLElement, options: PdfViewerOptions = {}) {
    this.doc = doc;
    this.container = container;
    if (options.onChange) this.onChange = options.onChange;

    container.style.position = 'relative';

    // ── Outer flex shell ──────────────────────────────────────────────────
    const shell = document.createElement('div');
    shell.style.cssText = 'position:absolute;inset:0;display:flex;overflow:hidden;';
    container.appendChild(shell);

    // ── Main column (toolbar + scroll area) ───────────────────────────────
    const mainCol = document.createElement('div');
    mainCol.style.cssText = 'flex:1;min-width:0;display:flex;flex-direction:column;overflow:hidden;';
    shell.appendChild(mainCol);

    // ── Zoom toolbar ──────────────────────────────────────────────────────
    const toolbar = document.createElement('div');
    toolbar.style.cssText = [
      'display:flex', 'align-items:center', 'justify-content:center',
      'gap:4px', 'padding:5px 12px', 'background:#1e1e1e',
      'border-bottom:1px solid #333', 'flex-shrink:0', 'user-select:none',
    ].join(';');
    mainCol.appendChild(toolbar);

    const styleBtn = (btn: HTMLButtonElement) => {
      btn.style.cssText = [
        'background:#2e2e2e', 'border:1px solid #444', 'border-radius:4px',
        'color:#ddd', 'cursor:pointer', 'font-size:14px', 'line-height:1',
        'padding:3px 9px', 'transition:background 0.1s',
      ].join(';');
      btn.addEventListener('mouseover', () => { btn.style.background = '#3a3a3a'; });
      btn.addEventListener('mouseout',  () => { btn.style.background = '#2e2e2e'; });
    };

    const zoomOutBtn = document.createElement('button');
    zoomOutBtn.textContent = '−';
    zoomOutBtn.title = 'Zoom out';
    styleBtn(zoomOutBtn);
    zoomOutBtn.addEventListener('click', () => { this.zoomOut(); });

    this.zoomLabel = document.createElement('span');
    this.zoomLabel.style.cssText = [
      'color:#bbb', 'font-size:12px', 'min-width:46px', 'text-align:center',
      'cursor:pointer', 'padding:2px 4px', 'border-radius:3px',
    ].join(';');
    this.zoomLabel.title = 'Reset to fit width';
    this.zoomLabel.addEventListener('click',      () => { this.zoomFit(); });
    this.zoomLabel.addEventListener('mouseover',  () => { this.zoomLabel.style.background = '#2e2e2e'; });
    this.zoomLabel.addEventListener('mouseout',   () => { this.zoomLabel.style.background = 'transparent'; });

    const zoomInBtn = document.createElement('button');
    zoomInBtn.textContent = '+';
    zoomInBtn.title = 'Zoom in';
    styleBtn(zoomInBtn);
    zoomInBtn.addEventListener('click', () => { this.zoomIn(); });

    toolbar.appendChild(zoomOutBtn);
    toolbar.appendChild(this.zoomLabel);
    toolbar.appendChild(zoomInBtn);

    // ── Scroll area ───────────────────────────────────────────────────────
    this.scrollEl = document.createElement('div');
    this.scrollEl.style.cssText = 'flex:1;min-width:0;overflow-y:auto;overflow-x:auto;padding:28px 0;';
    mainCol.appendChild(this.scrollEl);

    // ── Thumbnail sidebar ─────────────────────────────────────────────────
    this.sidebarEl = document.createElement('div');
    this.sidebarEl.style.cssText = [
      `width:${SIDEBAR_W}px`, 'flex-shrink:0', 'overflow-y:auto', 'overflow-x:hidden',
      'background:#2e2e2e', 'border-left:1px solid #444', 'padding:12px 0',
      'display:flex', 'flex-direction:column', 'align-items:center', 'gap:8px',
    ].join(';');
    shell.appendChild(this.sidebarEl);

    // ── Build page placeholders and kick off async rendering ──────────────
    this.buildPages();

    this.applyZoom();
    this.updateActiveThumb(0);

    this.resizeObserver = new ResizeObserver(() => {
      if (this.userZoom === null) this.applyZoom();
    });
    this.resizeObserver.observe(container);

    this.setupIntersection();

    if (options.keyboard !== false) this.enableKeyboard();

    const start = options.startIndex ?? 0;
    if (start > 0) requestAnimationFrame(() => { this.goTo(start); });

    this.onChange?.(0, this.count);
  }

  // ── Page + thumbnail construction ─────────────────────────────────────────

  private buildPages(): void {
    for (let i = 0; i < this.doc.pages.length; i++) {
      const page = this.doc.pages[i] as PdfPage;

      // ── Main page container (white placeholder, correct natural dimensions) ─
      const pageContainer = document.createElement('div');
      pageContainer.dataset.pageIndex = String(i);
      pageContainer.style.cssText = [
        `width:${page.widthPx}px`, `height:${page.heightPx}px`,
        'background:#fff', 'box-shadow:0 2px 8px rgba(0,0,0,0.4)',
        'position:relative', 'overflow:hidden',
      ].join(';');

      const wrapper = document.createElement('div');
      wrapper.style.cssText = 'display:flex;justify-content:center;margin-bottom:28px;';
      wrapper.appendChild(pageContainer);
      this.scrollEl.appendChild(wrapper);
      this.pageContainers.push(pageContainer);

      // Render at 2× for retina; CSS size = natural page dimensions.
      this.doc.renderPage(i, RENDER_SCALE)
        .then((canvas) => {
          canvas.style.width  = `${page.widthPx}px`;
          canvas.style.height = `${page.heightPx}px`;
          pageContainer.appendChild(canvas);
        })
        .catch(() => {
          Object.assign(pageContainer.style, {
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            fontSize: '13px', color: '#e57373', fontFamily: 'system-ui, sans-serif',
          });
          pageContainer.textContent = `Failed to render page ${i + 1}`;
        });

      // ── Thumbnail ─────────────────────────────────────────────────────────
      const thumbScale = THUMB_W / page.widthPx;
      const thumbH = Math.round(page.heightPx * thumbScale);

      const thumb = document.createElement('div');
      thumb.style.cssText = [
        `width:${THUMB_W}px`, `height:${thumbH}px`, 'flex-shrink:0',
        'border:2px solid transparent', 'border-radius:2px', 'cursor:pointer',
        'position:relative', 'overflow:hidden', 'transition:border-color 0.15s', 'background:#fff',
      ].join(';');
      thumb.title = `Page ${i + 1}`;

      // Render thumb at a scale that produces exactly THUMB_W × thumbH pixels.
      const thumbRenderScale = thumbScale * 1.5; // 1.5× for acceptable sharpness
      this.doc.renderPage(i, thumbRenderScale)
        .then((canvas) => {
          canvas.style.width  = `${THUMB_W}px`;
          canvas.style.height = `${thumbH}px`;
          thumb.appendChild(canvas);
        })
        .catch(() => { /* non-critical — placeholder background suffices */ });

      const label = document.createElement('div');
      label.textContent = String(i + 1);
      label.style.cssText = [
        'position:absolute', 'bottom:0', 'left:0', 'right:0',
        'text-align:center', 'font-size:9px', 'color:#555',
        'background:rgba(255,255,255,0.8)', 'line-height:14px', 'pointer-events:none',
      ].join(';');
      thumb.appendChild(label);

      thumb.addEventListener('click', () => { this.goTo(i); });
      this.sidebarEl.appendChild(thumb);
      this.thumbContainers.push(thumb);
    }
  }

  // ── Public navigation / zoom API ──────────────────────────────────────────

  get count(): number { return this.pageContainers.length; }
  get currentIndex(): number { return this.index; }

  goTo(index: number): void {
    const clamped = Math.max(0, Math.min(this.pageContainers.length - 1, index));
    const wrapper = this.pageContainers[clamped]?.parentElement;
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
    const page = this.doc.pages[0];
    const pageW = page ? page.widthPx : 595;
    return cw > 0 ? (cw - 48) / pageW : 1;
  }

  private applyZoom(): void {
    const cw = this.container.clientWidth - SIDEBAR_W;
    if (cw <= 0) return;

    for (let i = 0; i < this.pageContainers.length; i++) {
      const container = this.pageContainers[i] as HTMLElement;
      const page = this.doc.pages[i] as PdfPage;
      const zoom = this.userZoom !== null ? this.userZoom : (cw - 48) / page.widthPx;
      (container.style as CSSStyleDeclaration & { zoom: string }).zoom = String(zoom);
    }

    this.zoomLabel.textContent = `${Math.round(this.currentZoomValue() * 100)}%`;
  }

  private updateActiveThumb(idx: number): void {
    for (let i = 0; i < this.thumbContainers.length; i++) {
      const th = this.thumbContainers[i] as HTMLElement;
      th.style.borderColor = i === idx ? '#4a9eff' : 'transparent';
      th.style.boxShadow   = i === idx ? '0 0 0 1px #4a9eff' : 'none';
    }
    this.thumbContainers[idx]?.scrollIntoView({ block: 'nearest' });
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

    for (const el of this.pageContainers) this.intersectionObserver.observe(el);
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
        case '-':           e.preventDefault(); this.zoomOut(); break;
        case '0':           e.preventDefault(); this.zoomFit(); break;
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

export function createPdfViewer(
  doc: PdfDocument,
  container: HTMLElement,
  options?: PdfViewerOptions,
): PdfViewer {
  return new PdfViewer(doc, container, options);
}
