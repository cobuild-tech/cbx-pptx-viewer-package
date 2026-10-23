/**
 * Viewer controller: mounts a {@link DocxDocument} into a container and renders
 * its pages as a vertical, scrollable stack scaled to fit the container width —
 * the embedded-document style (vs PPTX's one-slide-at-a-time fit).
 *
 * Navigation (next/prev/goTo) scrolls to a page; the current page index is
 * reported via `onChange` as the user scrolls.
 */
import type { DocxDocument } from '../document/document.js';
import { renderPage, layoutTabs, renderCommentRail, layoutComments, COMMENT_RAIL_PX, type RevisionView } from '../render/dom.js';
import { installFallbackFonts, fontStack } from '../../oxml/fontMetrics.js';
import { installWebFontFamilies, type FontInstallation, type WebFontOptions } from '../../oxml/webfonts.js';
import { collectDocxFonts } from './fonts.js';
import { paginate, clearLayoutCache, blockSignature } from './paginate.js';
import { DocxEditContext } from '../edit/context.js';
import { DocxEditSession, type DocxParaFormat } from '../edit/session.js';
import { DocxEditController } from '../edit/controller.js';
import type { TextBoxOutline } from '../../oxml/edit/styles.js';
import type { RunFormat } from '../../oxml/edit/format.js';
import type { DocxBlock, DocxPage } from '../model.js';

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
  /**
   * Edit the document as in Word's Print Layout: type straight onto the pages,
   * with the caret and selection running across paragraphs, tables and page
   * breaks, and the pages re-laid out as the text changes. Header, footer and
   * footnote text stays read-only. Off by default.
   */
  editable?: boolean;
  /** Called after each edit, undo or redo. */
  onEdit?: () => void;
  /** Called when the caret moves, with the character formatting in effect there. */
  onSelectionChange?: (format: RunFormat) => void;
  /** Called when the caret moves, with the paragraph formatting in effect there. */
  onParaSelectionChange?: (format: DocxParaFormat) => void;
  /** @deprecated Paragraphs are no longer outlined while editing; ignored. */
  textBoxOutline?: TextBoxOutline;
  /**
   * Fetch fonts the document uses but the machine lacks from Google Fonts —
   * metric-compatible clones for Office fonts (Carlito for Calibri…) — and
   * re-paginate once they arrive. Default true; `false` keeps everything local
   * (a size-adjusted system font still stands in, so pages break in the same
   * places). An object customises the source.
   */
  webFonts?: boolean | WebFontOptions;
  /**
   * How tracked changes show: 'markup' (default) — insertions underlined and
   * deletions struck through in each reviewer's colour, with change bars;
   * 'final' — as if every change were accepted; 'original' — as if rejected.
   */
  trackChanges?: RevisionView;
  /** Show comments: highlighted ranges and cards beside the page. Default true. */
  comments?: boolean;
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
  private pageModels: DocxPage[];
  private editable: boolean;
  private editCtx: DocxEditContext | null = null;
  private session: DocxEditSession | null = null;
  private controller: DocxEditController | null = null;
  private readonly onEdit: DocxViewerOptions['onEdit'];
  private readonly onSelectionChange: DocxViewerOptions['onSelectionChange'];
  private readonly onParaSelectionChange: DocxViewerOptions['onParaSelectionChange'];
  private readonly onChange: DocxViewerOptions['onChange'];
  private readonly onScaleChange: DocxViewerOptions['onScaleChange'];
  /** 'fit-width' recomputes on resize; a number is a fixed user zoom. */
  private zoomMode: number | 'fit-width';
  private index = 0;
  private scale = 1;
  private resizeObserver: ResizeObserver | null = null;
  private keyHandler: ((e: KeyboardEvent) => void) | null = null;
  private scrollHandler: (() => void) | null = null;
  private fonts: FontInstallation;
  private readonly revisions: RevisionView;
  private readonly showComments: boolean;
  /** Width the comment rail adds beside pages that have comments (0 when none do). */
  private railW = 0;
  private disposeFallbackFonts: (() => void) | null = null;
  private destroyed = false;

  constructor(doc: DocxDocument, container: HTMLElement, options: DocxViewerOptions = {}) {
    this.doc = doc;
    this.container = container;
    this.zoomMode = options.zoom ?? 'fit-width';
    if (options.onChange) this.onChange = options.onChange;
    if (options.onScaleChange) this.onScaleChange = options.onScaleChange;
    if (options.onEdit) this.onEdit = options.onEdit;
    if (options.onSelectionChange) this.onSelectionChange = options.onSelectionChange;
    if (options.onParaSelectionChange) this.onParaSelectionChange = options.onParaSelectionChange;
    this.editable = options.editable ?? false;
    this.revisions = options.trackChanges ?? 'markup';
    this.showComments = options.comments ?? true;

    container.style.overflow = 'auto';
    container.style.background = '#525659';

    this.holder = document.createElement('div');
    this.holder.style.margin = '0 auto';

    this.pagesEl = document.createElement('div');
    this.pagesEl.style.display = 'flex';
    this.pagesEl.style.flexDirection = 'column';
    this.pagesEl.style.alignItems = 'center';
    this.pagesEl.style.gap = `${PAGE_GAP}px`;
    this.pagesEl.style.transformOrigin = 'top left';
    this.holder.appendChild(this.pagesEl);
    container.appendChild(this.holder);

    // Metric-matched stand-ins for fonts the machine lacks, before anything is
    // measured, so the first pagination already breaks where Word does.
    this.disposeFallbackFonts = installFallbackFonts(container.ownerDocument);
    const families = collectDocxFonts(doc.sections);
    this.fonts =
      options.webFonts === false
        ? { ready: Promise.resolve(), dispose() {} }
        : installWebFontFamilies(families, typeof options.webFonts === 'object' ? options.webFonts : {});
    // Local stand-in faces load asynchronously on first use, too.
    const fontDoc = container.ownerDocument;
    const localLoads =
      typeof fontDoc.fonts?.load === 'function'
        ? Promise.all(
            [...families].flatMap((f) =>
              ['', 'bold '].map((w) => fontDoc.fonts.load(`${w}16px ${fontStack(f)}`).catch(() => [])),
            ),
          )
        : Promise.resolve();
    Promise.all([this.fonts.ready, localLoads]).then(() => this.refreshAfterFonts());

    this.pageModels = [];
    if (this.editable) {
      this.editCtx = new DocxEditContext(doc);
      this.enableEditing();
    }
    this.renderContent();

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
    const maxPageW = this.pages.reduce((m, p) => Math.max(m, p.size.wPx), 1) + this.railW;
    // clientWidth already excludes the vertical scrollbar, so filling it exactly
    // leaves no horizontal scrollbar and no side gutter around the page.
    const avail = this.container.clientWidth || maxPageW;

    // Fit-width shrinks a wide page to the container but never upscales past
    // 100% — otherwise a wide panel opens the document zoomed in.
    const next =
      this.zoomMode === 'fit-width' ? Math.min(avail / maxPageW, 1) : this.zoomMode;
    this.scale = clamp(next, MIN_ZOOM, MAX_ZOOM);
    // Pin the (unscaled) stack width to the widest page so the top-left scale
    // fills the holder exactly — otherwise the stack fills the holder's already
    // scaled width and centering offsets get amplified by the transform.
    this.pagesEl.style.width = `${maxPageW}px`;
    this.pagesEl.style.transform = this.scale === 1 ? '' : `scale(${this.scale})`;

    // Transform doesn't change the layout box; size the holder to the scaled
    // dimensions so the scrollbars and `margin:auto` centering are correct.
    const naturalH = this.pagesEl.offsetHeight;
    this.holder.style.width = `${maxPageW * this.scale}px`;
    this.holder.style.height = `${naturalH * this.scale}px`;
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
      // PageUp/PageDown/Home/End all mean something else inside text. Without
      // this, an editor would jump pages mid-word.
      if (this.isEditingTarget(e.target)) return;
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

  // ─── Editing ───────────────────────────────────────────────────────────────

  /** Lay the document out into pages and draw them (with edit stamps when editing). */
  private renderContent(): void {
    const deps = {
      imageUrl: (p: string) => this.doc.imageUrl(p),
      revisions: this.revisions,
      comments: this.showComments,
      ...(this.editCtx ? { edit: this.editCtx } : {}),
    };
    this.editCtx?.reset();
    this.pageEls.length = 0;
    this.controller?.invalidate();

    // Flow sections into fixed-size pages (needs the DOM for measurement).
    this.pageModels = paginate(this.doc.sections, deps);
    const previous = this.drawn;
    this.drawn = [];
    let rail = false;
    this.pageModels.forEach((page, i) => {
      // A page whose content — and whose edit stamps — are what is already on
      // screen is kept: editing re-lays out the document on every keystroke,
      // and redrawing twenty unchanged pages each time is what made it lag.
      const sig = pageSignature(page, this.editCtx);
      const old = previous[i];
      const kept = old && old.sig === sig ? old : undefined;
      const entry = kept ?? { sig, el: this.drawPage(page, deps), rail: false };
      if (!kept) entry.rail = entry.el.firstElementChild?.nextElementSibling?.hasAttribute('data-cbx-comment-rail') ?? false;
      if (entry.rail) rail = true;
      this.drawn.push(entry);
      this.pageEls.push(entry.el);
      const at = this.pagesEl.children[i];
      if (at !== entry.el) this.pagesEl.insertBefore(entry.el, at ?? null);
      if (!kept) {
        layoutTabs(entry.el);
        if (entry.rail) layoutComments(entry.el);
      }
    });
    while (this.pagesEl.children.length > this.drawn.length) this.pagesEl.lastElementChild!.remove();
    this.railW = rail ? COMMENT_RAIL_PX + 12 : 0;
  }

  /** Pages on screen, with the signature each was drawn from. */
  private drawn: { sig: string; el: HTMLElement; rail: boolean }[] = [];

  /** Draw one page — beside its comment cards, as Word's markup pane does. */
  private drawPage(page: DocxPage, deps: Parameters<typeof renderPage>[1]): HTMLElement {
    const sheet = renderPage(page, deps);
    const rail = this.showComments ? renderCommentRail(page, deps) : undefined;
    if (!rail) return sheet;
    const el = document.createElement('div');
    el.style.display = 'flex';
    el.style.alignItems = 'flex-start';
    sheet.style.margin = '0';
    el.append(sheet, rail);
    return el;
  }

  /**
   * Fonts arrived: text now measures differently, so pages may break in other
   * places. Re-paginate, keeping the reader on the page they were on — unless
   * they are typing, when a re-render would throw away the caret.
   */
  private refreshAfterFonts(): void {
    if (this.destroyed) return;
    clearLayoutCache();
    // Tab stops and comment cards were placed by measuring the old fonts.
    this.drawn = [];
    // Re-laying out mid-typing would lose the caret: keep it across the redraw.
    const sel = this.controller?.readSelection();
    const scroll = this.container.scrollTop;
    const page = this.index;
    this.renderContent();
    this.applyScale();
    if (this.editable) {
      this.container.scrollTop = scroll;
      if (sel) this.controller?.setSelection(sel);
      this.onChange?.(this.index, this.count);
    } else if (page > 0) this.goTo(page);
    else this.onChange?.(this.index, this.count);
  }

  /** True if the event target is inside the editable page stack. */
  private isEditingTarget(target: EventTarget | null): boolean {
    return this.editable && target instanceof Node && this.pagesEl.contains(target);
  }

  private enableEditing(): void {
    this.editCtx ??= new DocxEditContext(this.doc);
    this.session ??= new DocxEditSession(this.doc, { onChange: () => this.onEdit?.() });
    this.controller = new DocxEditController({
      host: this.pagesEl,
      scroller: this.container,
      session: this.session,
      ctx: this.editCtx,
      render: () => {
        const scroll = this.container.scrollTop;
        this.renderContent();
        this.applyScale();
        this.container.scrollTop = scroll;
        this.updateCurrentFromScroll();
      },
      onSelectionChange: (format, para) => {
        this.onSelectionChange?.(format);
        this.onParaSelectionChange?.(para);
      },
    });
  }

  /** Kept for compatibility: edits now apply as they are made, so there is nothing to commit. */
  commitActive(): boolean {
    return false;
  }

  /** Apply character formatting to the selection (or, with none, to what is typed next). */
  applyFormat(format: RunFormat): boolean {
    if (!this.controller) return false;
    this.controller.format(format);
    return true;
  }

  /** Set alignment, spacing or the paragraph style on the selected paragraphs. */
  applyParagraphFormat(format: DocxParaFormat): void {
    this.controller?.paragraphFormat(format);
  }

  /** Turn the selected paragraphs into a bulleted or numbered list, or back out of one. */
  toggleList(kind: 'bullet' | 'number'): void {
    this.controller?.toggleList(kind);
  }

  /** Increase (+1) or decrease (−1) the indent — a list item's level, else ½ inch. */
  indent(delta: 1 | -1): void {
    this.controller?.indent(delta);
  }

  /** The document's paragraph styles, for a style picker (quick styles first). */
  paragraphStyles(): { id: string; name: string; quick: boolean }[] {
    const session = this.session ?? new DocxEditSession(this.doc);
    return session.styles().list.map(({ id, name, quick }) => ({ id, name, quick }));
  }

  /** Turn editing on or off. Both modes show the same pages. */
  setEditable(on: boolean): void {
    if (on === this.editable) return;
    this.editable = on;
    if (on) {
      this.editCtx = new DocxEditContext(this.doc);
      this.enableEditing();
    } else {
      this.teardownEditing();
      this.editCtx = null;
    }
    this.renderContent();
    this.applyScale();
  }

  get canUndo(): boolean {
    return this.session?.canUndo ?? false;
  }

  get canRedo(): boolean {
    return this.session?.canRedo ?? false;
  }

  undo(): void {
    this.controller?.undo();
  }

  redo(): void {
    this.controller?.redo();
  }

  /** True if the document has unsaved edits. */
  get hasEdits(): boolean {
    return this.session?.hasEdits ?? false;
  }

  /** Re-zip the document, edits included, as a .docx Blob. */
  exportBlob(): Blob {
    return this.doc.exportBlob();
  }

  private teardownEditing(): void {
    this.controller?.destroy();
    this.controller = null;
  }

  destroy(): void {
    this.destroyed = true;
    this.fonts.dispose();
    this.disposeFallbackFonts?.();
    this.resizeObserver?.disconnect();
    if (this.keyHandler) this.container.removeEventListener('keydown', this.keyHandler);
    if (this.scrollHandler) this.container.removeEventListener('scroll', this.scrollHandler);
    this.teardownEditing();
    this.pagesEl.remove();
    this.holder.remove();
  }
}

/** Everything a drawn page depends on, as one string (see {@link blockSignature}). */
function pageSignature(page: DocxPage, ctx: DocxEditContext | null): string {
  const blocks = (list: DocxBlock[] | undefined) => (list ?? []).map(blockSignature).join('\u0001');
  const { elements, header, footer, footnotes, bands, ...rest } = page;
  const bandSig = (bands ?? []).map((b) => `${JSON.stringify({ ...b, blocks: undefined })}:${b.blocks.map(blocks).join('\u0002')}`);
  const stamps = ctx ? [elements, header, footer, footnotes].map((l) => ctx.fingerprint(l)).join('\u0003') : '';
  return [JSON.stringify(rest), blocks(elements), blocks(header), blocks(footer), blocks(footnotes), bandSig.join('\u0004'), stamps].join('\u0005');
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
