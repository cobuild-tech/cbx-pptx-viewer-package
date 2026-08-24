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
import { DocxEditContext } from '../edit/context.js';
import { DocxEditSession } from '../edit/session.js';
import { reconcileParagraph } from '../edit/reconcile.js';
import { renderFlow } from '../edit/flow.js';
import { readRunFormat } from '../edit/format.js';
import { EDIT_ATTR } from '../../oxml/edit/attrs.js';
import { installEditStyles, type TextBoxOutline } from '../../oxml/edit/styles.js';
import { applyFormatToSelection, formatAtSelection } from '../../oxml/edit/selection.js';
import type { RunFormat } from '../../oxml/edit/format.js';
import type { DocxParagraph } from '../model.js';
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
  /**
   * Edit the document's body text in place. Editing renders the document as one
   * continuous column rather than fixed pages — see edit/flow.ts for why — and
   * re-paginates when editing is switched off. Header and footer text stays
   * read-only. Off by default.
   */
  editable?: boolean;
  /** Called after a committed edit, undo or redo. */
  onEdit?: () => void;
  /** Called when the caret moves, with the formatting in effect there. */
  onSelectionChange?: (format: RunFormat) => void;
  /** How editable paragraphs are outlined. Ignored unless `editable`. */
  textBoxOutline?: TextBoxOutline;
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
  private flowEl: HTMLElement | null = null;
  private readonly onEdit: DocxViewerOptions['onEdit'];
  private readonly onSelectionChange: DocxViewerOptions['onSelectionChange'];
  private focusOutHandler: ((e: FocusEvent) => void) | null = null;
  private selectionHandler: (() => void) | null = null;
  private disposeStyles: (() => void) | null = null;
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
    if (options.onEdit) this.onEdit = options.onEdit;
    if (options.onSelectionChange) this.onSelectionChange = options.onSelectionChange;
    this.editable = options.editable ?? false;
    if (this.editable) {
      this.editCtx = new DocxEditContext(doc);
      this.session = new DocxEditSession(doc, { onChange: () => this.onEdit?.() });
      this.disposeStyles = installEditStyles(
        container.ownerDocument,
        options.textBoxOutline ?? 'hover',
      );
    }

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

    this.pageModels = [];
    this.renderContent();

    if (options.keyboard !== false) this.enableKeyboard();
    if (this.editable) this.enableEditing();
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

  /** Render either the paginated page stack or the continuous edit flow. */
  private renderContent(): void {
    const deps = {
      imageUrl: (p: string) => this.doc.imageUrl(p),
      ...(this.editCtx ? { edit: this.editCtx } : {}),
    };
    this.editCtx?.reset();
    this.pagesEl.replaceChildren();
    this.pageEls.length = 0;
    this.flowEl = null;

    if (this.editable) {
      this.flowEl = renderFlow(this.doc.sections, deps);
      this.pagesEl.appendChild(this.flowEl);
      // Pages are meaningless in flow mode; keep one entry so counts stay sane.
      this.pageModels = this.pageModels.length ? this.pageModels : [];
      return;
    }

    // Flow sections into fixed-size pages (needs the DOM for measurement).
    this.pageModels = paginate(this.doc.sections, deps);
    for (const page of this.pageModels) {
      const el = renderPage(page, deps);
      this.pageEls.push(el);
      this.pagesEl.appendChild(el);
    }
  }

  /** The editable paragraph element an event target sits inside, if any. */
  private paragraphElOf(target: EventTarget | null): HTMLElement | null {
    if (!this.editable || !(target instanceof Node)) return null;
    let node: Node | null = target;
    while (node && node !== this.pagesEl) {
      // The marker attribute is the signal: the renderer only stamps it on
      // paragraphs it also made contentEditable. (isContentEditable would be
      // equivalent in a browser but is unimplemented in jsdom.)
      if (node.nodeType === 1 && (node as Element).hasAttribute(EDIT_ATTR.para)) {
        return node as HTMLElement;
      }
      node = node.parentNode;
    }
    return null;
  }

  private isEditingTarget(target: EventTarget | null): boolean {
    return this.paragraphElOf(target) !== null;
  }

  private enableEditing(): void {
    // Commit when focus leaves a paragraph — including when it moves to another
    // paragraph, which focusout reports and blur does not.
    this.focusOutHandler = (e: FocusEvent) => {
      const from = this.paragraphElOf(e.target);
      if (!from) return;
      const to = this.paragraphElOf(e.relatedTarget);
      if (to === from) return;
      this.commitParagraphEl(from);
    };
    this.container.addEventListener('focusout', this.focusOutHandler, true);

    if (this.onSelectionChange) {
      this.selectionHandler = () => {
        if (!this.flowEl || !this.editCtx) return;
        const sel = this.container.ownerDocument.getSelection();
        if (!sel || sel.rangeCount === 0) return;
        if (!this.isEditingTarget(sel.getRangeAt(0).startContainer)) return;
        this.onSelectionChange?.(
          formatAtSelection(this.flowEl, (k) => this.editCtx?.resolve(k), readRunFormat),
        );
      };
      this.container.ownerDocument.addEventListener('selectionchange', this.selectionHandler);
    }
  }

  /** Read one edited paragraph back into the XML and re-render. */
  private commitParagraphEl(el: HTMLElement): boolean {
    if (!this.editCtx || !this.session) return false;
    const para = this.editCtx.resolve(el.getAttribute(EDIT_ATTR.para)) as
      | DocxParagraph
      | undefined;
    if (!para) return false;

    const edits = reconcileParagraph(el, (k) => this.editCtx?.resolve(k));
    if (!this.session.commitParagraph(para, edits)) return false;
    // Re-render so the committed XML — not the browser's improvised markup —
    // is what the user is looking at.
    this.renderContent();
    this.applyScale();
    return true;
  }

  /** Commit whatever paragraph currently has focus, if any. */
  commitActive(): boolean {
    const el = this.paragraphElOf(this.container.ownerDocument.activeElement);
    return el ? this.commitParagraphEl(el) : false;
  }

  /** Apply formatting to the current selection and commit it. */
  applyFormat(format: RunFormat): boolean {
    if (!this.editable || !this.flowEl) return false;
    const changed = applyFormatToSelection(this.flowEl, format);
    if (!changed) return false;
    const el = this.paragraphElOf(changed) ?? changed;
    return this.commitParagraphEl(el as HTMLElement);
  }

  /** Turn editing on or off, re-rendering into the matching layout. */
  setEditable(on: boolean): void {
    if (on === this.editable) return;
    if (this.editable) this.commitActive();
    this.editable = on;
    if (on) {
      this.editCtx ??= new DocxEditContext(this.doc);
      this.session ??= new DocxEditSession(this.doc, { onChange: () => this.onEdit?.() });
      this.disposeStyles ??= installEditStyles(this.container.ownerDocument, 'hover');
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
    if (this.session?.undo()) {
      this.renderContent();
      this.applyScale();
    }
  }

  redo(): void {
    if (this.session?.redo()) {
      this.renderContent();
      this.applyScale();
    }
  }

  /** True if the document has unsaved edits. */
  get hasEdits(): boolean {
    return this.session?.hasEdits ?? false;
  }

  /** Re-zip the document, edits included, as a .docx Blob. */
  exportBlob(): Blob {
    // Flush anything still being typed before packaging.
    this.commitActive();
    return this.doc.exportBlob();
  }

  private teardownEditing(): void {
    if (this.focusOutHandler) {
      this.container.removeEventListener('focusout', this.focusOutHandler, true);
      this.focusOutHandler = null;
    }
    if (this.selectionHandler) {
      this.container.ownerDocument.removeEventListener('selectionchange', this.selectionHandler);
      this.selectionHandler = null;
    }
    this.disposeStyles?.();
    this.disposeStyles = null;
  }

  destroy(): void {
    this.resizeObserver?.disconnect();
    if (this.keyHandler) this.container.removeEventListener('keydown', this.keyHandler);
    if (this.scrollHandler) this.container.removeEventListener('scroll', this.scrollHandler);
    this.teardownEditing();
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
