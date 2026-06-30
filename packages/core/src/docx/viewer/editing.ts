/**
 * Inline WYSIWYG editing controller for the DocxViewer.
 *
 * Strategy: the rendered DOM is only a *view + input surface*. Paragraphs are
 * made contenteditable; on blur (commit) we reconcile each run span's text back
 * into replaceText ops, apply them atomically, and the document re-renders from
 * the (mutated) OOXML tree — so contenteditable can never drift the model.
 * Formatting reads the live selection and maps it to run ranges. Structural
 * actions act on the last-focused paragraph/cell (tracked here).
 */
import type { DocxDocument } from '../document/document.js';
import { selectionToRunSegments, domRunText, type RunSegment } from '../edit/selection.js';
import type { RunPropPatch } from '../edit/ops.js';

export interface EditContext {
  runId?: string;
  paragraphId?: string;
  cellId?: string;
}

export class DocxEditController {
  private readonly doc: DocxDocument;
  private root: HTMLElement | null = null;
  private committing = false;
  private ctx: EditContext = {};
  /** Last non-empty selection, captured before a toolbar input can clear it. */
  private lastSegments: RunSegment[] = [];

  private readonly onFocusOut = (e: FocusEvent) => {
    const para = e.target instanceof HTMLElement ? e.target.closest<HTMLElement>('.docx-para') : null;
    if (para) this.commitParagraph(para);
  };
  private readonly onFocusIn = (e: FocusEvent) => this.updateContext(e.target as Node | null);
  private readonly onSelect = () => {
    if (typeof window === 'undefined' || !this.root) return;
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) this.updateContext(sel.getRangeAt(0).startContainer);
    // Remember any non-empty selection so formatting survives a toolbar input
    // (color picker / size select) stealing focus and collapsing the selection.
    const segs = selectionToRunSegments(this.root);
    if (segs.length > 0) this.lastSegments = segs;
  };
  private readonly onKeyDown = (e: KeyboardEvent) => {
    // Enter commits the paragraph instead of injecting <div>/<br> markup.
    if (e.key === 'Enter' && !e.shiftKey && e.target instanceof HTMLElement) {
      e.preventDefault();
      e.target.blur();
    }
  };

  constructor(doc: DocxDocument) {
    this.doc = doc;
  }

  /** Make the current page DOM editable and wire listeners. Idempotent per render. */
  attach(root: HTMLElement): void {
    this.root = root;
    // Fresh DOM (e.g. after a re-render): previous selection/context is stale.
    this.ctx = {};
    this.lastSegments = [];
    injectEditStyles();
    root.querySelectorAll<HTMLElement>('.docx-para').forEach((p) => {
      p.contentEditable = 'true';
      p.spellcheck = false;
      p.style.outline = 'none';
    });
    // Snapshot each run's text so commit can detect what actually changed.
    root.querySelectorAll<HTMLElement>('.docx-run[data-docx-id]').forEach((span) => {
      span.dataset.origText = domRunText(span);
    });
    root.addEventListener('focusout', this.onFocusOut, true);
    root.addEventListener('focusin', this.onFocusIn, true);
    root.addEventListener('mouseup', this.onSelect);
    root.addEventListener('keyup', this.onSelect);
    root.addEventListener('keydown', this.onKeyDown, true);
  }

  detach(): void {
    const root = this.root;
    if (root) {
      root.removeEventListener('focusout', this.onFocusOut, true);
      root.removeEventListener('focusin', this.onFocusIn, true);
      root.removeEventListener('mouseup', this.onSelect);
      root.removeEventListener('keyup', this.onSelect);
      root.removeEventListener('keydown', this.onKeyDown, true);
    }
    this.root = null;
  }

  /** The run/paragraph/cell the user last interacted with (for structural actions). */
  get context(): EditContext {
    return this.ctx;
  }

  /** Toggle a boolean run prop based on the focused run's current rendered state. */
  toggleProp(prop: 'bold' | 'italic' | 'underline' | 'strike'): void {
    const span = this.focusedRunSpan();
    const current = span ? readRenderedBool(span, prop) : false;
    this.format({ [prop]: !current } as RunPropPatch);
  }

  private focusedRunSpan(): HTMLElement | null {
    if (!this.root || !this.ctx.runId) return null;
    // nodeIds contain no '"' or '\\', so a quoted attribute selector is safe.
    return this.root.querySelector<HTMLElement>(`.docx-run[data-docx-id="${this.ctx.runId}"]`);
  }

  /**
   * Apply formatting to the current selection. If the selection spans runs they
   * are all formatted in one undo step; a collapsed cursor formats the run it's in.
   */
  format(props: RunPropPatch): void {
    if (!this.root) return;
    const live = selectionToRunSegments(this.root);
    const segs = live.length > 0 ? live : this.lastSegments;
    try {
      if (segs.length > 0) {
        this.doc.formatRunRanges(
          segs.map((s) => ({ runId: s.runId, start: s.start, end: s.end })),
          props,
        );
      } else if (this.ctx.runId) {
        this.doc.setRunProps(this.ctx.runId, props);
      }
    } catch {
      // Stale ids after a re-render — ignore; the user can reselect and retry.
    }
    // After applying, the selection no longer maps to the re-rendered DOM.
    this.lastSegments = [];
  }

  /** Reconcile a committed paragraph's DOM text into per-run replaceText ops. */
  private commitParagraph(para: HTMLElement): void {
    if (this.committing) return;
    const edits: { kind: 'replaceText'; target: string; text: string }[] = [];
    para.querySelectorAll<HTMLElement>('.docx-run[data-docx-id]').forEach((span) => {
      const runId = span.dataset.docxId!;
      const text = domRunText(span);
      if (text !== (span.dataset.origText ?? '')) {
        edits.push({ kind: 'replaceText', target: runId, text });
      }
    });
    if (edits.length === 0) return;
    this.committing = true;
    try {
      this.doc.applyEdits(edits);
    } finally {
      this.committing = false;
    }
  }

  private updateContext(node: Node | null): void {
    const el = node instanceof HTMLElement ? node : node?.parentElement ?? null;
    if (!el) return;
    this.ctx = {
      runId: el.closest<HTMLElement>('.docx-run[data-docx-id]')?.dataset.docxId,
      paragraphId: el.closest<HTMLElement>('.docx-para[data-docx-id]')?.dataset.docxId,
      cellId: el.closest<HTMLElement>('.docx-cell[data-docx-id]')?.dataset.docxId,
    };
  }
}

let editStylesInjected = false;
/** One-time CSS so users can see which paragraph is being edited. */
function injectEditStyles(): void {
  if (editStylesInjected || typeof document === 'undefined') return;
  const style = document.createElement('style');
  style.textContent =
    '.docx-para[contenteditable="true"]:hover{cursor:text;}' +
    '.docx-para[contenteditable="true"]:focus{background:rgba(43,87,154,0.07);border-radius:2px;}';
  document.head.appendChild(style);
  editStylesInjected = true;
}

/** Read a run span's current rendered boolean style (for toggle state). */
function readRenderedBool(el: HTMLElement, prop: 'bold' | 'italic' | 'underline' | 'strike'): boolean {
  if (typeof getComputedStyle === 'undefined') return false;
  const s = getComputedStyle(el);
  switch (prop) {
    case 'bold':
      return s.fontWeight === 'bold' || Number(s.fontWeight) >= 600;
    case 'italic':
      return s.fontStyle === 'italic';
    case 'underline':
      return `${s.textDecorationLine} ${s.textDecoration}`.includes('underline');
    case 'strike':
      return `${s.textDecorationLine} ${s.textDecoration}`.includes('line-through');
  }
}
