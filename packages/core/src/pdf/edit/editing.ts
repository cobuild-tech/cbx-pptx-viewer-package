/**
 * PdfEditController — manages interactive contenteditable text overlays
 * on top of PDF canvas pages.
 *
 * ── Existing text block editing ───────────────────────────────────────────
 *  Click a block in edit mode → contenteditable div with full formatting toolbar.
 *  Drag the grab-bar to reposition. Box auto-expands (fit-content) — no resize handle.
 *  Shift+Enter = newline · Enter = commit · Escape = cancel.
 *  On commit: emits replaceText (if text changed) + styleBlock (if style changed).
 *
 * ── Free-form annotation features ────────────────────────────────────────
 *  • Click empty page area in edit mode → place new text annotation.
 *  • Inline toolbar: 50+ Google Fonts font picker, Bold, Italic, alignment,
 *    size ±, 18 color swatches, custom OS color picker.
 *  • Drag from grab-bar. Box auto-expands (fit-content) — no resize handle.
 *  • Shift+Enter = newline · Enter = commit · Escape = cancel.
 *  • Clear all text + blur = delete annotation (removeAnnotation op).
 *
 * ── Canvas overdraw (existing blocks only) ───────────────────────────────
 *  getImageData() saves originals; putImageData() restores on undo.
 */
import type { PdfTextBlock, PdfAnnotation, PdfEditOp, PdfBlockStyle } from '../model.js';
import { showFontPicker }                                               from './fontPicker.js';
import { resolveCssFontStack }                                          from './fonts.js';

// ── Constants ──────────────────────────────────────────────────────────────

const ANN = {
  DEFAULT_FONT_SIZE: 12,
  MIN_FONT_SIZE:      6,
  MAX_FONT_SIZE:     96,
  DEFAULT_WIDTH:    200,
  MIN_WIDTH:         80,
  MIN_HEIGHT:        24,
  GRAB_BAR_H:        18,
} as const;

const PALETTE: ReadonlyArray<{ hex: string; label: string }> = [
  { hex: '#000000', label: 'Black'     },
  { hex: '#1f2937', label: 'Dark gray' },
  { hex: '#6b7280', label: 'Gray'      },
  { hex: '#d1d5db', label: 'Silver'    },
  { hex: '#ffffff', label: 'White'     },
  { hex: '#b91c1c', label: 'Dark red'  },
  { hex: '#dc2626', label: 'Red'       },
  { hex: '#ea580c', label: 'Orange'    },
  { hex: '#d97706', label: 'Amber'     },
  { hex: '#16a34a', label: 'Green'     },
  { hex: '#0891b2', label: 'Cyan'      },
  { hex: '#1d4ed8', label: 'Blue'      },
  { hex: '#1e3a8a', label: 'Dark blue' },
  { hex: '#7c3aed', label: 'Violet'    },
  { hex: '#db2777', label: 'Pink'      },
  { hex: '#92400e', label: 'Brown'     },
  { hex: '#065f46', label: 'Teal'      },
  { hex: '#78350f', label: 'Caramel'   },
];

// ── Helpers ────────────────────────────────────────────────────────────────

function relativeLuminance(r: number, g: number, b: number): number {
  const lin = (c: number) => { c /= 255; return c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4; };
  return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
}

// ── Types ──────────────────────────────────────────────────────────────────

/** Commits one or more ops as a single undoable group. */
export type EditCommitCallback = (ops: PdfEditOp[]) => void;

interface OverlayState {
  pageIndex:     number;
  overlayEl:     HTMLDivElement;
  blockEls:      Map<string, HTMLDivElement>; // blockId → wrapper div
  annotationEls: Map<string, HTMLDivElement>; // annotationId → wrapper div
}

interface SavedRegion { imageData: ImageData; x: number; y: number; }

/** Shared style state used by both block and annotation toolbars. */
interface ElementStyle {
  color:      string;
  fontSize:   number;
  fontWeight: 'normal' | 'bold';
  fontStyle:  'normal' | 'italic';
  fontFamily: string;   // display name (e.g. 'Inter', 'Arial')
  textAlign:  'left' | 'center' | 'right';
}

// ── Controller ─────────────────────────────────────────────────────────────

export class PdfEditController {
  private readonly blocks:          PdfTextBlock[][];
  private readonly annotations:     Map<string, PdfAnnotation>;
  private readonly liveEdits:       Map<string, string>;
  private readonly blockStylesMap:  Map<string, PdfBlockStyle>;
  private readonly onCommit:        EditCommitCallback;
  private readonly pageContainers:  HTMLElement[];

  private readonly formattingSlot?: HTMLElement;
  private overlays:    OverlayState[] = [];
  private editable = false;
  private annCounter = 0;

  /** Carries last-used style to the next new annotation or block edit. */
  private styleDefaults: Omit<ElementStyle, 'fontSize'> = {
    color: '#000000', fontWeight: 'normal', fontStyle: 'normal',
    fontFamily: 'Arial', textAlign: 'left',
  };

  private readonly bgCache     = new Map<string, { bg: string; textColor: string }>();
  private readonly savedPixels = new Map<string, SavedRegion>();

  constructor(
    pageContainers: HTMLElement[],
    blocks:         PdfTextBlock[][],
    liveEdits:      Map<string, string>,
    annotations:    Map<string, PdfAnnotation>,
    blockStylesMap: Map<string, PdfBlockStyle>,
    onCommit:       EditCommitCallback,
    formattingSlot?: HTMLElement,
  ) {
    this.blocks          = blocks;
    this.liveEdits       = liveEdits;
    this.annotations     = annotations;
    this.blockStylesMap  = blockStylesMap;
    this.onCommit        = onCommit;
    this.pageContainers  = pageContainers;
    this.formattingSlot  = formattingSlot;
    this.mountOverlays(pageContainers);
  }

  // ── Public API ─────────────────────────────────────────────────────────────

  setEditable(enabled: boolean): void {
    this.editable = enabled;
    const flatBlocks = this.blocks.flat();
    for (const { overlayEl, blockEls, annotationEls } of this.overlays) {
      overlayEl.style.pointerEvents = enabled ? 'auto' : 'none';
      for (const [id, wrapper] of blockEls) {
        const block = flatBlocks.find(b => b.id === id);
        if (block) this.applyBlockWrapperStyle(wrapper, block, this.blockStylesMap.get(id), enabled);
      }
      for (const [id, wrapper] of annotationEls) {
        const ann = this.annotations.get(id);
        if (ann) this.applyWrapperStyle(wrapper, ann, enabled);
      }
    }
    if (!enabled && this.formattingSlot) this.formattingSlot.innerHTML = '';
  }

  refreshAll(): void {
    const flatBlocks = this.blocks.flat();
    for (const { pageIndex, overlayEl, blockEls, annotationEls } of this.overlays) {
      // ── Refresh blocks ───────────────────────────────────────────────────
      for (const [id, wrapper] of blockEls) {
        const block      = flatBlocks.find(b => b.id === id);
        const blockStyle = this.blockStylesMap.get(id);
        if (!block) continue;
        const contentDiv = this.blockContentOf(wrapper);
        const text       = this.liveEdits.get(id) ?? block.text;
        if (contentDiv && contentDiv.innerText !== text) contentDiv.innerText = text;
        this.applyBlockWrapperStyle(wrapper, block, blockStyle, this.editable);
      }

      // ── Refresh annotations ──────────────────────────────────────────────
      for (const [id, wrapper] of annotationEls) {
        if (!this.annotations.has(id)) { wrapper.remove(); annotationEls.delete(id); }
      }
      for (const ann of [...this.annotations.values()].filter(a => a.pageIndex === pageIndex)) {
        if (!annotationEls.has(ann.id)) {
          const wrapper = this.buildCommittedEl(ann, pageIndex, annotationEls);
          overlayEl.appendChild(wrapper);
          annotationEls.set(ann.id, wrapper);
        }
        const wrapper    = annotationEls.get(ann.id)!;
        const contentDiv = this.contentOf(wrapper);
        if (contentDiv && contentDiv.innerText !== ann.text) contentDiv.innerText = ann.text;
        this.applyWrapperStyle(wrapper, ann, this.editable);
      }
    }
  }

  unmount(): void {
    if (this.formattingSlot) this.formattingSlot.innerHTML = '';
    for (const { overlayEl } of this.overlays) overlayEl.remove();
    this.overlays = [];
  }

  // ── Overlay mounting ────────────────────────────────────────────────────────

  private mountOverlays(containers: HTMLElement[]): void {
    for (let pi = 0; pi < containers.length; pi++) {
      const container  = containers[pi]!;
      const pageBlocks = this.blocks[pi] ?? [];

      const overlayEl = document.createElement('div');
      overlayEl.style.cssText = 'position:absolute;inset:0;pointer-events:none;overflow:visible;z-index:2;';

      const blockEls:      Map<string, HTMLDivElement> = new Map();
      const annotationEls: Map<string, HTMLDivElement> = new Map();

      for (const block of pageBlocks) {
        const { wrapper, grabBar, contentDiv, rightHandle, bottomHandle } = this.createBlockWrapper(block);
        this.setupBlock(wrapper, grabBar, contentDiv, rightHandle, bottomHandle, block, container);
        this.applyBlockWrapperStyle(wrapper, block, this.blockStylesMap.get(block.id), false);
        blockEls.set(block.id, wrapper);
        overlayEl.appendChild(wrapper);
      }

      for (const ann of this.annotations.values()) {
        if (ann.pageIndex !== pi) continue;
        const wrapper = this.buildCommittedEl(ann, pi, annotationEls);
        annotationEls.set(ann.id, wrapper);
        overlayEl.appendChild(wrapper);
      }

      overlayEl.addEventListener('mousedown', (e: MouseEvent) => {
        if (!this.editable || e.target !== overlayEl) return;
        e.preventDefault();
        const rect = overlayEl.getBoundingClientRect();
        const zoom = this.zoom(container);
        this.startNewAnnotation(
          pi,
          (e.clientX - rect.left) / zoom,
          (e.clientY - rect.top)  / zoom,
          container, overlayEl, annotationEls,
        );
      });

      container.appendChild(overlayEl);
      this.overlays.push({ pageIndex: pi, overlayEl, blockEls, annotationEls });
    }
  }

  // ── Block wrapper helpers ───────────────────────────────────────────────────

  private blockContentOf(wrapper: HTMLDivElement): HTMLDivElement | null {
    return wrapper.querySelector<HTMLDivElement>('[data-block-content]');
  }

  private createBlockWrapper(block: PdfTextBlock): {
    wrapper: HTMLDivElement; grabBar: HTMLDivElement; contentDiv: HTMLDivElement;
    rightHandle: HTMLDivElement; bottomHandle: HTMLDivElement;
  } {
    const wrapper = document.createElement('div');
    wrapper.dataset.blockId = block.id;
    wrapper.style.cssText = [
      'position:absolute',
      `left:${block.cssX}px`, `top:${block.cssY}px`,
      `width:${block.cssWidth}px`,
      'box-sizing:border-box', 'z-index:3',
    ].join(';');

    const grabBar = document.createElement('div');
    grabBar.dataset.blockGrab = '';
    grabBar.style.cssText = [
      'display:flex', 'align-items:center', 'justify-content:center',
      'overflow:hidden', 'height:0', 'cursor:move',
      'border-radius:4px 4px 0 0',
      'font-size:10px', 'letter-spacing:2px',
      'color:rgba(147,197,253,0.7)',
      'background:rgba(59,130,246,0.12)',
      'user-select:none',
    ].join(';');
    grabBar.textContent = '• • • • • • •';
    wrapper.appendChild(grabBar);

    const contentDiv = document.createElement('div');
    contentDiv.dataset.blockContent = '';
    contentDiv.style.cssText = [
      'width:100%', 'min-height:1.4em', 'height:auto',
      `font-size:${block.fontSize}px`,
      'font-weight:normal', 'font-style:normal',
      'font-family:Helvetica,Arial,sans-serif',
      'text-align:left', 'line-height:1.4',
      'white-space:pre-wrap', 'word-break:break-word',
      'box-sizing:border-box', 'padding:3px 5px',
      'background:transparent', 'color:transparent',
      'outline:none', 'border:none', 'cursor:default',
    ].join(';');
    contentDiv.innerText = this.liveEdits.get(block.id) ?? block.text;
    wrapper.appendChild(contentDiv);

    const rightHandle = document.createElement('div');
    rightHandle.dataset.blockResizeRight = '';
    rightHandle.style.cssText = [
      'position:absolute', 'top:0', 'right:0',
      'width:6px', 'height:100%', 'cursor:ew-resize',
      'z-index:5', 'display:none',
      'background:rgba(59,130,246,0.3)',
      'border-radius:0 2px 2px 0',
    ].join(';');
    wrapper.appendChild(rightHandle);

    const bottomHandle = document.createElement('div');
    bottomHandle.dataset.blockResizeBottom = '';
    bottomHandle.style.cssText = [
      'position:absolute', 'bottom:0', 'left:0',
      'width:100%', 'height:6px', 'cursor:ns-resize',
      'z-index:5', 'display:none',
      'background:rgba(59,130,246,0.3)',
      'border-radius:0 0 2px 2px',
    ].join(';');
    wrapper.appendChild(bottomHandle);

    return { wrapper, grabBar, contentDiv, rightHandle, bottomHandle };
  }

  private setupBlock(
    wrapper:      HTMLDivElement,
    grabBar:      HTMLDivElement,
    contentDiv:   HTMLDivElement,
    rightHandle:  HTMLDivElement,
    bottomHandle: HTMLDivElement,
    block:        PdfTextBlock,
    container:    HTMLElement,
  ): void {
    // Mutable style state — re-synced from blockStylesMap on every focus.
    const blockState: ElementStyle = {
      color: '#000000', fontSize: block.fontSize,
      fontWeight: 'normal', fontStyle: 'normal',
      fontFamily: 'Arial', textAlign: 'left',
    };

    let removeBlockToolbar: (() => void) | null = null;

    wrapper.addEventListener('click', (e: MouseEvent) => {
      if (!this.editable) return;
      const t = e.target as HTMLElement;
      if (t.dataset.blockGrab !== undefined) return;
      contentDiv.contentEditable = 'true';
      contentDiv.style.cursor    = 'text';
      contentDiv.focus();
    });

    contentDiv.addEventListener('focus', () => {
      // Re-sync from blockStylesMap so undo/redo state is always reflected.
      const cur = this.blockStylesMap.get(block.id) ?? {};
      blockState.color      = cur.color      ?? this.styleDefaults.color;
      blockState.fontSize   = cur.fontSize   ?? block.fontSize;
      blockState.fontWeight = cur.fontWeight ?? 'normal';
      blockState.fontStyle  = cur.fontStyle  ?? 'normal';
      blockState.fontFamily = cur.fontFamily ?? this.styleDefaults.fontFamily;
      blockState.textAlign  = cur.textAlign  ?? 'left';

      contentDiv.style.fontFamily = resolveCssFontStack(blockState.fontFamily);
      contentDiv.style.fontWeight = blockState.fontWeight;
      contentDiv.style.fontStyle  = blockState.fontStyle;
      contentDiv.style.fontSize   = `${blockState.fontSize}px`;
      contentDiv.style.color      = blockState.color;
      contentDiv.style.textAlign  = blockState.textAlign;
      contentDiv.style.background = '#ffffff';

      removeBlockToolbar?.();
      removeBlockToolbar = this.showToolbar(wrapper, blockState, contentDiv, (pickerColor) => {
        // Color picker stole focus — commit the color change immediately.
        const old = this.blockStylesMap.get(block.id) ?? {};
        const upd: PdfBlockStyle = { ...old, color: pickerColor };
        this.blockStylesMap.set(block.id, upd);
        this.onCommit([{ kind: 'styleBlock', blockId: block.id, pageIndex: block.pageIndex, oldStyle: old, newStyle: upd }]);
      });
    });

    contentDiv.addEventListener('keydown', (e: KeyboardEvent) => {
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); contentDiv.blur(); return; }
      if (e.key === 'Escape') {
        e.preventDefault();
        contentDiv.innerText = this.liveEdits.get(block.id) ?? block.text;
        contentDiv.contentEditable = 'false'; contentDiv.style.cursor = 'default';
        removeBlockToolbar?.(); removeBlockToolbar = null;
        this.applyBlockWrapperStyle(wrapper, block, this.blockStylesMap.get(block.id), this.editable);
        return;
      }
      e.stopPropagation();
    });

    contentDiv.addEventListener('blur', () => {
      removeBlockToolbar?.(); removeBlockToolbar = null;

      const newText  = contentDiv.innerText ?? '';
      const prevText = this.liveEdits.get(block.id) ?? block.text;
      const oldStyle = this.blockStylesMap.get(block.id) ?? {};

      const ops: PdfEditOp[] = [];

      if (newText !== prevText) {
        ops.push({
          kind: 'replaceText', blockId: block.id, pageIndex: block.pageIndex,
          oldText: prevText, newText,
        });
      }

      // Build new style, preserving position overrides set by drag.
      const rawNewStyle: PdfBlockStyle = {
        ...(oldStyle.cssX  !== undefined  && { cssX:  oldStyle.cssX  }),
        ...(oldStyle.cssY  !== undefined  && { cssY:  oldStyle.cssY  }),
        ...(oldStyle.cssWidth !== undefined && { cssWidth: oldStyle.cssWidth }),
        color:      blockState.color,
        fontSize:   blockState.fontSize,
        fontWeight: blockState.fontWeight,
        fontStyle:  blockState.fontStyle,
        fontFamily: blockState.fontFamily,
        textAlign:  blockState.textAlign,
      };

      if (JSON.stringify(oldStyle) !== JSON.stringify(rawNewStyle)) {
        this.blockStylesMap.set(block.id, rawNewStyle);
        ops.push({
          kind: 'styleBlock', blockId: block.id, pageIndex: block.pageIndex,
          oldStyle, newStyle: rawNewStyle,
        });
      }

      if (ops.length > 0) this.onCommit(ops);

      contentDiv.contentEditable = 'false'; contentDiv.style.cursor = 'default';
      this.applyBlockWrapperStyle(wrapper, block, this.blockStylesMap.get(block.id), this.editable);
    });

    this.wireBlockDrag(grabBar, wrapper, block, container, contentDiv, rightHandle, bottomHandle);
    this.wireBlockDrag(wrapper, wrapper, block, container, contentDiv, grabBar, rightHandle, bottomHandle);

    this.wireResize(rightHandle, wrapper, 'h', container, (newW) => {
      const cur = this.blockStylesMap.get(block.id) ?? {};
      const old = { ...cur };
      const upd: PdfBlockStyle = { ...cur, cssWidth: newW };
      this.blockStylesMap.set(block.id, upd);
      this.onCommit([{ kind: 'styleBlock', blockId: block.id, pageIndex: block.pageIndex, oldStyle: old, newStyle: upd }]);
    });

    this.wireResize(bottomHandle, wrapper, 'v', container, (newH) => {
      const cur = this.blockStylesMap.get(block.id) ?? {};
      const old = { ...cur };
      const upd: PdfBlockStyle = { ...cur, cssHeight: newH };
      this.blockStylesMap.set(block.id, upd);
      this.onCommit([{ kind: 'styleBlock', blockId: block.id, pageIndex: block.pageIndex, oldStyle: old, newStyle: upd }]);
    });
  }

  private applyBlockWrapperStyle(
    wrapper:    HTMLDivElement,
    block:      PdfTextBlock,
    blockStyle: PdfBlockStyle | undefined,
    editMode:   boolean,
  ): void {
    const contentDiv  = this.blockContentOf(wrapper);
    const grabBar     = wrapper.querySelector<HTMLDivElement>('[data-block-grab]');
    const rightHandle = wrapper.querySelector<HTMLDivElement>('[data-block-resize-right]');
    const bottomHandle = wrapper.querySelector<HTMLDivElement>('[data-block-resize-bottom]');

    const x     = blockStyle?.cssX      ?? block.cssX;
    const y     = blockStyle?.cssY      ?? block.cssY;
    const w     = blockStyle?.cssWidth  ?? block.cssWidth;
    const fs    = blockStyle?.fontSize  ?? block.fontSize;
    const fw    = blockStyle?.fontWeight ?? 'normal';
    const fi    = blockStyle?.fontStyle  ?? 'normal';
    const ff    = blockStyle?.fontFamily ?? '';
    const color = blockStyle?.color     ?? '';
    const align = blockStyle?.textAlign ?? 'left';

    wrapper.style.left = `${x}px`;
    wrapper.style.top  = `${y}px`;

    const storedText  = this.liveEdits.get(block.id);
    const isEdited    = storedText !== undefined && storedText !== block.text;
    const isStyled    = !!blockStyle;
    const showOverlay = editMode || isEdited || isStyled;

    // Only erase canvas when the block has been physically repositioned — opaque overlay
    // background handles all other cases (text edits, style changes) without touching the canvas.
    const hasMoved = !!blockStyle &&
      ((blockStyle.cssX !== undefined && blockStyle.cssX !== block.cssX) ||
       (blockStyle.cssY !== undefined && blockStyle.cssY !== block.cssY));

    const applyContentStyle = () => {
      if (!contentDiv) return;
      contentDiv.style.fontSize   = `${fs}px`;
      contentDiv.style.fontWeight = fw;
      contentDiv.style.fontStyle  = fi;
      contentDiv.style.fontFamily = ff ? resolveCssFontStack(ff) : 'Helvetica,Arial,sans-serif';
      contentDiv.style.textAlign  = align;
    };

    if (editMode) {
      // fit-content expands naturally; min-width honours any explicit user resize.
      const pc = this.pageContainers[block.pageIndex];
      wrapper.style.width     = 'fit-content';
      wrapper.style.minWidth  = `${w}px`;
      wrapper.style.maxWidth  = pc ? `${pc.clientWidth - x}px` : 'none';
      wrapper.style.minHeight = blockStyle?.cssHeight ? `${blockStyle.cssHeight}px` : '';

      wrapper.style.border        = '1px dashed rgba(59,130,246,0.6)';
      wrapper.style.cursor        = 'move';
      wrapper.style.pointerEvents = 'auto';
      if (contentDiv) {
        const { bg, textColor } = this.sampleBlockBg(block);
        if (hasMoved) this.eraseBlockOnCanvas(block, bg);
        contentDiv.style.width         = 'auto';
        contentDiv.style.cursor        = 'default';
        contentDiv.style.background    = bg;
        contentDiv.style.color         = color || textColor;
        contentDiv.style.pointerEvents = 'auto';
        applyContentStyle();
      }
      if (grabBar)      grabBar.style.height      = `${ANN.GRAB_BAR_H}px`;
      if (rightHandle)  rightHandle.style.display  = 'block';
      if (bottomHandle) bottomHandle.style.display = 'block';

    } else if (showOverlay) {
      wrapper.style.width     = `${w}px`;
      wrapper.style.minWidth  = '';
      wrapper.style.maxWidth  = '';
      wrapper.style.minHeight = blockStyle?.cssHeight ? `${blockStyle.cssHeight}px` : '';

      wrapper.style.border        = 'none';
      wrapper.style.cursor        = 'default';
      wrapper.style.pointerEvents = 'none';
      if (contentDiv) {
        const { bg, textColor } = this.sampleBlockBg(block);
        if (hasMoved) this.eraseBlockOnCanvas(block, bg);
        contentDiv.style.width         = '100%';
        contentDiv.contentEditable     = 'false';
        contentDiv.style.cursor        = 'default';
        contentDiv.style.background    = bg;
        contentDiv.style.color         = color || textColor;
        contentDiv.style.pointerEvents = 'none';
        applyContentStyle();
      }
      if (grabBar)      grabBar.style.height      = '0';
      if (rightHandle)  rightHandle.style.display  = 'none';
      if (bottomHandle) bottomHandle.style.display = 'none';

    } else {
      this.restoreBlockOnCanvas(block);
      wrapper.style.width     = `${w}px`;
      wrapper.style.minWidth  = '';
      wrapper.style.maxWidth  = '';
      wrapper.style.minHeight = '';

      wrapper.style.border        = 'none';
      wrapper.style.cursor        = 'default';
      wrapper.style.pointerEvents = 'none';
      if (contentDiv) {
        contentDiv.style.width         = '100%';
        contentDiv.contentEditable     = 'false';
        contentDiv.style.cursor        = 'default';
        contentDiv.style.background    = 'transparent';
        contentDiv.style.color         = 'transparent';
        contentDiv.style.pointerEvents = 'none';
      }
      if (grabBar)      grabBar.style.height      = '0';
      if (rightHandle)  rightHandle.style.display  = 'none';
      if (bottomHandle) bottomHandle.style.display = 'none';
    }
  }

  // ── Block drag & resize ─────────────────────────────────────────────────────

  private wireBlockDrag(
    triggerEl:   HTMLElement,
    wrapper:     HTMLDivElement,
    block:       PdfTextBlock,
    container:   HTMLElement,
    ...excludeEls: Array<HTMLElement | undefined>
  ): void {
    triggerEl.addEventListener('mousedown', (e: MouseEvent) => {
      if (!this.editable) return;
      for (const ex of excludeEls) {
        if (ex && (e.target === ex || ex.contains(e.target as Node))) return;
      }
      e.preventDefault(); e.stopPropagation();

      const cur   = this.blockStylesMap.get(block.id) ?? {};
      const origX = cur.cssX ?? block.cssX;
      const origY = cur.cssY ?? block.cssY;
      const sx = e.clientX, sy = e.clientY;
      let moved = false;

      const onMove = (me: MouseEvent) => {
        const z = this.zoom(container);
        const dx = (me.clientX - sx) / z, dy = (me.clientY - sy) / z;
        if (!moved && (Math.abs(dx) > 2 || Math.abs(dy) > 2)) moved = true;
        wrapper.style.left = `${origX + dx}px`;
        wrapper.style.top  = `${origY + dy}px`;
      };

      const onUp = (me: MouseEvent) => {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup',   onUp);
        if (!moved) return;
        const z    = this.zoom(container);
        const newX = Math.max(0, origX + (me.clientX - sx) / z);
        const newY = Math.max(0, origY + (me.clientY - sy) / z);
        const old  = { ...cur };
        const upd: PdfBlockStyle  = { ...cur, cssX: newX, cssY: newY };
        this.blockStylesMap.set(block.id, upd);
        this.onCommit([{ kind: 'styleBlock', blockId: block.id, pageIndex: block.pageIndex, oldStyle: old, newStyle: upd }]);
      };

      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup',   onUp);
    });
  }

  // ── Shared formatting toolbar ───────────────────────────────────────────────

  /**
   * Build and show the formatting toolbar.
   * When a formattingSlot is provided the toolbar is rendered inline inside it
   * (scrollable in the viewer's top bar). Otherwise it floats above the element.
   *
   * All controls use mousedown+preventDefault to avoid focus loss from the
   * contenteditable — EXCEPT the native color picker (stopPropagation only).
   *
   * Returns a cleanup fn that removes the toolbar and any open font picker.
   */
  private showToolbar(
    wrapperEl:    HTMLElement,
    state:        ElementStyle,
    contentDiv:   HTMLDivElement,
    onPickerColor: (color: string) => void,
  ): () => void {
    const inSlot = !!this.formattingSlot;

    const toolbar = document.createElement('div');
    toolbar.setAttribute('data-ann-toolbar', '');

    if (inSlot) {
      toolbar.style.cssText = [
        'display:flex', 'align-items:center', 'flex-wrap:nowrap', 'gap:2px',
        'padding:2px 4px', 'user-select:none',
      ].join(';');
    } else {
      toolbar.style.cssText = [
        'position:fixed', 'z-index:10000',
        'display:flex', 'align-items:center', 'flex-wrap:nowrap', 'gap:2px',
        'padding:5px 8px',
        'background:#111827',
        'border:1px solid rgba(255,255,255,0.1)',
        'border-radius:7px',
        'box-shadow:0 6px 20px rgba(0,0,0,0.5)',
        'user-select:none',
      ].join(';');
    }

    toolbar.addEventListener('mousedown', (e) => e.preventDefault());

    const sep = () => {
      const d = document.createElement('div');
      d.style.cssText = 'width:1px;height:16px;background:rgba(255,255,255,0.12);margin:0 4px;flex-shrink:0;';
      return d;
    };

    // ── Font family picker button ──────────────────────────────────────────────
    let pickerCleanup: (() => void) | null = null;

    const fontBtn = this.mkBtn(state.fontFamily, () => {
      if (pickerCleanup) { pickerCleanup(); pickerCleanup = null; return; }
      pickerCleanup = showFontPicker(fontBtn, state.fontFamily, (font) => {
        state.fontFamily = font.name;
        contentDiv.style.fontFamily = font.cssStack;
        this.styleDefaults.fontFamily = font.name;
        fontBtn.textContent = font.name;
        fontBtn.style.fontFamily = font.cssStack;
        pickerCleanup = null;
      });
    });
    fontBtn.style.fontFamily  = resolveCssFontStack(state.fontFamily);
    fontBtn.style.maxWidth    = '120px';
    fontBtn.style.minWidth    = '60px';
    fontBtn.style.overflow    = 'hidden';
    fontBtn.style.textOverflow = 'ellipsis';
    fontBtn.style.whiteSpace  = 'nowrap';
    fontBtn.style.textAlign   = 'left';
    toolbar.appendChild(fontBtn);

    toolbar.appendChild(sep());

    // ── Bold / Italic ──────────────────────────────────────────────────────────
    const boldBtn = this.mkBtn('B', () => {
      state.fontWeight = state.fontWeight === 'bold' ? 'normal' : 'bold';
      contentDiv.style.fontWeight = state.fontWeight;
      this.styleDefaults.fontWeight = state.fontWeight;
      updateStyleBtns();
    });
    boldBtn.style.fontWeight = 'bold';

    const italicBtn = this.mkBtn('I', () => {
      state.fontStyle = state.fontStyle === 'italic' ? 'normal' : 'italic';
      contentDiv.style.fontStyle = state.fontStyle;
      this.styleDefaults.fontStyle = state.fontStyle;
      updateStyleBtns();
    });
    italicBtn.style.fontStyle = 'italic';

    const updateStyleBtns = () => {
      boldBtn.style.background   = state.fontWeight === 'bold'   ? 'rgba(96,165,250,0.25)' : 'transparent';
      boldBtn.style.color        = state.fontWeight === 'bold'   ? '#93c5fd' : '#d1d5db';
      italicBtn.style.background = state.fontStyle  === 'italic' ? 'rgba(96,165,250,0.25)' : 'transparent';
      italicBtn.style.color      = state.fontStyle  === 'italic' ? '#93c5fd' : '#d1d5db';
    };
    updateStyleBtns();

    toolbar.appendChild(boldBtn);
    toolbar.appendChild(italicBtn);
    toolbar.appendChild(sep());

    // ── Alignment ──────────────────────────────────────────────────────────────
    const alignments: Array<{ key: 'left' | 'center' | 'right'; icon: string }> = [
      { key: 'left',   icon: '⫤' },
      { key: 'center', icon: '≡' },
      { key: 'right',  icon: '⫥' },
    ];
    for (const al of alignments) {
      const btn = this.mkBtn(al.icon, () => {
        state.textAlign = al.key;
        contentDiv.style.textAlign = al.key;
        this.styleDefaults.textAlign = al.key;
        updateAlignBtns();
      });
      btn.dataset.alignKey = al.key;
      toolbar.appendChild(btn);
    }
    const updateAlignBtns = () => {
      for (const b of toolbar.querySelectorAll<HTMLButtonElement>('[data-align-key]')) {
        const active = b.dataset.alignKey === state.textAlign;
        b.style.background = active ? 'rgba(96,165,250,0.25)' : 'transparent';
        b.style.color      = active ? '#93c5fd' : '#d1d5db';
      }
    };
    updateAlignBtns();

    toolbar.appendChild(sep());

    // ── Font size ──────────────────────────────────────────────────────────────
    const sizeEl = document.createElement('span');
    sizeEl.style.cssText = 'min-width:24px;text-align:center;font-size:12px;color:#d1d5db;font-family:monospace;';
    sizeEl.textContent   = String(state.fontSize);

    const sizeBtn = (delta: number, symbol: string) => {
      const btn = this.mkBtn(symbol, () => {
        state.fontSize = Math.max(ANN.MIN_FONT_SIZE, Math.min(ANN.MAX_FONT_SIZE, state.fontSize + delta));
        contentDiv.style.fontSize = `${state.fontSize}px`;
        sizeEl.textContent = String(state.fontSize);
      });
      btn.style.border = '1px solid rgba(255,255,255,0.15)';
      return btn;
    };

    toolbar.appendChild(sizeBtn(-1, '−'));
    toolbar.appendChild(sizeEl);
    toolbar.appendChild(sizeBtn(+1, '+'));
    toolbar.appendChild(sep());

    // ── Color palette ──────────────────────────────────────────────────────────
    const updateSwatches = (active: string) => {
      for (const b of toolbar.querySelectorAll<HTMLButtonElement>('[data-color]')) {
        b.style.outline       = b.dataset.color === active ? '2px solid #60a5fa' : '2px solid transparent';
        b.style.outlineOffset = '1px';
      }
    };

    for (const { hex, label } of PALETTE) {
      const swatch = document.createElement('button');
      swatch.dataset.color  = hex;
      swatch.title          = label;
      swatch.style.cssText  = [
        'width:15px', 'height:15px',
        `background:${hex}`,
        hex === '#ffffff' ? 'border:1px solid rgba(255,255,255,0.3)' : 'border:none',
        'border-radius:2px', 'cursor:pointer', 'flex-shrink:0', 'padding:0',
      ].join(';');
      swatch.addEventListener('mousedown', (e) => {
        e.preventDefault();
        state.color = hex;
        contentDiv.style.color = hex;
        this.styleDefaults.color = hex;
        updateSwatches(hex);
      });
      toolbar.appendChild(swatch);
    }

    // ── Custom color picker ────────────────────────────────────────────────────
    const colorInput = document.createElement('input');
    colorInput.type  = 'color';
    colorInput.value = state.color.startsWith('#') ? state.color : '#000000';
    colorInput.title = 'Custom color';
    colorInput.style.cssText = [
      'width:22px', 'height:22px', 'border:none',
      'border-radius:3px', 'cursor:pointer', 'padding:0',
      'background:transparent', 'flex-shrink:0',
    ].join(';');
    colorInput.addEventListener('mousedown', (e) => e.stopPropagation());
    colorInput.addEventListener('input', () => {
      const c = colorInput.value;
      state.color = c;
      contentDiv.style.color = c;
      this.styleDefaults.color = c;
      updateSwatches(c);
      onPickerColor(c);
    });
    toolbar.appendChild(colorInput);

    updateSwatches(state.color);

    if (inSlot) {
      this.formattingSlot!.innerHTML = '';
      this.formattingSlot!.appendChild(toolbar);
    } else {
      document.body.appendChild(toolbar);
      // Position above the wrapper; fall back to below if no room.
      const anchor   = wrapperEl.getBoundingClientRect();
      const tbRect   = toolbar.getBoundingClientRect();
      const topAbove = anchor.top - tbRect.height - 6;
      toolbar.style.left = `${Math.max(8, Math.min(anchor.left, window.innerWidth - tbRect.width - 8))}px`;
      toolbar.style.top  = `${topAbove >= 8 ? topAbove : anchor.bottom + 6}px`;
    }

    return () => {
      pickerCleanup?.();
      toolbar.remove();
      if (inSlot && this.formattingSlot) this.formattingSlot.innerHTML = '';
    };
  }

  private mkBtn(label: string, cb: () => void): HTMLButtonElement {
    const btn = document.createElement('button');
    btn.textContent = label;
    btn.style.cssText = [
      'min-width:26px', 'height:24px', 'padding:0 4px',
      'border:none', 'border-radius:3px', 'cursor:pointer',
      'background:transparent', 'color:#d1d5db',
      'font-size:12px', 'font-family:Helvetica,Arial,sans-serif',
      'flex-shrink:0',
    ].join(';');
    btn.addEventListener('mousedown', (e) => { e.preventDefault(); cb(); });
    return btn;
  }

  // ── Annotation element builders ─────────────────────────────────────────────

  private contentOf(wrapper: HTMLDivElement): HTMLDivElement | null {
    return wrapper.querySelector<HTMLDivElement>('[data-ann-content]');
  }

  private buildWrapper(ann: PdfAnnotation): {
    wrapper: HTMLDivElement; grabBar: HTMLDivElement; contentDiv: HTMLDivElement;
    rightHandle: HTMLDivElement; bottomHandle: HTMLDivElement;
  } {
    const wrapper = document.createElement('div');
    wrapper.dataset.annotationId = ann.id;
    wrapper.style.cssText = [
      'position:absolute',
      `left:${ann.cssX}px`, `top:${ann.cssY}px`,
      'box-sizing:border-box', 'z-index:3',
    ].join(';');

    const grabBar = document.createElement('div');
    grabBar.dataset.annGrab = '';
    grabBar.style.cssText = [
      'display:flex', 'align-items:center', 'justify-content:center',
      'overflow:hidden', 'height:0', 'cursor:move',
      'border-radius:4px 4px 0 0',
      'font-size:10px', 'letter-spacing:2px',
      'color:rgba(147,197,253,0.7)',
      'background:rgba(59,130,246,0.12)',
      'user-select:none', 'transition:background 0.1s',
    ].join(';');
    grabBar.textContent = '• • • • • • •';
    wrapper.appendChild(grabBar);

    const contentDiv = document.createElement('div');
    contentDiv.dataset.annContent = '';
    contentDiv.style.cssText = [
      'width:100%', 'min-height:1.4em', 'height:auto',
      `font-size:${ann.fontSize}px`,
      `font-weight:${ann.fontWeight}`,
      `font-style:${ann.fontStyle}`,
      `font-family:${resolveCssFontStack(ann.fontFamily)}`,
      `text-align:${ann.textAlign}`,
      'line-height:1.4', 'white-space:pre-wrap', 'word-break:break-word',
      'box-sizing:border-box', 'padding:3px 5px',
      'background:transparent', `color:${ann.color}`,
      'outline:none', 'border:none', 'cursor:default',
    ].join(';');
    contentDiv.innerText = ann.text;
    wrapper.appendChild(contentDiv);

    const rightHandle = document.createElement('div');
    rightHandle.dataset.annResizeRight = '';
    rightHandle.style.cssText = [
      'position:absolute', 'top:0', 'right:0',
      'width:6px', 'height:100%', 'cursor:ew-resize',
      'z-index:5', 'display:none',
      'background:rgba(59,130,246,0.3)',
      'border-radius:0 2px 2px 0',
    ].join(';');
    wrapper.appendChild(rightHandle);

    const bottomHandle = document.createElement('div');
    bottomHandle.dataset.annResizeBottom = '';
    bottomHandle.style.cssText = [
      'position:absolute', 'bottom:0', 'left:0',
      'width:100%', 'height:6px', 'cursor:ns-resize',
      'z-index:5', 'display:none',
      'background:rgba(59,130,246,0.3)',
      'border-radius:0 0 2px 2px',
    ].join(';');
    wrapper.appendChild(bottomHandle);

    return { wrapper, grabBar, contentDiv, rightHandle, bottomHandle };
  }

  private applyWrapperStyle(wrapper: HTMLDivElement, ann: PdfAnnotation, editMode: boolean): void {
    const contentDiv  = this.contentOf(wrapper);
    const grabBar     = wrapper.querySelector<HTMLDivElement>('[data-ann-grab]');
    const rightHandle  = wrapper.querySelector<HTMLDivElement>('[data-ann-resize-right]');
    const bottomHandle = wrapper.querySelector<HTMLDivElement>('[data-ann-resize-bottom]');

    wrapper.style.left = `${ann.cssX}px`;
    wrapper.style.top  = `${ann.cssY}px`;

    if (contentDiv) {
      contentDiv.style.fontSize   = `${ann.fontSize}px`;
      contentDiv.style.fontWeight = ann.fontWeight;
      contentDiv.style.fontStyle  = ann.fontStyle;
      contentDiv.style.fontFamily = resolveCssFontStack(ann.fontFamily);
      contentDiv.style.textAlign  = ann.textAlign;
      contentDiv.style.color      = ann.color;
    }

    if (editMode) {
      const pc = this.pageContainers[ann.pageIndex];
      // fit-content auto-expands with content; min-width honours any explicit user resize.
      wrapper.style.width     = 'fit-content';
      wrapper.style.minWidth  = `${ann.width}px`;
      wrapper.style.maxWidth  = pc ? `${pc.clientWidth - ann.cssX}px` : 'none';
      wrapper.style.minHeight = ann.height ? `${ann.height}px` : '';
      wrapper.style.border        = '1px dashed rgba(59,130,246,0.6)';
      wrapper.style.cursor        = 'move';
      wrapper.style.pointerEvents = 'auto';
      if (grabBar)      grabBar.style.height      = `${ANN.GRAB_BAR_H}px`;
      if (rightHandle)  rightHandle.style.display  = 'block';
      if (bottomHandle) bottomHandle.style.display = 'block';
    } else {
      wrapper.style.width     = `${ann.width}px`;
      wrapper.style.minWidth  = '';
      wrapper.style.maxWidth  = '';
      wrapper.style.minHeight = ann.height ? `${ann.height}px` : '';
      wrapper.style.border        = 'none';
      wrapper.style.cursor        = 'default';
      wrapper.style.pointerEvents = 'none';
      if (grabBar)      grabBar.style.height      = '0';
      if (rightHandle)  rightHandle.style.display  = 'none';
      if (bottomHandle) bottomHandle.style.display = 'none';
    }
  }

  // ── New annotation (transient) ──────────────────────────────────────────────

  private startNewAnnotation(
    pageIndex: number, cssX: number, cssY: number,
    container: HTMLElement, overlayEl: HTMLDivElement,
    annotationEls: Map<string, HTMLDivElement>,
  ): void {
    const id = `ann-${pageIndex}-${++this.annCounter}`;
    const ann: PdfAnnotation = {
      id, pageIndex, cssX, cssY,
      width:      ANN.DEFAULT_WIDTH,
      fontSize:   ANN.DEFAULT_FONT_SIZE,
      text:       '',
      color:      this.styleDefaults.color,
      fontWeight: this.styleDefaults.fontWeight,
      fontStyle:  this.styleDefaults.fontStyle,
      fontFamily: this.styleDefaults.fontFamily,
      textAlign:  this.styleDefaults.textAlign,
    };

    const { wrapper, grabBar, contentDiv, rightHandle, bottomHandle } = this.buildWrapper(ann);

    const state: ElementStyle = {
      color: ann.color, fontSize: ann.fontSize, fontWeight: ann.fontWeight,
      fontStyle: ann.fontStyle, fontFamily: ann.fontFamily, textAlign: ann.textAlign,
    };

    const pc = this.pageContainers[pageIndex];
    grabBar.style.height       = `${ANN.GRAB_BAR_H}px`;
    wrapper.style.width        = 'fit-content';
    wrapper.style.minWidth     = `${ANN.MIN_WIDTH}px`;
    wrapper.style.maxWidth     = pc ? `${pc.clientWidth - cssX}px` : 'none';
    wrapper.style.border       = '1px dashed rgba(59,130,246,0.8)';
    wrapper.style.cursor       = 'move';
    wrapper.style.pointerEvents = 'auto';
    rightHandle.style.display  = 'block';
    bottomHandle.style.display = 'block';
    contentDiv.contentEditable = 'true';
    contentDiv.style.cursor    = 'text';

    let removeToolbar: (() => void) | null = null;

    contentDiv.addEventListener('focus', () => {
      removeToolbar?.();
      removeToolbar = this.showToolbar(wrapper, state, contentDiv, () => { /* new ann — no extra op */ });
    });
    contentDiv.addEventListener('keydown', (e: KeyboardEvent) => {
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); contentDiv.blur(); return; }
      if (e.key === 'Escape') { e.preventDefault(); removeToolbar?.(); wrapper.remove(); return; }
      e.stopPropagation();
    });
    contentDiv.addEventListener('blur', () => {
      removeToolbar?.(); removeToolbar = null;
      const text = contentDiv.innerText.trim();
      if (!text) { wrapper.remove(); return; }

      ann.text = text; ann.color = state.color; ann.fontSize = state.fontSize;
      ann.fontWeight = state.fontWeight; ann.fontStyle = state.fontStyle;
      ann.fontFamily = state.fontFamily; ann.textAlign = state.textAlign;
      ann.width = Math.max(ANN.MIN_WIDTH, wrapper.offsetWidth || ann.width);
      Object.assign(this.styleDefaults, {
        color: state.color, fontWeight: state.fontWeight,
        fontStyle: state.fontStyle, fontFamily: state.fontFamily, textAlign: state.textAlign,
      });

      this.annotations.set(id, { ...ann });
      wrapper.remove();
      const committed = this.buildCommittedEl(ann, pageIndex, annotationEls);
      overlayEl.appendChild(committed);
      annotationEls.set(id, committed);
      this.onCommit([{ kind: 'addAnnotation', annotation: { ...ann } }]);
    });

    this.wireDrag(grabBar, wrapper, ann, container);
    this.wireDrag(wrapper, wrapper, ann, container, contentDiv, rightHandle, bottomHandle);

    // Resize handles — update ann directly (will be committed on blur).
    this.wireResize(rightHandle, wrapper, 'h', container, (newW) => { ann.width  = newW; });
    this.wireResize(bottomHandle, wrapper, 'v', container, (newH) => { ann.height = newH; });

    overlayEl.appendChild(wrapper);
    setTimeout(() => contentDiv.focus(), 0);
  }

  // ── Committed annotation element ────────────────────────────────────────────

  private buildCommittedEl(
    ann: PdfAnnotation, pageIndex: number,
    annotationEls: Map<string, HTMLDivElement>,
  ): HTMLDivElement {
    const container = this.pageContainers[pageIndex]!;
    const overlayEl = this.overlays.find(o => o.pageIndex === pageIndex)?.overlayEl;

    const { wrapper, grabBar, contentDiv, rightHandle, bottomHandle } = this.buildWrapper(ann);
    this.applyWrapperStyle(wrapper, ann, this.editable);

    const state: ElementStyle = {
      color: ann.color, fontSize: ann.fontSize, fontWeight: ann.fontWeight,
      fontStyle: ann.fontStyle, fontFamily: ann.fontFamily, textAlign: ann.textAlign,
    };

    let removeToolbar: (() => void) | null = null;

    wrapper.addEventListener('click', (e: MouseEvent) => {
      if (!this.editable) return;
      const t = e.target as HTMLElement;
      if (t.dataset.annGrab !== undefined) return;
      contentDiv.contentEditable = 'true';
      contentDiv.style.cursor    = 'text';
      contentDiv.focus();
    });

    contentDiv.addEventListener('focus', () => {
      removeToolbar?.();
      removeToolbar = this.showToolbar(wrapper, state, contentDiv, (pickerColor) => {
        const oldAnnotation = { ...ann };
        ann.color = pickerColor;
        const newAnnotation = { ...ann };
        this.annotations.set(ann.id, newAnnotation);
        if (contentDiv.isConnected) contentDiv.style.color = pickerColor;
        this.onCommit([{ kind: 'updateAnnotation', annotationId: ann.id, pageIndex: ann.pageIndex, oldAnnotation, newAnnotation }]);
      });
    });

    contentDiv.addEventListener('keydown', (e: KeyboardEvent) => {
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); contentDiv.blur(); return; }
      if (e.key === 'Escape') {
        e.preventDefault();
        contentDiv.innerText = ann.text;
        state.color = ann.color; state.fontSize = ann.fontSize;
        state.fontWeight = ann.fontWeight; state.fontStyle = ann.fontStyle;
        state.fontFamily = ann.fontFamily; state.textAlign = ann.textAlign;
        this.applyContentStyle(contentDiv, ann);
        contentDiv.contentEditable = 'false'; contentDiv.style.cursor = 'default';
        removeToolbar?.(); removeToolbar = null;
        this.applyWrapperStyle(wrapper, ann, this.editable);
        return;
      }
      e.stopPropagation();
    });

    contentDiv.addEventListener('blur', () => {
      removeToolbar?.(); removeToolbar = null;
      const newText = contentDiv.innerText.trim();

      if (!newText) {
        wrapper.remove(); annotationEls.delete(ann.id); this.annotations.delete(ann.id);
        this.onCommit([{ kind: 'removeAnnotation', annotationId: ann.id, pageIndex: ann.pageIndex, annotation: { ...ann } }]);
        return;
      }

      const newWidth = Math.max(ANN.MIN_WIDTH, wrapper.offsetWidth || ann.width);
      const changed = newText !== ann.text || state.color !== ann.color ||
        state.fontSize !== ann.fontSize || state.fontWeight !== ann.fontWeight ||
        state.fontStyle !== ann.fontStyle || state.fontFamily !== ann.fontFamily ||
        state.textAlign !== ann.textAlign || newWidth !== ann.width;

      if (changed) {
        const oldAnnotation = { ...ann };
        ann.text = newText; ann.color = state.color; ann.fontSize = state.fontSize;
        ann.fontWeight = state.fontWeight; ann.fontStyle = state.fontStyle;
        ann.fontFamily = state.fontFamily; ann.textAlign = state.textAlign;
        ann.width = newWidth;
        const newAnnotation = { ...ann };
        this.annotations.set(ann.id, newAnnotation);
        Object.assign(this.styleDefaults, {
          color: state.color, fontWeight: state.fontWeight,
          fontStyle: state.fontStyle, fontFamily: state.fontFamily, textAlign: state.textAlign,
        });
        this.onCommit([{ kind: 'updateAnnotation', annotationId: ann.id, pageIndex: ann.pageIndex, oldAnnotation, newAnnotation }]);
      }

      contentDiv.contentEditable = 'false'; contentDiv.style.cursor = 'default';
      this.applyWrapperStyle(wrapper, ann, this.editable);
    });

    if (overlayEl) {
      this.wireDrag(grabBar, wrapper, ann, container);
      this.wireDrag(wrapper, wrapper, ann, container, contentDiv, rightHandle, bottomHandle);
    }

    this.wireResize(rightHandle, wrapper, 'h', container, (newW) => {
      const oldAnn = { ...ann };
      ann.width = newW;
      const newAnn = { ...ann };
      this.annotations.set(ann.id, newAnn);
      this.onCommit([{ kind: 'updateAnnotation', annotationId: ann.id, pageIndex: ann.pageIndex, oldAnnotation: oldAnn, newAnnotation: newAnn }]);
    });

    this.wireResize(bottomHandle, wrapper, 'v', container, (newH) => {
      const oldAnn = { ...ann };
      ann.height = newH;
      const newAnn = { ...ann };
      this.annotations.set(ann.id, newAnn);
      this.onCommit([{ kind: 'updateAnnotation', annotationId: ann.id, pageIndex: ann.pageIndex, oldAnnotation: oldAnn, newAnnotation: newAnn }]);
    });

    return wrapper;
  }

  private applyContentStyle(contentDiv: HTMLDivElement, ann: PdfAnnotation): void {
    contentDiv.style.fontSize   = `${ann.fontSize}px`;
    contentDiv.style.fontWeight = ann.fontWeight;
    contentDiv.style.fontStyle  = ann.fontStyle;
    contentDiv.style.fontFamily = resolveCssFontStack(ann.fontFamily);
    contentDiv.style.textAlign  = ann.textAlign;
    contentDiv.style.color      = ann.color;
  }

  // ── Annotation drag & resize ────────────────────────────────────────────────

  private wireDrag(
    triggerEl:   HTMLElement,
    wrapper:     HTMLDivElement,
    ann:         PdfAnnotation,
    container:   HTMLElement,
    ...excludeEls: Array<HTMLElement | undefined>
  ): void {
    triggerEl.addEventListener('mousedown', (e: MouseEvent) => {
      if (!this.editable) return;
      for (const ex of excludeEls) {
        if (ex && (e.target === ex || ex.contains(e.target as Node))) return;
      }

      e.preventDefault(); e.stopPropagation();

      const origX = ann.cssX, origY = ann.cssY;
      const sx = e.clientX, sy = e.clientY;
      let moved = false;

      const onMove = (me: MouseEvent) => {
        const z = this.zoom(container);
        const dx = (me.clientX - sx) / z, dy = (me.clientY - sy) / z;
        if (!moved && (Math.abs(dx) > 2 || Math.abs(dy) > 2)) moved = true;
        wrapper.style.left = `${origX + dx}px`;
        wrapper.style.top  = `${origY + dy}px`;
      };
      const onUp = (me: MouseEvent) => {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup',   onUp);
        if (!moved) return;
        const z = this.zoom(container);
        const oldAnn = { ...ann };
        ann.cssX = Math.max(0, origX + (me.clientX - sx) / z);
        ann.cssY = Math.max(0, origY + (me.clientY - sy) / z);
        const newAnn = { ...ann };
        this.annotations.set(ann.id, newAnn);
        this.onCommit([{ kind: 'updateAnnotation', annotationId: ann.id, pageIndex: ann.pageIndex, oldAnnotation: oldAnn, newAnnotation: newAnn }]);
      };
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup',   onUp);
    });
  }

  /**
   * Wire a resize handle (right edge or bottom edge) to adjust the wrapper's
   * width or height. `onCommit` is called once on mouse-up with the new size.
   */
  private wireResize(
    handle:    HTMLDivElement,
    wrapper:   HTMLDivElement,
    direction: 'h' | 'v',
    container: HTMLElement,
    onCommit:  (newSize: number) => void,
  ): void {
    handle.addEventListener('mousedown', (e: MouseEvent) => {
      if (!this.editable) return;
      e.preventDefault(); e.stopPropagation();

      const origSize   = direction === 'h' ? wrapper.offsetWidth : wrapper.offsetHeight;
      const startCoord = direction === 'h' ? e.clientX : e.clientY;
      const minSize    = direction === 'h' ? ANN.MIN_WIDTH : ANN.MIN_HEIGHT;
      let moved = false;

      const onMove = (me: MouseEvent) => {
        const z     = this.zoom(container);
        const delta = ((direction === 'h' ? me.clientX : me.clientY) - startCoord) / z;
        const size  = Math.max(minSize, origSize + delta);
        if (!moved && Math.abs(delta) > 2) moved = true;
        if (direction === 'h') {
          wrapper.style.width    = `${size}px`;
          wrapper.style.minWidth = `${minSize}px`;
        } else {
          wrapper.style.minHeight = `${size}px`;
        }
      };

      const onUp = (me: MouseEvent) => {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup',   onUp);
        if (!moved) return;
        const z    = this.zoom(container);
        const delta = ((direction === 'h' ? me.clientX : me.clientY) - startCoord) / z;
        onCommit(Math.max(minSize, origSize + delta));
      };

      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup',   onUp);
    });
  }

  // ── Utility ─────────────────────────────────────────────────────────────────

  private zoom(container: HTMLElement): number {
    return parseFloat((container.style as CSSStyleDeclaration & { zoom?: string }).zoom ?? '1') || 1;
  }

  // ── Canvas sampling ─────────────────────────────────────────────────────────

  private sampleBlockBg(block: PdfTextBlock): { bg: string; textColor: string } {
    const cached = this.bgCache.get(block.id);
    if (cached) return cached;
    const result = this.sampleCanvasAt(
      block.pageIndex,
      block.cssX + block.cssWidth * 0.5, block.cssY - 3,
      block.cssX + block.cssWidth * 0.5, block.cssY + block.cssHeight + 3,
    );
    this.bgCache.set(block.id, result);
    return result;
  }

  private sampleCanvasAt(pageIndex: number, ...pts: number[]): { bg: string; textColor: string } {
    const fb = { bg: '#ffffff', textColor: '#000000' };
    const container = this.pageContainers[pageIndex];
    const canvas    = container?.querySelector<HTMLCanvasElement>('canvas');
    const ctx       = canvas?.getContext('2d');
    if (!ctx || !canvas) return fb;

    const cssW = parseFloat(canvas.style.width)  || canvas.width;
    const cssH = parseFloat(canvas.style.height) || canvas.height;
    const sx   = canvas.width  / cssW;
    const sy   = canvas.height / cssH;
    let r = 0, g = 0, b = 0, n = 0;

    for (let i = 0; i + 1 < pts.length; i += 2) {
      const cx = pts[i]!, cy = pts[i + 1]!;
      if (cx < 0 || cy < 0 || cx >= cssW || cy >= cssH) continue;
      const px = Math.max(0, Math.min(canvas.width  - 1, Math.floor(cx * sx)));
      const py = Math.max(0, Math.min(canvas.height - 1, Math.floor(cy * sy)));
      try { const d = ctx.getImageData(px, py, 1, 1).data; r += d[0]!; g += d[1]!; b += d[2]!; n++; }
      catch { return fb; }
    }
    if (n === 0) return fb;
    return {
      bg:        `rgb(${Math.round(r/n)},${Math.round(g/n)},${Math.round(b/n)})`,
      textColor: relativeLuminance(r/n, g/n, b/n) > 0.179 ? '#000000' : '#ffffff',
    };
  }

  // ── Canvas overdraw ─────────────────────────────────────────────────────────

  private eraseBlockOnCanvas(block: PdfTextBlock, bgColor: string): void {
    const container = this.pageContainers[block.pageIndex];
    const canvas    = container?.querySelector<HTMLCanvasElement>('canvas');
    const ctx       = canvas?.getContext('2d');
    if (!ctx || !canvas) return;

    const cssW = parseFloat(canvas.style.width)  || canvas.width;
    const cssH = parseFloat(canvas.style.height) || canvas.height;
    const scX  = canvas.width  / cssW;
    const scY  = canvas.height / cssH;

    const m  = 4;
    const x  = Math.max(0, Math.floor((block.cssX - m) * scX));
    const y  = Math.max(0, Math.floor((block.cssY - m) * scY));
    const x2 = Math.min(canvas.width,  Math.ceil((block.cssX + block.cssWidth  + m) * scX));
    const y2 = Math.min(canvas.height, Math.ceil((block.cssY + block.cssHeight + m) * scY));
    const w  = x2 - x, h = y2 - y;
    if (w <= 0 || h <= 0) return;

    if (!this.savedPixels.has(block.id)) {
      try { this.savedPixels.set(block.id, { imageData: ctx.getImageData(x, y, w, h), x, y }); }
      catch { return; }
    }
    ctx.save(); ctx.fillStyle = bgColor; ctx.fillRect(x, y, w, h); ctx.restore();
  }

  private restoreBlockOnCanvas(block: PdfTextBlock): void {
    const saved = this.savedPixels.get(block.id);
    if (!saved) return;
    const container = this.pageContainers[block.pageIndex];
    const canvas    = container?.querySelector<HTMLCanvasElement>('canvas');
    const ctx       = canvas?.getContext('2d');
    if (!ctx) return;
    try { ctx.putImageData(saved.imageData, saved.x, saved.y); } catch { /* tainted */ }
    this.savedPixels.delete(block.id);
    this.bgCache.delete(block.id);
  }
}
