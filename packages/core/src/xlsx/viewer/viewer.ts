/**
 * Viewer controller: mounts a {@link Workbook} into a container as a grid with
 * a sheet-tab bar, and — when `editable` — owns the edit session that turns
 * what the user types into worksheet XML.
 *
 * The commit cycle mirrors the other formats: the grid reports the raw text of
 * a cell, the session writes it and re-parses the sheet, and the viewer renders
 * the sheet again from the committed model, restoring the selection. So what
 * the user is looking at is always the XML, never the browser's improvisation.
 */
import type { Workbook } from '../workbook/workbook.js';
import type { XlsxCell, XlsxSheet } from '../model.js';
import { createSheetView, type XlsxSheetView } from '../render/dom.js';
import { XlsxEditSession } from '../edit/session.js';
import { readCellPatch } from '../edit/format.js';
import type { CellFormatPatch } from '../edit/styleWrite.js';

export interface XlsxViewerOptions {
  /** Initial sheet name or 0-indexed position (defaults to 0) */
  initialSheet?: number | string;
  /** Callback fired when active sheet changes */
  onSheetChange?: (sheetName: string, index: number) => void;
  /**
   * Edit cell values and formatting in place. Cells on a protected sheet, cells
   * covered by a merge, and array/shared-formula hosts stay read-only. Off by
   * default.
   */
  editable?: boolean;
  /** Called after a committed edit, undo or redo. */
  onEdit?: () => void;
  /** Called when the selection moves, with the formatting in effect there. */
  onSelectionChange?: (format: CellFormatPatch, ref: string) => void;
}

export class XlsxViewer {
  private readonly workbook: Workbook;
  private readonly container: HTMLElement;
  private readonly options: XlsxViewerOptions;
  private activeSheetIndex = 0;
  private editable: boolean;
  private session: XlsxEditSession | null = null;
  private view: XlsxSheetView | null = null;
  private activeRef = 'A1';

  constructor(workbook: Workbook, container: HTMLElement, options: XlsxViewerOptions = {}) {
    this.workbook = workbook;
    this.container = container;
    this.options = options;
    this.editable = options.editable ?? false;
    if (this.editable) this.session = this.newSession();
    this.init();
  }

  static create(workbook: Workbook, container: HTMLElement, options?: XlsxViewerOptions): XlsxViewer {
    return new XlsxViewer(workbook, container, options);
  }

  goToSheet(key: number | string): void {
    let idx = 0;
    if (typeof key === 'number') {
      idx = Math.max(0, Math.min(key, this.workbook.sheetSummaries.length - 1));
    } else {
      const found = this.workbook.sheetSummaries.findIndex((s) => s.name === key || s.id === key);
      if (found !== -1) idx = found;
    }
    if (idx === this.activeSheetIndex) return;
    this.view?.commitEdit();
    this.activeSheetIndex = idx;
    this.activeRef = 'A1';
    this.render();
  }

  /** The sheet currently on screen. */
  get sheet(): XlsxSheet | undefined {
    return this.workbook.getSheet(this.activeSheetIndex);
  }

  /** The cell the selection is on. */
  get activeCell(): XlsxCell | undefined {
    return this.session && this.sheet
      ? this.session.contextFor(this.sheet).cellAt(this.activeRef)
      : undefined;
  }

  // ─── Editing ───────────────────────────────────────────────────────────────

  /** Turn editing on or off, re-rendering into the matching grid. */
  setEditable(on: boolean): void {
    if (on === this.editable) return;
    if (this.editable) this.view?.commitEdit();
    this.editable = on;
    this.session ??= this.newSession();
    this.render();
  }

  /** Format the selected cells (editable mode only). */
  applyFormat(patch: CellFormatPatch): boolean {
    const sheet = this.sheet;
    if (!this.editable || !this.session || !sheet || !this.view) return false;
    const refs = this.view.selectedRefs();
    const rebuilt = this.session.applyFormat(sheet, refs, patch);
    if (!rebuilt) return false;
    this.render();
    return true;
  }

  /** Commit the cell being typed into, without waiting for it to lose focus. */
  commitActive(): boolean {
    return this.view?.commitEdit() ?? false;
  }

  get canUndo(): boolean {
    return this.session?.canUndo ?? false;
  }

  get canRedo(): boolean {
    return this.session?.canRedo ?? false;
  }

  /** True if the workbook has unsaved edits. */
  get hasEdits(): boolean {
    return this.session?.hasEdits ?? false;
  }

  undo(): void {
    if (this.session?.undo()) this.render();
  }

  redo(): void {
    if (this.session?.redo()) this.render();
  }

  /** Re-zip the workbook, edits included, as a .xlsx Blob. */
  exportBlob(): Blob {
    // Flush anything still being typed before packaging.
    this.commitActive();
    return this.workbook.exportBlob();
  }

  destroy(): void {
    this.view?.commitEdit();
    this.container.innerHTML = '';
    this.view = null;
  }

  private newSession(): XlsxEditSession {
    return new XlsxEditSession(this.workbook, { onChange: () => this.options.onEdit?.() });
  }

  private init(): void {
    if (this.options.initialSheet !== undefined) {
      if (typeof this.options.initialSheet === 'number') {
        this.activeSheetIndex = this.options.initialSheet;
      } else {
        const found = this.workbook.sheetSummaries.findIndex(
          (s) => s.name === this.options.initialSheet || s.id === this.options.initialSheet,
        );
        if (found !== -1) this.activeSheetIndex = found;
      }
    }
    this.render();
  }

  /** Write one cell and re-render from the committed model. */
  private commitCell(ref: string, text: string, nextRef?: string): void {
    const sheet = this.sheet;
    if (!this.session || !sheet) return;
    const rebuilt = this.session.commitCell(sheet, ref, text);
    if (rebuilt) {
      this.activeRef = nextRef ?? ref;
      this.render();
      this.view?.focusGrid();
      return;
    }
    // Nothing changed (unchanged text, or a read-only cell): the grid is still
    // valid, so just move the selection where the user was heading.
    if (nextRef) this.view?.select(nextRef);
  }

  private clearCells(refs: string[]): void {
    const sheet = this.sheet;
    if (!this.session || !sheet) return;
    if (this.session.clearCells(sheet, refs)) {
      this.render();
      this.view?.focusGrid();
    }
  }

  private render(): void {
    this.container.innerHTML = '';

    const outerContainer = document.createElement('div');
    outerContainer.style.cssText = `
      display: flex;
      flex-direction: column;
      width: 100%;
      height: 100%;
      position: relative;
      background: #f3f4f6;
    `;

    const sheet = this.sheet;
    if (sheet) {
      const session = this.editable ? this.session : null;
      this.view = createSheetView(sheet, {
        activeCellRef: this.activeRef,
        ...(session ? { edit: session.contextFor(sheet) } : {}),
        onCommit: (ref, text, nextRef) => this.commitCell(ref, text, nextRef),
        onClearCells: (refs) => this.clearCells(refs),
        onUndo: () => this.undo(),
        onRedo: () => this.redo(),
        onSelectionChange: (_refs, ref) => {
          this.activeRef = ref;
          this.options.onSelectionChange?.(
            readCellPatch(session?.contextFor(sheet).cellAt(ref)),
            ref,
          );
        },
      });
      const sheetDom = this.view.root;
      sheetDom.style.flex = '1';
      sheetDom.style.minHeight = '0';
      outerContainer.appendChild(sheetDom);
      // Report where the selection landed, so a toolbar starts in sync.
      this.options.onSelectionChange?.(
        readCellPatch(session?.contextFor(sheet).cellAt(this.activeRef)),
        this.activeRef,
      );
    } else {
      this.view = null;
      const emptyMsg = document.createElement('div');
      emptyMsg.style.cssText = 'margin: auto; color: #6b7280; font-size: 14px;';
      emptyMsg.textContent = 'No sheet data available.';
      outerContainer.appendChild(emptyMsg);
    }

    // Render Bottom Sheet Tabs Bar
    if (this.workbook.sheetSummaries.length > 0) {
      const tabBar = document.createElement('div');
      tabBar.className = 'cbx-xlsx-tab-bar';
      tabBar.style.cssText = `
        display: flex;
        align-items: center;
        gap: 2px;
        padding: 4px 8px 0 8px;
        background: #e5e7eb;
        border-top: 1px solid #d1d5db;
        overflow-x: auto;
        flex-shrink: 0;
      `;

      this.workbook.sheetSummaries.forEach((sum, idx) => {
        const isActive = idx === this.activeSheetIndex;
        const tab = document.createElement('button');
        tab.textContent = sum.name;
        tab.style.cssText = `
          padding: 5px 14px;
          border: 1px solid ${isActive ? '#d1d5db' : 'transparent'};
          border-bottom: none;
          border-radius: 4px 4px 0 0;
          background: ${isActive ? '#ffffff' : '#f3f4f6'};
          color: ${isActive ? '#107c41' : '#4b5563'};
          font-weight: ${isActive ? '600' : 'normal'};
          font-size: 12px;
          cursor: pointer;
          white-space: nowrap;
          transition: background 0.15s, color 0.15s;
        `;

        tab.addEventListener('click', () => {
          if (this.activeSheetIndex !== idx) {
            this.goToSheet(idx);
            this.options.onSheetChange?.(sum.name, idx);
          }
        });

        tabBar.appendChild(tab);
      });

      outerContainer.appendChild(tabBar);
    }

    this.container.appendChild(outerContainer);
  }
}

export function createXlsxViewer(
  workbook: Workbook,
  container: HTMLElement,
  options?: XlsxViewerOptions,
): XlsxViewer {
  return XlsxViewer.create(workbook, container, options);
}
