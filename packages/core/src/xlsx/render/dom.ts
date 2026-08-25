/**
 * Model -> DOM for one worksheet: a sticky-header grid plus the formula bar.
 *
 * The grid is a real `<table>` rather than absolutely-positioned divs, because a
 * spreadsheet *is* a table: merged cells map onto rowspan/colspan, and column
 * widths and row heights come straight off the model.
 *
 * Editing is opt-in and follows the same rule as the other formats: the browser
 * is left to do the typing, and the text is read back out of the DOM when the
 * cell is committed. The unit is one cell, so `contentEditable` is switched on
 * for the cell being edited and off again the moment it commits — which is also
 * what makes arrow-key navigation possible between edits.
 */
import type { XlsxSheet, XlsxCell, XlsxMergeCell } from '../model.js';
import { indexToColAlpha, parseCellRef } from '../sheets/sheet.js';
import { editableText } from '../edit/values.js';
import type { XlsxEditRenderContext } from '../edit/context.js';

/** DOM attribute carrying a cell's reference, so events can be mapped back. */
export const XLSX_CELL_ATTR = 'data-cbx-cell';

const ACCENT = '#107c41';
const DEFAULT_COL_PX = 88;
const DEFAULT_ROW_PX = 22;
/** Blank rows/columns rendered past the data, so the sheet can be grown. */
const EDIT_HEADROOM_ROWS = 12;
const EDIT_HEADROOM_COLS = 4;

export interface RenderXlsxOptions {
  activeCellRef?: string;
  onSelectCell?: (cell: XlsxCell | undefined, ref: string) => void;
  /**
   * Enables in-place cell editing. The context decides which cells accept
   * typing; a protected sheet, a merged-over cell or an array-formula host does
   * not.
   */
  edit?: XlsxEditRenderContext;
  /**
   * Called with the raw text the user typed into a cell. `nextRef` is where the
   * caret was heading — Enter moves down, Tab moves right, a click moves to the
   * clicked cell — so the caller can restore the selection after the re-render
   * a commit triggers.
   */
  onCommit?: (ref: string, text: string, nextRef?: string) => void;
  /** Called when the user clears the selection with Delete/Backspace. */
  onClearCells?: (refs: string[]) => void;
  /** Called when the selected range changes. */
  onSelectionChange?: (refs: string[], activeRef: string) => void;
  onUndo?: () => void;
  onRedo?: () => void;
}

/** A mounted sheet grid, and the handful of things a viewer drives it with. */
export interface XlsxSheetView {
  readonly root: HTMLElement;
  /** The cell the caret is on. */
  readonly activeRef: string;
  /** Every reference in the current rectangular selection. */
  selectedRefs(): string[];
  /** Move (or, with `extend`, grow) the selection. */
  select(ref: string, extend?: boolean): void;
  /** Start typing into the active cell, optionally replacing it with `seed`. */
  beginEdit(seed?: string): void;
  /** Read the cell being edited back out and report it. */
  commitEdit(): boolean;
  /** Abandon the edit in progress, restoring what was rendered. */
  cancelEdit(): void;
  focusGrid(): void;
}

/**
 * Render a worksheet. Returns the root element; use {@link createSheetView}
 * when the caller also needs to drive selection and editing.
 */
export function renderXlsxSheet(sheet: XlsxSheet, options: RenderXlsxOptions = {}): HTMLElement {
  return createSheetView(sheet, options).root;
}

export function createSheetView(
  sheet: XlsxSheet,
  options: RenderXlsxOptions = {},
): XlsxSheetView {
  const editing = options.edit;
  const root = document.createElement('div');
  root.className = 'cbx-xlsx-wrapper';
  root.style.cssText = `
    display: flex;
    flex-direction: column;
    width: 100%;
    height: 100%;
    background: #ffffff;
    color: #1f1f1f;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
    font-size: 13px;
    user-select: text;
    overflow: hidden;
  `;

  // ── Formula bar ────────────────────────────────────────────────────────────
  const formulaBar = document.createElement('div');
  formulaBar.className = 'cbx-xlsx-formula-bar';
  formulaBar.style.cssText = `
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 6px 12px;
    background: #f8f9fa;
    border-bottom: 1px solid #e0e0e0;
    font-family: monospace;
    font-size: 12px;
    flex-shrink: 0;
  `;

  const nameBox = document.createElement('div');
  nameBox.className = 'cbx-xlsx-name-box';
  nameBox.style.cssText = `
    font-weight: 600;
    padding: 2px 8px;
    background: #ffffff;
    border: 1px solid #d1d5db;
    border-radius: 3px;
    min-width: 48px;
    text-align: center;
    color: ${ACCENT};
  `;

  const fxLabel = document.createElement('span');
  fxLabel.textContent = 'fx';
  fxLabel.style.cssText = 'color: #6b7280; font-style: italic; font-weight: bold;';

  const fxInput = document.createElement('div');
  fxInput.className = 'cbx-xlsx-fx-input';
  fxInput.style.cssText = `
    flex: 1;
    padding: 3px 8px;
    background: #ffffff;
    border: 1px solid #d1d5db;
    border-radius: 3px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    color: #111827;
    outline: none;
  `;

  formulaBar.appendChild(nameBox);
  formulaBar.appendChild(fxLabel);
  formulaBar.appendChild(fxInput);
  root.appendChild(formulaBar);

  // ── Grid ───────────────────────────────────────────────────────────────────
  const gridContainer = document.createElement('div');
  gridContainer.className = 'cbx-xlsx-grid-container';
  gridContainer.style.cssText = `
    flex: 1;
    min-height: 0;
    overflow: auto;
    position: relative;
    outline: none;
  `;
  if (editing) gridContainer.tabIndex = 0;

  const table = document.createElement('table');
  table.className = 'cbx-xlsx-table';
  table.style.cssText = `
    border-collapse: collapse;
    table-layout: fixed;
    background: #ffffff;
  `;

  const rowCount = sheet.maxRow + (editing ? EDIT_HEADROOM_ROWS : 0);
  const colCount = sheet.maxCol + (editing ? EDIT_HEADROOM_COLS : 0);

  // Merged cells: the origin spans the range, every other cell is not rendered.
  const mergeLookup = new Map<string, { merge: XlsxMergeCell; isOrigin: boolean }>();
  for (const m of sheet.mergeCells) {
    for (let r = m.startRow; r <= m.endRow; r++) {
      for (let c = m.startCol; c <= m.endCol; c++) {
        mergeLookup.set(`${r}:${c}`, {
          merge: m,
          isOrigin: r === m.startRow && c === m.startCol,
        });
      }
    }
  }

  const colWidthMap = new Map<number, number>();
  for (const colDef of sheet.cols) {
    for (let c = colDef.min - 1; c <= colDef.max - 1; c++) {
      if (colDef.widthPx) colWidthMap.set(c, colDef.widthPx);
    }
  }

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  const cornerTh = document.createElement('th');
  cornerTh.style.cssText = `
    width: 46px;
    min-width: 46px;
    height: 24px;
    background: #f3f4f6;
    border: 1px solid #d1d5db;
    position: sticky;
    top: 0;
    left: 0;
    z-index: 3;
  `;
  headerRow.appendChild(cornerTh);

  const colHeaders: HTMLTableCellElement[] = [];
  for (let c = 0; c < colCount; c++) {
    const th = document.createElement('th');
    const wPx = colWidthMap.get(c) ?? DEFAULT_COL_PX;
    th.style.cssText = `
      width: ${wPx}px;
      min-width: ${wPx}px;
      height: 24px;
      background: #f3f4f6;
      border: 1px solid #d1d5db;
      color: #4b5563;
      font-weight: 600;
      font-size: 11px;
      text-align: center;
      vertical-align: middle;
      position: sticky;
      top: 0;
      z-index: 2;
    `;
    th.textContent = indexToColAlpha(c);
    colHeaders.push(th);
    headerRow.appendChild(th);
  }
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  const cellEls = new Map<string, HTMLTableCellElement>();
  const rowHeaders: HTMLTableCellElement[] = [];

  for (let r = 0; r < rowCount; r++) {
    const tr = document.createElement('tr');
    const rowModel = sheet.rows.get(r);
    tr.style.height = `${rowModel?.heightPx ?? DEFAULT_ROW_PX}px`;
    if (rowModel?.hidden) tr.style.display = 'none';

    const rowTh = document.createElement('th');
    rowTh.style.cssText = `
      width: 46px;
      min-width: 46px;
      background: #f3f4f6;
      border: 1px solid #d1d5db;
      color: #4b5563;
      font-weight: 600;
      font-size: 11px;
      text-align: center;
      vertical-align: middle;
      position: sticky;
      left: 0;
      z-index: 1;
    `;
    rowTh.textContent = `${r + 1}`;
    rowHeaders.push(rowTh);
    tr.appendChild(rowTh);

    for (let c = 0; c < colCount; c++) {
      const mergeInfo = mergeLookup.get(`${r}:${c}`);
      if (mergeInfo && !mergeInfo.isOrigin) continue;

      const ref = `${indexToColAlpha(c)}${r + 1}`;
      const cell = rowModel?.cells.get(c);
      const td = document.createElement('td');
      td.setAttribute(XLSX_CELL_ATTR, ref);
      applyCellStyle(td, cell);

      if (mergeInfo?.isOrigin) {
        const rowSpan = mergeInfo.merge.endRow - mergeInfo.merge.startRow + 1;
        const colSpan = mergeInfo.merge.endCol - mergeInfo.merge.startCol + 1;
        if (rowSpan > 1) td.rowSpan = rowSpan;
        if (colSpan > 1) td.colSpan = colSpan;
      }

      td.textContent = cell?.formattedValue ?? '';
      cellEls.set(ref, td);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  gridContainer.appendChild(table);
  root.appendChild(gridContainer);

  // ── Selection ──────────────────────────────────────────────────────────────
  const cellAt = (ref: string): XlsxCell | undefined => {
    const { rowIndex, colIndex } = parseCellRef(ref);
    return sheet.rows.get(rowIndex)?.cells.get(colIndex);
  };

  let activeRef = options.activeCellRef ?? 'A1';
  if (!cellEls.has(activeRef)) activeRef = 'A1';
  let anchorRef = activeRef;
  let editingRef: string | null = null;
  /** Where the selection is heading once the edit in progress commits. */
  let pendingNext: string | null = null;

  const rangeRefs = (): string[] => {
    const a = parseCellRef(anchorRef);
    const b = parseCellRef(activeRef);
    const refs: string[] = [];
    for (let r = Math.min(a.rowIndex, b.rowIndex); r <= Math.max(a.rowIndex, b.rowIndex); r++) {
      for (let c = Math.min(a.colIndex, b.colIndex); c <= Math.max(a.colIndex, b.colIndex); c++) {
        const ref = `${indexToColAlpha(c)}${r + 1}`;
        if (cellEls.has(ref)) refs.push(ref);
      }
    }
    return refs.length ? refs : [activeRef];
  };

  const paintSelection = (): void => {
    for (const [ref, el] of cellEls) {
      el.style.outline = '';
      el.style.outlineOffset = '';
      el.style.boxShadow = '';
      if (ref === activeRef) {
        el.style.outline = `2px solid ${ACCENT}`;
        el.style.outlineOffset = '-2px';
      }
    }
    if (anchorRef !== activeRef) {
      // A tint that works over whatever fill the cell already has.
      for (const ref of rangeRefs()) {
        if (ref === activeRef) continue;
        const el = cellEls.get(ref);
        if (el) el.style.boxShadow = `inset 0 0 0 9999px rgba(16, 124, 65, 0.10)`;
      }
    }
    const { rowIndex, colIndex } = parseCellRef(activeRef);
    colHeaders.forEach((th, i) => {
      th.style.color = i === colIndex ? ACCENT : '#4b5563';
      th.style.background = i === colIndex ? '#e6f2ea' : '#f3f4f6';
    });
    rowHeaders.forEach((th, i) => {
      th.style.color = i === rowIndex ? ACCENT : '#4b5563';
      th.style.background = i === rowIndex ? '#e6f2ea' : '#f3f4f6';
    });
  };

  const showInFormulaBar = (): void => {
    const cell = cellAt(activeRef);
    nameBox.textContent = activeRef;
    fxInput.textContent = editableText(cell);
  };

  const select = (ref: string, extend = false): void => {
    if (!cellEls.has(ref)) return;
    if (editingRef && editingRef !== ref) {
      // Committing re-renders the sheet, so the new selection has to travel
      // with the commit rather than be applied to a view that is about to go.
      pendingNext = ref;
      commitEdit();
      return;
    }
    activeRef = ref;
    if (!extend) anchorRef = ref;
    paintSelection();
    showInFormulaBar();
    // Not implemented in jsdom, and only ever a convenience.
    cellEls.get(ref)?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
    options.onSelectCell?.(cellAt(ref), ref);
    options.onSelectionChange?.(rangeRefs(), ref);
  };

  /** The reference `dRow`/`dCol` away from the active cell, clamped to the grid. */
  const offsetRef = (dRow: number, dCol: number): string => {
    const { rowIndex, colIndex } = parseCellRef(activeRef);
    const r = Math.max(0, Math.min(rowCount - 1, rowIndex + dRow));
    const c = Math.max(0, Math.min(colCount - 1, colIndex + dCol));
    return `${indexToColAlpha(c)}${r + 1}`;
  };

  const move = (dRow: number, dCol: number, extend = false): void => {
    select(offsetRef(dRow, dCol), extend);
  };

  // ── Editing ────────────────────────────────────────────────────────────────
  const isEditable = (ref: string): boolean => !!editing && editing.editable(ref);

  const beginEdit = (seed?: string): void => {
    if (!editing || editingRef || !isEditable(activeRef)) return;
    const td = cellEls.get(activeRef);
    if (!td) return;
    editingRef = activeRef;
    // setAttribute rather than the property: it is what actually makes the cell
    // editable in a browser, and unlike the property it is observable in jsdom.
    td.setAttribute('contenteditable', 'true');
    td.style.whiteSpace = 'pre-wrap';
    td.style.overflow = 'visible';
    td.textContent = seed ?? editableText(cellAt(activeRef));
    td.focus();
    placeCaretAtEnd(td);
  };

  const restoreRendered = (ref: string, td: HTMLTableCellElement): void => {
    td.removeAttribute('contenteditable');
    td.textContent = cellAt(ref)?.formattedValue ?? '';
  };

  const commitEdit = (): boolean => {
    if (!editingRef) return false;
    const ref = editingRef;
    const td = cellEls.get(ref);
    editingRef = null;
    if (!td) return false;
    const text = (td.textContent ?? '').replace(/\r/g, '');
    const next = pendingNext ?? undefined;
    pendingNext = null;
    restoreRendered(ref, td);
    options.onCommit?.(ref, text, next);
    return true;
  };

  const cancelEdit = (): void => {
    if (!editingRef) return;
    const ref = editingRef;
    editingRef = null;
    const td = cellEls.get(ref);
    if (td) restoreRendered(ref, td);
  };

  // ── Events ─────────────────────────────────────────────────────────────────
  gridContainer.addEventListener('mousedown', (e) => {
    const td = cellElOf(e.target);
    if (!td) return;
    const ref = td.getAttribute(XLSX_CELL_ATTR);
    if (!ref || ref === editingRef) return;
    // Let the cell being edited keep the caret; otherwise the grid owns focus.
    if (!editing) {
      select(ref, e.shiftKey);
      return;
    }
    e.preventDefault();
    select(ref, e.shiftKey);
    gridContainer.focus();
  });

  gridContainer.addEventListener('dblclick', (e) => {
    if (!editing) return;
    const td = cellElOf(e.target);
    if (!td) return;
    beginEdit();
  });

  if (editing) {
    gridContainer.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        cancelEdit();
        gridContainer.focus();
        return;
      }
      if (e.key === 'Enter') {
        e.preventDefault();
        const next = offsetRef(e.shiftKey ? -1 : 1, 0);
        if (editingRef) {
          pendingNext = next;
          commitEdit();
          gridContainer.focus();
        } else {
          select(next);
        }
        return;
      }
      if (e.key === 'Tab') {
        e.preventDefault();
        const next = offsetRef(0, e.shiftKey ? -1 : 1);
        if (editingRef) {
          pendingNext = next;
          commitEdit();
          gridContainer.focus();
        } else {
          select(next);
        }
        return;
      }
      if (editingRef) return; // everything else belongs to the cell's editor

      if ((e.ctrlKey || e.metaKey) && (e.key === 'z' || e.key === 'Z')) {
        e.preventDefault();
        if (e.shiftKey) options.onRedo?.();
        else options.onUndo?.();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || e.key === 'Y')) {
        e.preventDefault();
        options.onRedo?.();
        return;
      }

      switch (e.key) {
        case 'ArrowUp':
          e.preventDefault();
          move(-1, 0, e.shiftKey);
          return;
        case 'ArrowDown':
          e.preventDefault();
          move(1, 0, e.shiftKey);
          return;
        case 'ArrowLeft':
          e.preventDefault();
          move(0, -1, e.shiftKey);
          return;
        case 'ArrowRight':
          e.preventDefault();
          move(0, 1, e.shiftKey);
          return;
        case 'Home':
          e.preventDefault();
          select(`A${parseCellRef(activeRef).rowIndex + 1}`, e.shiftKey);
          return;
        case 'F2':
          e.preventDefault();
          beginEdit();
          return;
        case 'Delete':
        case 'Backspace':
          e.preventDefault();
          options.onClearCells?.(rangeRefs());
          return;
      }

      // A printable character starts an edit that replaces the cell, as Excel does.
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        beginEdit(e.key);
      }
    });

    gridContainer.addEventListener(
      'focusout',
      (e) => {
        if (!editingRef) return;
        const to = e.relatedTarget;
        if (to instanceof Node && cellEls.get(editingRef)?.contains(to)) return;
        commitEdit();
      },
      true,
    );

    // The formula bar edits the active cell, like Excel's does.
    fxInput.contentEditable = 'true';
    fxInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const ref = activeRef;
        const text = fxInput.textContent ?? '';
        const next = offsetRef(1, 0);
        gridContainer.focus();
        if (isEditable(ref)) options.onCommit?.(ref, text, next);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        showInFormulaBar();
        gridContainer.focus();
      }
    });
  }

  paintSelection();
  showInFormulaBar();
  options.onSelectCell?.(cellAt(activeRef), activeRef);

  return {
    root,
    get activeRef() {
      return activeRef;
    },
    selectedRefs: rangeRefs,
    select,
    beginEdit,
    commitEdit,
    cancelEdit,
    focusGrid: () => gridContainer.focus(),
  };
}

/** The `<td>` an event happened inside, if any. */
function cellElOf(target: EventTarget | null): HTMLTableCellElement | null {
  let node: Node | null = target instanceof Node ? target : null;
  while (node) {
    if (node.nodeType === 1 && (node as Element).hasAttribute(XLSX_CELL_ATTR)) {
      return node as HTMLTableCellElement;
    }
    node = node.parentNode;
  }
  return null;
}

function placeCaretAtEnd(el: HTMLElement): void {
  const doc = el.ownerDocument;
  const sel = doc.getSelection?.();
  if (!sel || typeof doc.createRange !== 'function') return;
  const range = doc.createRange();
  range.selectNodeContents(el);
  range.collapse(false);
  sel.removeAllRanges();
  sel.addRange(range);
}

/** Paint one cell's resolved style onto its `<td>`. */
function applyCellStyle(td: HTMLTableCellElement, cell: XlsxCell | undefined): void {
  td.style.cssText = `
    border: 1px solid #e5e7eb;
    padding: 2px 6px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    vertical-align: bottom;
    background: #ffffff;
  `;
  // Excel right-aligns numbers and dates unless told otherwise.
  if (cell && (cell.type === 'n' || cell.type === 'b' || cell.type === 'e')) {
    td.style.textAlign = cell.type === 'n' ? 'right' : 'center';
  }
  const s = cell?.style;
  if (!s) return;

  if (s.font) {
    if (s.font.name) td.style.fontFamily = s.font.name;
    if (s.font.sizePt) td.style.fontSize = `${s.font.sizePt}pt`;
    if (s.font.bold) td.style.fontWeight = 'bold';
    if (s.font.italic) td.style.fontStyle = 'italic';
    if (s.font.colorHex) td.style.color = s.font.colorHex;
    const decorations: string[] = [];
    if (s.font.underline) decorations.push('underline');
    if (s.font.strike) decorations.push('line-through');
    if (decorations.length) td.style.textDecoration = decorations.join(' ');
  }

  if (s.fill?.patternType === 'solid' && s.fill.fgColorHex) {
    td.style.backgroundColor = s.fill.fgColorHex;
  }

  if (s.alignment) {
    if (s.alignment.horizontal) td.style.textAlign = s.alignment.horizontal;
    if (s.alignment.vertical) td.style.verticalAlign = s.alignment.vertical;
    if (s.alignment.wrapText) td.style.whiteSpace = 'pre-wrap';
  }

  if (s.border) {
    const side = (
      name: 'Top' | 'Bottom' | 'Left' | 'Right',
      def: { style?: string; colorHex?: string } | undefined,
    ) => {
      if (!def?.style || def.style === 'none') return;
      const width = def.style === 'thick' ? 2 : 1;
      const kind = def.style === 'dashed' || def.style === 'dotted' ? def.style : 'solid';
      td.style[`border${name}` as 'borderTop'] =
        `${width}px ${kind} ${def.colorHex ?? '#d1d5db'}`;
    };
    side('Top', s.border.top);
    side('Bottom', s.border.bottom);
    side('Left', s.border.left);
    side('Right', s.border.right);
  }
}
