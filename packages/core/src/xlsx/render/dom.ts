import type { XlsxSheet, XlsxCell, XlsxMergeCell, XlsxColumn } from '../model.js';
import { indexToColAlpha } from '../sheets/sheet.js';

export interface RenderXlsxOptions {
  activeCellRef?: string;
  onSelectCell?: (cell: XlsxCell | undefined, ref: string) => void;
}

export function renderXlsxSheet(sheet: XlsxSheet, options: RenderXlsxOptions = {}): HTMLElement {
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

  // Formula Bar
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
    color: #107c41;
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
  `;

  formulaBar.appendChild(nameBox);
  formulaBar.appendChild(fxLabel);
  formulaBar.appendChild(fxInput);
  root.appendChild(formulaBar);

  // Scrollable Table Grid Container
  const gridContainer = document.createElement('div');
  gridContainer.className = 'cbx-xlsx-grid-container';
  gridContainer.style.cssText = `
    flex: 1;
    min-height: 0;
    overflow: auto;
    position: relative;
  `;

  const table = document.createElement('table');
  table.className = 'cbx-xlsx-table';
  table.style.cssText = `
    border-collapse: collapse;
    table-layout: fixed;
    background: #ffffff;
  `;

  // Process Merged Cells Map
  const mergeLookup = new Map<string, { merge: XlsxMergeCell; isOrigin: boolean }>();
  for (const m of sheet.mergeCells) {
    for (let r = m.startRow; r <= m.endRow; r++) {
      for (let c = m.startCol; c <= m.endCol; c++) {
        const key = `${r}:${c}`;
        const isOrigin = r === m.startRow && c === m.startCol;
        mergeLookup.set(key, { merge: m, isOrigin });
      }
    }
  }

  // Build Table Header (Column headers A, B, C...)
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');

  // Corner header cell
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

  // Column width lookup
  const colWidthMap = new Map<number, number>();
  for (const colDef of sheet.cols) {
    for (let c = colDef.min - 1; c <= colDef.max - 1; c++) {
      if (colDef.widthPx) colWidthMap.set(c, colDef.widthPx);
    }
  }

  for (let c = 0; c < sheet.maxCol; c++) {
    const th = document.createElement('th');
    const wPx = colWidthMap.get(c) ?? 88;
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
    headerRow.appendChild(th);
  }
  thead.appendChild(headerRow);
  table.appendChild(thead);

  // Table Body
  const tbody = document.createElement('tbody');
  let selectedCellTd: HTMLTableCellElement | null = null;

  const selectCell = (cell: XlsxCell | undefined, ref: string, td: HTMLTableCellElement) => {
    if (selectedCellTd) {
      selectedCellTd.style.outline = '';
      selectedCellTd.style.outlineOffset = '';
    }
    selectedCellTd = td;
    td.style.outline = '2px solid #107c41';
    td.style.outlineOffset = '-2px';

    nameBox.textContent = ref;
    if (cell?.formula) {
      fxInput.textContent = `=${cell.formula}`;
    } else {
      fxInput.textContent = cell?.formattedValue ?? cell?.rawValue ?? '';
    }

    if (options.onSelectCell) {
      options.onSelectCell(cell, ref);
    }
  };

  for (let r = 0; r < sheet.maxRow; r++) {
    const tr = document.createElement('tr');
    const rowModel = sheet.rows.get(r);
    if (rowModel?.heightPx) {
      tr.style.height = `${rowModel.heightPx}px`;
    } else {
      tr.style.height = '22px';
    }

    // Row Index Header cell (1, 2, 3...)
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
    tr.appendChild(rowTh);

    for (let c = 0; c < sheet.maxCol; c++) {
      const mergeInfo = mergeLookup.get(`${r}:${c}`);
      if (mergeInfo && !mergeInfo.isOrigin) {
        // Covered by merge span, skip rendering cell
        continue;
      }

      const td = document.createElement('td');
      const cellRef = `${indexToColAlpha(c)}${r + 1}`;
      const cell = rowModel?.cells.get(c);

      td.style.cssText = `
        border: 1px solid #e5e7eb;
        padding: 2px 6px;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        vertical-align: bottom;
        background: #ffffff;
      `;

      if (mergeInfo?.isOrigin) {
        const rowSpan = mergeInfo.merge.endRow - mergeInfo.merge.startRow + 1;
        const colSpan = mergeInfo.merge.endCol - mergeInfo.merge.startCol + 1;
        if (rowSpan > 1) td.rowSpan = rowSpan;
        if (colSpan > 1) td.colSpan = colSpan;
      }

      // Apply cell styling if present
      if (cell?.style) {
        const s = cell.style;
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

        if (s.fill?.fgColorHex) {
          td.style.backgroundColor = s.fill.fgColorHex;
        }

        if (s.alignment) {
          if (s.alignment.horizontal) td.style.textAlign = s.alignment.horizontal;
          if (s.alignment.vertical) td.style.verticalAlign = s.alignment.vertical;
          if (s.alignment.wrapText) td.style.whiteSpace = 'pre-wrap';
        }

        if (s.border) {
          if (s.border.top?.style && s.border.top.style !== 'none') {
            td.style.borderTop = `1px solid ${s.border.top.colorHex ?? '#d1d5db'}`;
          }
          if (s.border.bottom?.style && s.border.bottom.style !== 'none') {
            td.style.borderBottom = `1px solid ${s.border.bottom.colorHex ?? '#d1d5db'}`;
          }
          if (s.border.left?.style && s.border.left.style !== 'none') {
            td.style.borderLeft = `1px solid ${s.border.left.colorHex ?? '#d1d5db'}`;
          }
          if (s.border.right?.style && s.border.right.style !== 'none') {
            td.style.borderRight = `1px solid ${s.border.right.colorHex ?? '#d1d5db'}`;
          }
        }
      }

      td.textContent = cell?.formattedValue ?? '';

      // Default selection on A1 or active cell
      if ((options.activeCellRef && cellRef === options.activeCellRef) || (!options.activeCellRef && r === 0 && c === 0)) {
        selectCell(cell, cellRef, td);
      }

      td.addEventListener('click', () => {
        selectCell(cell, cellRef, td);
      });

      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  gridContainer.appendChild(table);
  root.appendChild(gridContainer);

  return root;
}
