import { type XmlNode, children, child, attr, attrNum } from '../../oxml/xml.js';
import type {
  XlsxSheet,
  XlsxRow,
  XlsxColumn,
  XlsxCell,
  XlsxMergeCell,
  CellType,
} from '../model.js';
import type { XlsxStyles } from '../styles/styles.js';

export function colAlphaToIndex(colStr: string): number {
  let index = 0;
  const upper = colStr.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

export function indexToColAlpha(index: number): string {
  let colStr = '';
  let num = index + 1;
  while (num > 0) {
    const rem = (num - 1) % 26;
    colStr = String.fromCharCode(65 + rem) + colStr;
    num = Math.floor((num - 1) / 26);
  }
  return colStr;
}

export function parseCellRef(ref: string): { rowIndex: number; colIndex: number } {
  const match = ref.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) return { rowIndex: 0, colIndex: 0 };
  const [, colStr, rowStr] = match;
  return {
    colIndex: colAlphaToIndex(colStr!),
    rowIndex: parseInt(rowStr!, 10) - 1,
  };
}

export function parseMergeRef(ref: string): XlsxMergeCell {
  const [start, end] = ref.split(':');
  const p1 = parseCellRef(start ?? 'A1');
  const p2 = parseCellRef(end ?? start ?? 'A1');
  return {
    ref,
    startRow: Math.min(p1.rowIndex, p2.rowIndex),
    startCol: Math.min(p1.colIndex, p2.colIndex),
    endRow: Math.max(p1.rowIndex, p2.rowIndex),
    endCol: Math.max(p1.colIndex, p2.colIndex),
  };
}

export function parseSharedStrings(xmlNode?: XmlNode): string[] {
  if (!xmlNode) return [];
  const strings: string[] = [];
  for (const si of children(xmlNode, 'si')) {
    const tNode = child(si, 't');
    if (tNode) {
      strings.push(tNode.text ?? '');
    } else {
      // Rich text runs <r><t>...</t></r>
      let combined = '';
      for (const r of children(si, 'r')) {
        const rt = child(r, 't');
        if (rt?.text) combined += rt.text;
      }
      strings.push(combined);
    }
  }
  return strings;
}

function formatCellText(rawValue: string, type: CellType, sharedStrings: string[], numFmtCode?: string): string {
  if (type === 's') {
    const idx = parseInt(rawValue, 10);
    return sharedStrings[idx] ?? rawValue;
  }
  if (type === 'b') {
    return rawValue === '1' || rawValue.toLowerCase() === 'true' ? 'TRUE' : 'FALSE';
  }
  if (type === 'e') {
    return rawValue; // Error e.g. #N/A, #VALUE!
  }
  if (numFmtCode && !isNaN(Number(rawValue))) {
    const num = Number(rawValue);
    if (numFmtCode.includes('%')) {
      return `${(num * 100).toFixed(numFmtCode.includes('.00') ? 2 : 0)}%`;
    }
    if (numFmtCode.includes('.00')) {
      return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
  }
  return rawValue;
}

export function parseSheetXml(
  id: string,
  name: string,
  sheetXml: XmlNode,
  sharedStrings: string[],
  styles: XlsxStyles,
  /**
   * Called with every parsed cell and the `<c>` element it came from. The edit
   * layer uses this to address the XML a cell must be written back into; a
   * read-only parse simply omits it.
   */
  recordSource?: (cell: XlsxCell, node: XmlNode) => void,
): XlsxSheet {
  const rows = new Map<number, XlsxRow>();
  const cols: XlsxColumn[] = [];
  const mergeCells: XlsxMergeCell[] = [];

  let maxRow = 0;
  let maxCol = 0;

  // Parse columns <cols><col min="1" max="1" width="15" customWidth="1"/></cols>
  const colsNode = child(sheetXml, 'cols');
  if (colsNode) {
    for (const cNode of children(colsNode, 'col')) {
      const min = attrNum(cNode, 'min') ?? 1;
      const max = attrNum(cNode, 'max') ?? min;
      const width = attrNum(cNode, 'width');
      const hidden = attr(cNode, 'hidden') === '1' || attr(cNode, 'hidden') === 'true';
      // Excel 1 width unit is approximately 7-8px (char width + padding)
      const widthPx = width !== undefined ? Math.round(width * 8) : undefined;
      cols.push({ min, max, widthPx, hidden });
    }
  }

  // Parse sheet data <sheetData><row r="1">...
  const sheetData = child(sheetXml, 'sheetData');
  if (sheetData) {
    for (const rowNode of children(sheetData, 'row')) {
      const rNum = attrNum(rowNode, 'r');
      const rowIndex = rNum !== undefined ? rNum - 1 : rows.size;
      const ht = attrNum(rowNode, 'ht');
      const rowHidden = attr(rowNode, 'hidden') === '1' || attr(rowNode, 'hidden') === 'true';

      const row: XlsxRow = {
        rowIndex,
        heightPx: ht !== undefined ? Math.round((ht * 96) / 72) : undefined, // pt -> px
        hidden: rowHidden,
        cells: new Map<number, XlsxCell>(),
      };

      for (const cNode of children(rowNode, 'c')) {
        const ref = attr(cNode, 'r') ?? `${indexToColAlpha(row.cells.size)}${rowIndex + 1}`;
        const { colIndex } = parseCellRef(ref);
        const type = (attr(cNode, 't') as CellType) ?? 'n';
        const styleId = attrNum(cNode, 's');
        const style = styles.getStyle(styleId);

        let rawValue = '';
        const vNode = child(cNode, 'v');
        if (vNode?.text) {
          rawValue = vNode.text;
        } else {
          const isNode = child(cNode, 'is');
          if (isNode) {
            const tNode = child(isNode, 't');
            if (tNode?.text) rawValue = tNode.text;
          }
        }

        let formula: string | undefined;
        const fNode = child(cNode, 'f');
        if (fNode?.text) {
          formula = fNode.text;
        }

        const formattedValue = formatCellText(rawValue, type, sharedStrings, style?.numFmtCode);

        const cell: XlsxCell = {
          ref,
          colIndex,
          rowIndex,
          type,
          rawValue,
          formattedValue,
          formula,
          styleId,
          style,
        };

        recordSource?.(cell, cNode);
        row.cells.set(colIndex, cell);
        maxCol = Math.max(maxCol, colIndex + 1);
      }

      rows.set(rowIndex, row);
      maxRow = Math.max(maxRow, rowIndex + 1);
    }
  }

  // Parse merged cells <mergeCells><mergeCell ref="A1:C3"/></mergeCells>
  const mergeNodes = child(sheetXml, 'mergeCells');
  if (mergeNodes) {
    for (const mNode of children(mergeNodes, 'mergeCell')) {
      const ref = attr(mNode, 'ref');
      if (ref) {
        mergeCells.push(parseMergeRef(ref));
      }
    }
  }

  return {
    id,
    name,
    maxRow: Math.max(maxRow, 15), // At least 15 rows for display grid
    maxCol: Math.max(maxCol, 10), // At least 10 columns for display grid
    rows,
    cols,
    mergeCells,
  };
}
