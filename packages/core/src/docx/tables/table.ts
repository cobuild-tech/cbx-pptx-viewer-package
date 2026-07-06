/**
 * Table parser for WordprocessingML.
 *
 * Converts a <w:tbl> element into a DocxTable IR node, handling column widths,
 * cell spans, borders, fills, and vertical merge groups.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { twipsToPx, borderSzToPx } from '../units.js';
import type { DocxTable, DocxTableCell, DocxParagraph, Fill, Stroke } from '../model.js';
import { parseParagraph, type ParagraphParseCtx } from '../paragraphs/paragraph.js';
import { encodeNodeId } from '../edit/nodeId.js';

export function parseTable(tblEl: XmlNode, ctx: ParagraphParseCtx, path?: number[]): DocxTable {
  // Column widths from <w:tblGrid>.
  const tblGrid = child(tblEl, 'tblGrid');
  const colWidths: number[] = [];
  for (const gridCol of children(tblGrid, 'gridCol')) {
    const w = attrNum(gridCol, 'w:w') ?? attrNum(gridCol, 'w') ?? 0;
    colWidths.push(twipsToPx(w));
  }

  // Table width (optional).
  const tblPr = child(tblEl, 'tblPr');
  const tblW = child(tblPr, 'tblW');
  let widthPx: number | undefined;
  if (tblW) {
    const tblWType = attr(tblW, 'w:type') ?? attr(tblW, 'type');
    const tblWVal = attrNum(tblW, 'w:w') ?? attrNum(tblW, 'w');
    if (tblWType === 'dxa' && tblWVal !== undefined) widthPx = twipsToPx(tblWVal);
  }

  // Table-level borders (outer + insideH/V) and default cell margins.
  const tblBorders = parseTableBorders(child(tblPr, 'tblBorders'));
  const defaultCellPadding = parseMarginSet(child(tblPr, 'tblCellMar'));

  const numCols = colWidths.length;
  const rawRows: RawCell[][] = [];

  // Indexed iteration (not children()) so cell paths use real XML child indices.
  for (let trIdx = 0; trIdx < tblEl.children.length; trIdx++) {
    const trEl = tblEl.children[trIdx]!;
    if (localName(trEl.name) !== 'tr') continue;
    const rawRow: RawCell[] = [];
    for (let tcIdx = 0; tcIdx < trEl.children.length; tcIdx++) {
      const tcEl = trEl.children[tcIdx]!;
      if (localName(tcEl.name) !== 'tc') continue;
      const cellPath = path ? [...path, trIdx, tcIdx] : undefined;
      rawRow.push(parseRawCell(tcEl, ctx, defaultCellPadding, cellPath));
    }
    rawRows.push(rawRow);
  }

  // Expand spans into a full grid (rows × cols), nulling covered cells.
  const rows: (DocxTableCell | null)[][] = buildGrid(rawRows, numCols);

  // Apply table-level border inheritance per cell position.
  applyTableBorders(rows, tblBorders);

  return { kind: 'table', widthPx, colWidths, rows };
}

// ─── Internal types ───────────────────────────────────────────────────────────

interface RawCell {
  cell: DocxTableCell;
  /** True if this cell is a vertical merge continuation (covered). */
  vMerge: boolean;
  /** True if this cell starts a vertical merge group. */
  vMergeRestart: boolean;
}

interface TableBorderDef {
  outer: Partial<Record<'l' | 't' | 'r' | 'b', Stroke>>;
  insideH?: Stroke;
  insideV?: Stroke;
}

// ─── Grid builder ─────────────────────────────────────────────────────────────

function buildGrid(rawRows: RawCell[][], numCols: number): (DocxTableCell | null)[][] {
  const numRows = rawRows.length;
  const grid: (DocxTableCell | null)[][] = Array.from({ length: numRows }, () =>
    new Array<DocxTableCell | null>(Math.max(numCols, 1)).fill(null),
  );

  // Track vertical merge groups: colIndex -> anchor cell
  const vMergeAnchor = new Map<number, { r: number; c: number }>();

  let col = 0;
  for (let r = 0; r < numRows; r++) {
    col = 0;
    for (const raw of rawRows[r]!) {
      // Skip columns already filled by a previous row's span.
      while (col < grid[r]!.length && grid[r]![col] !== null) col++;
      if (col >= (numCols || 1)) break;

      if (raw.vMerge && !raw.vMergeRestart) {
        // Continuation of a vertical merge: increment the anchor cell's rowSpan.
        const anchor = vMergeAnchor.get(col);
        if (anchor) {
          const anchorCell = grid[anchor.r]![anchor.c];
          if (anchorCell) anchorCell.rowSpan++;
          // Mark continuation cells as null.
          for (let span = 0; span < raw.cell.colSpan; span++) {
            if (col + span < grid[r]!.length) grid[r]![col + span] = null;
          }
        }
      } else {
        grid[r]![col] = raw.cell;
        if (raw.vMergeRestart) {
          vMergeAnchor.set(col, { r, c: col });
        } else {
          vMergeAnchor.delete(col);
        }
        // Fill spanned columns with null.
        for (let span = 1; span < raw.cell.colSpan; span++) {
          if (col + span < grid[r]!.length) grid[r]![col + span] = null;
        }
      }
      col += raw.cell.colSpan;
    }
  }
  return grid;
}

// ─── Border inheritance ────────────────────────────────────────────────────────

/**
 * Apply table-level border defaults to each cell based on its position in the
 * grid. Cell-level borders (already set on the cell) take priority. Outer table
 * borders apply to the outermost edges; insideH/V fill interior boundaries.
 */
function applyTableBorders(
  grid: (DocxTableCell | null)[][],
  tblBorders: TableBorderDef,
): void {
  const numRows = grid.length;
  if (numRows === 0) return;
  const numCols = grid[0]!.length;

  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const cell = grid[r]?.[c];
      if (!cell) continue;

      const existing = cell.borders ?? {};
      const resolved: Partial<Record<'l' | 't' | 'r' | 'b', Stroke>> = {};

      const lastRow = r + cell.rowSpan - 1;
      const lastCol = c + cell.colSpan - 1;

      if (existing.t !== undefined) resolved.t = existing.t;
      else if (r === 0 && tblBorders.outer.t) resolved.t = tblBorders.outer.t;
      else if (r > 0 && tblBorders.insideH) resolved.t = tblBorders.insideH;

      if (existing.b !== undefined) resolved.b = existing.b;
      else if (lastRow === numRows - 1 && tblBorders.outer.b) resolved.b = tblBorders.outer.b;
      else if (lastRow < numRows - 1 && tblBorders.insideH) resolved.b = tblBorders.insideH;

      if (existing.l !== undefined) resolved.l = existing.l;
      else if (c === 0 && tblBorders.outer.l) resolved.l = tblBorders.outer.l;
      else if (c > 0 && tblBorders.insideV) resolved.l = tblBorders.insideV;

      if (existing.r !== undefined) resolved.r = existing.r;
      else if (lastCol === numCols - 1 && tblBorders.outer.r) resolved.r = tblBorders.outer.r;
      else if (lastCol < numCols - 1 && tblBorders.insideV) resolved.r = tblBorders.insideV;

      cell.borders = resolved;
    }
  }
}

// ─── Cell parsing ─────────────────────────────────────────────────────────────

function parseRawCell(
  tcEl: XmlNode,
  ctx: ParagraphParseCtx,
  defaultCellPadding: { top: number; right: number; bottom: number; left: number } | undefined,
  path?: number[],
): RawCell {
  const tcPr = child(tcEl, 'tcPr');

  const colSpan = attrNum(child(tcPr, 'gridSpan'), 'w:val') ?? attrNum(child(tcPr, 'gridSpan'), 'val') ?? 1;

  const vMergeEl = child(tcPr, 'vMerge');
  const vMerge = vMergeEl !== undefined;
  const vMergeRestartVal = attr(vMergeEl, 'w:val') ?? attr(vMergeEl, 'val');
  const vMergeRestart = vMerge && vMergeRestartVal === 'restart';

  const fill = parseCellFill(tcPr);
  // Cell-level borders only; table-level defaults are applied after the grid is built.
  const borders = parseCellBorders(child(tcPr, 'tcBorders'));
  const cellPaddingPx = parseMarginSet(child(tcPr, 'tcMar')) ?? defaultCellPadding;

  const vAlignEl = child(tcPr, 'vAlign');
  const vAlignVal = attr(vAlignEl, 'w:val') ?? attr(vAlignEl, 'val');
  const vAlign = vAlignVal === 'center' ? 'center' : vAlignVal === 'bottom' ? 'bottom' : 'top';

  // Parse cell content (paragraphs only for now).
  const content: DocxParagraph[] = [];
  for (let i = 0; i < tcEl.children.length; i++) {
    const node = tcEl.children[i]!;
    if (localName(node.name) === 'p') {
      const pPath = path ? [...path, i] : undefined;
      const { paragraphs } = parseParagraph(node, ctx, pPath);
      content.push(...paragraphs);
    }
  }

  return {
    cell: {
      content,
      nodeId: path && ctx.partPath ? encodeNodeId(ctx.partPath, path) : undefined,
      fill,
      rowSpan: 1,
      colSpan,
      borders,
      vAlign: vAlign as 'top' | 'center' | 'bottom',
      cellPaddingPx,
    },
    vMerge,
    vMergeRestart,
  };
}

function parseCellFill(tcPr: XmlNode | undefined): Fill {
  if (!tcPr) return { type: 'none' };
  const shd = child(tcPr, 'shd');
  if (!shd) return { type: 'none' };
  const fill = attr(shd, 'w:fill') ?? attr(shd, 'fill');
  if (!fill || fill === 'auto' || fill === 'nil') return { type: 'none' };
  return { type: 'solid', color: { hex: fill.toUpperCase() } };
}

// ─── Border helpers ────────────────────────────────────────────────────────────

function parseBorderStroke(sideEl: XmlNode | undefined): Stroke | undefined {
  if (!sideEl) return undefined;
  const val = attr(sideEl, 'w:val') ?? attr(sideEl, 'val');
  if (!val || val === 'none' || val === 'nil') return undefined;
  const sz = attrNum(sideEl, 'w:sz') ?? attrNum(sideEl, 'sz') ?? 4;
  const colorHex = attr(sideEl, 'w:color') ?? attr(sideEl, 'color');
  const hex = colorHex && colorHex !== 'auto' ? colorHex.toUpperCase() : '000000';
  return { color: { hex }, width: borderSzToPx(sz) };
}

function parseCellBorders(
  bordersEl: XmlNode | undefined,
): Partial<Record<'l' | 't' | 'r' | 'b', Stroke>> {
  if (!bordersEl) return {};
  const out: Partial<Record<'l' | 't' | 'r' | 'b', Stroke>> = {};
  for (const [tag, key] of [['left', 'l'], ['top', 't'], ['right', 'r'], ['bottom', 'b']] as const) {
    const stroke = parseBorderStroke(child(bordersEl, tag));
    if (stroke) out[key] = stroke;
  }
  return out;
}

function parseTableBorders(bordersEl: XmlNode | undefined): TableBorderDef {
  if (!bordersEl) return { outer: {} };
  return {
    outer: parseCellBorders(bordersEl),
    insideH: parseBorderStroke(child(bordersEl, 'insideH')),
    insideV: parseBorderStroke(child(bordersEl, 'insideV')),
  };
}

// ─── Margin helpers ────────────────────────────────────────────────────────────

function parseMarginSet(
  marginEl: XmlNode | undefined,
): { top: number; right: number; bottom: number; left: number } | undefined {
  if (!marginEl) return undefined;
  let hasAny = false;

  function get(tag: string): number {
    const el = child(marginEl!, tag);
    if (!el) return 0;
    const type = attr(el, 'w:type') ?? attr(el, 'type');
    if (type === 'nil') { hasAny = true; return 0; }
    const w = attrNum(el, 'w:w') ?? attrNum(el, 'w') ?? 0;
    hasAny = true;
    return type === 'dxa' ? twipsToPx(w) : w;
  }

  const top = get('top');
  const right = get('right');
  const bottom = get('bottom');
  const left = get('left');
  return hasAny ? { top, right, bottom, left } : undefined;
}
