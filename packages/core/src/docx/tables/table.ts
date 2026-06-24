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

export function parseTable(tblEl: XmlNode, ctx: ParagraphParseCtx): DocxTable {
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

  // Table-level default borders (used when cells don't specify their own).
  const tblBorders = child(tblPr, 'tblBorders');
  const defaultBorders = parseBorderSet(tblBorders);

  const numCols = colWidths.length;
  const rawRows: RawCell[][] = [];

  for (const trEl of children(tblEl, 'tr')) {
    const rawRow: RawCell[] = [];
    for (const tcEl of children(trEl, 'tc')) {
      rawRow.push(parseRawCell(tcEl, ctx, defaultBorders));
    }
    rawRows.push(rawRow);
  }

  // Expand spans into a full grid (rows × cols), nulling covered cells.
  const rows: (DocxTableCell | null)[][] = buildGrid(rawRows, numCols);

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

// ─── Cell parsing ─────────────────────────────────────────────────────────────

function parseRawCell(
  tcEl: XmlNode,
  ctx: ParagraphParseCtx,
  defaultBorders: Partial<Record<'l' | 't' | 'r' | 'b', Stroke>>,
): RawCell {
  const tcPr = child(tcEl, 'tcPr');

  const colSpan = attrNum(child(tcPr, 'gridSpan'), 'w:val') ?? attrNum(child(tcPr, 'gridSpan'), 'val') ?? 1;

  const vMergeEl = child(tcPr, 'vMerge');
  const vMerge = vMergeEl !== undefined;
  const vMergeRestartVal = attr(vMergeEl, 'w:val') ?? attr(vMergeEl, 'val');
  const vMergeRestart = vMerge && vMergeRestartVal === 'restart';

  const fill = parseCellFill(tcPr);
  const borders = parseBorderSet(child(tcPr, 'tcBorders'));
  const mergedBorders = { ...defaultBorders, ...borders };

  const vAlignEl = child(tcPr, 'vAlign');
  const vAlignVal = attr(vAlignEl, 'w:val') ?? attr(vAlignEl, 'val');
  const vAlign = vAlignVal === 'center' ? 'center' : vAlignVal === 'bottom' ? 'bottom' : 'top';

  // Parse cell content (paragraphs only for now).
  const content: DocxParagraph[] = [];
  for (const node of tcEl.children) {
    const name = localName(node.name);
    if (name === 'p') {
      const { paragraphs } = parseParagraph(node, ctx);
      content.push(...paragraphs);
    }
  }

  return {
    cell: {
      content,
      fill,
      rowSpan: 1,
      colSpan,
      borders: mergedBorders as DocxTableCell['borders'],
      vAlign: vAlign as 'top' | 'center' | 'bottom',
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

function parseBorderSet(
  bordersEl: XmlNode | undefined,
): Partial<Record<'l' | 't' | 'r' | 'b', Stroke>> {
  if (!bordersEl) return {};
  const out: Partial<Record<'l' | 't' | 'r' | 'b', Stroke>> = {};
  const sides: Array<[string, 'l' | 't' | 'r' | 'b']> = [
    ['left', 'l'],
    ['top', 't'],
    ['right', 'r'],
    ['bottom', 'b'],
  ];
  for (const [tag, key] of sides) {
    const sideEl = child(bordersEl, tag);
    if (!sideEl) continue;
    const val = attr(sideEl, 'w:val') ?? attr(sideEl, 'val');
    if (!val || val === 'none' || val === 'nil') continue;
    const sz = attrNum(sideEl, 'w:sz') ?? attrNum(sideEl, 'sz') ?? 4;
    const colorHex = attr(sideEl, 'w:color') ?? attr(sideEl, 'color');
    const hex = colorHex && colorHex !== 'auto' ? colorHex.toUpperCase() : '000000';
    out[key] = { color: { hex }, width: borderSzToPx(sz) };
  }
  return out;
}
