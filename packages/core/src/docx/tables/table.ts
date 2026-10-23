/**
 * Table parsing: <w:tbl> -> DocxTable.
 *
 * Builds a rectangular grid from tblGrid column widths, resolving horizontal
 * spans (<w:gridSpan>) and vertical merges (<w:vMerge>) into colSpan/rowSpan on
 * origin cells with null placeholders for covered positions.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { twipToPx } from '../units.js';
import { logicalChildren, logicalChildrenNamed } from '../content.js';
import { parseParagraph, borderToStroke, type TableBase } from '../paragraphs/paragraph.js';
import type { DocxBlock, DocxRowHeight, DocxTable, DocxTableCell } from '../model.js';
import { applyContextualSpacing } from '../paragraphs/paragraph.js';
import {
  parseTableLook,
  tableCellFormat,
  type RawBorder,
  type RawBorders,
  type TableCond,
} from '../styles/styles.js';
import type { ParseContext } from '../document/context.js';

export function parseTable(tbl: XmlNode, ctx: ParseContext): DocxTable {
  const grid = children(child(tbl, 'tblGrid'), 'gridCol').map((c) => twipToPx(attrNum(c, 'w') ?? 0));

  const tblPr = child(tbl, 'tblPr');
  const style = ctx.styles.resolveTableStyle(attr(child(tblPr, 'tblStyle'), 'val'));
  const look = parseTableLook(child(tblPr, 'tblLook'));

  // Table-level borders: the style's, overlaid by any declared directly.
  const outerBorders: RawBorders = { ...style?.tblBorders };
  const directBorders = child(tblPr, 'tblBorders');
  if (directBorders) Object.assign(outerBorders, readSideBorders(directBorders));
  const insideH = readSide(directBorders, 'insideH') ?? style?.insideH;
  const insideV = readSide(directBorders, 'insideV') ?? style?.insideV;

  // Cell margins: the style's, side by side overlaid by the table's own, with
  // Word's built-in 0 / 0.08" as the floor.
  const styleCellMar = style?.cellMar
    ? {
        top: twipToPx(style.cellMar.top),
        right: twipToPx(style.cellMar.right),
        bottom: twipToPx(style.cellMar.bottom),
        left: twipToPx(style.cellMar.left),
      }
    : WORD_CELL_MARGINS;
  const defaultCellMar = { ...styleCellMar, ...readCellMargins(child(tblPr, 'tblCellMar')) };

  const trList = logicalChildrenNamed(tbl, 'tr');
  const rowCount = trList.length;
  const colCount = grid.length;

  const rows: (DocxTableCell | null)[][] = [];
  /** colIndex -> the origin cell currently being vertically merged. */
  const vmergeOrigin = new Map<number, DocxTableCell>();

  const rowHeights: (DocxRowHeight | undefined)[] = [];
  const cantSplit: boolean[] = [];
  let headerRows = 0;
  let headerRun = true;

  for (let r = 0; r < trList.length; r++) {
    const tr = trList[r]!;
    const row: (DocxTableCell | null)[] = [];
    const trPr = child(tr, 'trPr');
    const trHeight = child(trPr, 'trHeight');
    const hVal = attrNum(trHeight, 'val');
    if (hVal) rowHeights[r] = { px: twipToPx(hVal), rule: attr(trHeight, 'hRule') === 'exact' ? 'exact' : 'atLeast' };
    const rowCantSplit = child(trPr, 'cantSplit');
    if (rowCantSplit ? flagOn(rowCantSplit) : style?.cantSplit) cantSplit[r] = true;
    // Header rows are the unbroken run of <w:tblHeader> rows at the top.
    if (headerRun && flagOn(child(trPr, 'tblHeader'))) headerRows++;
    else headerRun = false;

    // <w:gridBefore>: grid columns skipped before the row's first cell.
    let col = 0;
    const gridBefore = attrNum(child(trPr, 'gridBefore'), 'val') ?? 0;
    for (let i = 0; i < gridBefore; i++) row.push(EMPTY_CELL());
    col += gridBefore;
    for (const tc of logicalChildrenNamed(tr, 'tc')) {
      const tcPr = child(tc, 'tcPr');
      const colSpan = attrNum(child(tcPr, 'gridSpan'), 'val') ?? 1;
      const vMerge = child(tcPr, 'vMerge');
      const vMergeVal = vMerge ? attr(vMerge, 'val') ?? 'continue' : undefined;

      if (vMergeVal === 'continue') {
        // Covered by a vertical merge from above: extend that origin's rowSpan.
        const origin = vmergeOrigin.get(col);
        if (origin) origin.rowSpan += 1;
        for (let i = 0; i < colSpan; i++) row.push(null);
        col += colSpan;
        continue;
      }

      const cond = style
        ? tableCellFormat(style, r, col, rowCount, colCount, look)
        : undefined;
      const edge = edgeBorders(r, col, colSpan, rowCount, colCount, outerBorders, insideH, insideV);
      const cell = buildCell(tc, tcPr, colSpan, defaultCellMar, ctx, cond, edge);
      row.push(cell);
      for (let i = 1; i < colSpan; i++) row.push(null);

      if (vMergeVal === 'restart') vmergeOrigin.set(col, cell);
      else vmergeOrigin.delete(col);

      col += colSpan;
    }
    // Pad a short row (<w:gridAfter>, or a malformed one) out to the grid.
    while (col < colCount) {
      row.push(EMPTY_CELL());
      col++;
    }
    rows.push(row);
  }

  const table: DocxTable = { kind: 'table', colWidths: grid, rows };
  const gridW = grid.reduce((a, b) => a + b, 0);
  if (gridW) table.widthPx = gridW;
  // A percentage width (fiftieths of a percent) is resolved against the text
  // column when the section is known — see fitTableWidth.
  const tblW = child(tblPr, 'tblW');
  if (attr(tblW, 'type') === 'pct') {
    const raw = attr(tblW, 'w') ?? '';
    const pct = raw.endsWith('%') ? parseFloat(raw) : (Number(raw) || 0) / 50;
    if (pct > 0) table.widthPct = pct;
  }

  const jc = attr(child(tblPr, 'jc'), 'val');
  if (jc === 'center') table.align = 'center';
  else if (jc === 'right' || jc === 'end') table.align = 'right';

  let indentPx = twipToPx(attrNum(child(tblPr, 'tblInd'), 'w') ?? 0);
  // Before Word 2013 the table's left *border* sat one cell margin left of the
  // indent, so the first cell's text — not its border — lined up with the margin.
  if (ctx.settings.compatMode < 15 && table.align === undefined) indentPx -= defaultCellMar.left;
  if (indentPx) table.indentPx = indentPx;

  if (rowHeights.length) table.rowHeights = rowHeights;
  if (cantSplit.length) table.cantSplit = cantSplit;
  if (headerRows > 0 && headerRows < rows.length) table.headerRows = headerRows;
  return table;
}

/** Word's built-in cell margins (the "Normal Table" style): 0 top/bottom, 0.08" sides. */
const WORD_CELL_MARGINS = { top: 0, right: twipToPx(108), bottom: 0, left: twipToPx(108) };

const EMPTY_CELL = (): DocxTableCell => ({ content: [], rowSpan: 1, colSpan: 1, empty: true });

function flagOn(el: XmlNode | undefined): boolean {
  if (!el) return false;
  const v = attr(el, 'val');
  return v === undefined || (v !== '0' && v !== 'false' && v !== 'off');
}

/**
 * Size a table against its section. A percentage width is resolved against
 * the text column. A table stated wider than the column is *not* shrunk —
 * Word lets it run into the right margin — only one that would run off the
 * page edge itself is scaled down to end there.
 */
export function fitTableWidth(table: DocxTable, contentW: number, pageRoomW: number = contentW): void {
  // A percentage-width table is sized to its share of the column first.
  if (table.widthPct !== undefined) {
    const target = (contentW * table.widthPct) / 100;
    const natural = table.colWidths.reduce((a, b) => a + b, 0);
    if (natural > 0 && target > 0) {
      table.colWidths = table.colWidths.map((w) => (w * target) / natural);
      table.widthPx = target;
    }
  }
  const indent = table.indentPx ?? 0;
  const avail = pageRoomW - indent; // negative indent widens the available space
  const natural = table.colWidths.reduce((a, b) => a + b, 0);
  if (natural <= 0 || avail <= 0 || natural <= avail) return;

  const scale = avail / natural;
  table.colWidths = table.colWidths.map((w) => w * scale);
  table.widthPx = avail;
}

/** The four border sides that apply to a cell from table-level borders. */
function edgeBorders(
  row: number,
  col: number,
  colSpan: number,
  rowCount: number,
  colCount: number,
  outer: RawBorders,
  insideH: RawBorder | undefined,
  insideV: RawBorder | undefined,
): RawBorders {
  const lastRow = row === rowCount - 1;
  const lastCol = col + colSpan - 1 === colCount - 1;
  const out: RawBorders = {};
  const top = row === 0 ? outer.top : insideH;
  const bottom = lastRow ? outer.bottom : insideH;
  const left = col === 0 ? outer.left : insideV;
  const right = lastCol ? outer.right : insideV;
  if (top) out.top = top;
  if (bottom) out.bottom = bottom;
  if (left) out.left = left;
  if (right) out.right = right;
  return out;
}

function buildCell(
  tc: XmlNode,
  tcPr: XmlNode | undefined,
  colSpan: number,
  defaultCellMar: NonNullable<DocxTableCell['cellPaddingPx']>,
  ctx: ParseContext,
  cond: TableCond | undefined,
  edge: RawBorders,
): DocxTableCell {
  const tableBase: TableBase | undefined = cond ? { pPr: cond.pPr, rPr: cond.rPr } : undefined;
  const content: DocxBlock[] = [];
  for (const node of logicalChildren(tc)) {
    const name = localName(node.name);
    if (name === 'p') content.push(...parseParagraph(node, ctx, tableBase));
    else if (name === 'tbl') content.push(parseTable(node, ctx));
  }
  applyContextualSpacing(content);

  const cell: DocxTableCell = { content, rowSpan: 1, colSpan };

  // Fill: direct cell shading, else the table style's conditional fill.
  const shdFill = attr(child(tcPr, 'shd'), 'fill');
  if (shdFill && shdFill !== 'auto') cell.fillHex = shdFill.replace(/^#/, '').toLowerCase();
  else if (cond?.tc.fillHex) cell.fillHex = cond.tc.fillHex;

  // Borders: table-level edges, overlaid by the style's conditional cell
  // borders, then any direct <w:tcBorders> — direct wins.
  const merged: RawBorders = { ...edge, ...cond?.tc.borders, ...readSideBordersRaw(child(tcPr, 'tcBorders')) };
  const borders = toStrokes(merged);
  if (borders) cell.borders = borders;

  const vAlign = attr(child(tcPr, 'vAlign'), 'val');
  if (vAlign === 'center' || vAlign === 'bottom' || vAlign === 'top') cell.vAlign = vAlign;
  else if (cond?.tc.vAlign) cell.vAlign = cond.tc.vAlign;

  const cellMar = { ...defaultCellMar, ...readCellMargins(child(tcPr, 'tcMar')) };
  cell.cellPaddingPx = cellMar;

  return cell;
}

/** Convert raw table/cell borders to render-ready per-side strokes. */
function toStrokes(b: RawBorders): DocxTableCell['borders'] {
  const out: NonNullable<DocxTableCell['borders']> = {};
  let any = false;
  const sides: [keyof NonNullable<DocxTableCell['borders']>, keyof RawBorders][] = [
    ['l', 'left'],
    ['t', 'top'],
    ['r', 'right'],
    ['b', 'bottom'],
  ];
  for (const [key, side] of sides) {
    const stroke = borderToStroke(b[side]);
    if (stroke) {
      out[key] = stroke;
      any = true;
    }
  }
  return any ? out : undefined;
}

/** The sides a margins element states (start/end are the bidi-neutral left/right). */
function readCellMargins(mar: XmlNode | undefined): Partial<NonNullable<DocxTableCell['cellPaddingPx']>> {
  const out: Partial<NonNullable<DocxTableCell['cellPaddingPx']>> = {};
  if (!mar) return out;
  const side = (...names: string[]) => {
    for (const name of names) {
      const el = child(mar, name);
      // Only dxa (twips) is meaningful here; a missing type means dxa.
      const w = attrNum(el, 'w');
      if (el && w !== undefined && (attr(el, 'type') ?? 'dxa') === 'dxa') return twipToPx(w);
    }
    return undefined;
  };
  const top = side('top');
  const right = side('right', 'end');
  const bottom = side('bottom');
  const left = side('left', 'start');
  if (top !== undefined) out.top = top;
  if (right !== undefined) out.right = right;
  if (bottom !== undefined) out.bottom = bottom;
  if (left !== undefined) out.left = left;
  return out;
}

/** Read one named border side (top/left/insideH/…) into a RawBorder. */
function readSide(parent: XmlNode | undefined, name: string): RawBorder | undefined {
  const e = child(parent, name);
  if (!e) return undefined;
  return { sz: attrNum(e, 'sz'), colorHex: hexOrUndef(attr(e, 'color')), val: attr(e, 'val') };
}

/** Read the four outer sides (top/bottom/left/right) of a borders element. */
function readSideBorders(parent: XmlNode | undefined): RawBorders {
  const out: RawBorders = {};
  for (const side of ['top', 'bottom', 'left', 'right'] as const) {
    const b = readSide(parent, side);
    if (b) out[side] = b;
  }
  return out;
}

/** Like {@link readSideBorders} but returns undefined-free only for present sides. */
function readSideBordersRaw(parent: XmlNode | undefined): RawBorders {
  return parent ? readSideBorders(parent) : {};
}

function hexOrUndef(v: string | undefined): string | undefined {
  return v && v !== 'auto' ? v.replace(/^#/, '').toLowerCase() : undefined;
}
