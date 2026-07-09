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
import type { DocxBlock, DocxTable, DocxTableCell } from '../model.js';
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

  const directCellMar = readCellMargins(child(tblPr, 'tblCellMar'));
  const styleCellMar = style?.cellMar
    ? {
        top: twipToPx(style.cellMar.top),
        right: twipToPx(style.cellMar.right),
        bottom: twipToPx(style.cellMar.bottom),
        left: twipToPx(style.cellMar.left),
      }
    : undefined;
  const defaultCellMar = directCellMar ?? styleCellMar;

  const trList = logicalChildrenNamed(tbl, 'tr');
  const rowCount = trList.length;
  const colCount = grid.length;

  const rows: (DocxTableCell | null)[][] = [];
  /** colIndex -> the origin cell currently being vertically merged. */
  const vmergeOrigin = new Map<number, DocxTableCell>();

  for (let r = 0; r < trList.length; r++) {
    const tr = trList[r]!;
    const row: (DocxTableCell | null)[] = [];
    let col = 0;
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
    rows.push(row);
  }

  const widthPx = grid.reduce((a, b) => a + b, 0) || undefined;
  const indentTwip = attrNum(child(tblPr, 'tblInd'), 'w');
  const indentPx = indentTwip !== undefined ? twipToPx(indentTwip) : undefined;
  return {
    kind: 'table',
    colWidths: grid,
    rows,
    ...(widthPx ? { widthPx } : {}),
    ...(indentPx ? { indentPx } : {}),
  };
}

/**
 * Scale a table (and its column widths) down so it fits the available width,
 * reproducing Word's autofit. A negative tblInd lets the table extend left into
 * the margin, adding to the width it may occupy. Narrower tables are untouched.
 */
export function fitTableWidth(table: DocxTable, contentW: number): void {
  const indent = table.indentPx ?? 0;
  const avail = contentW - indent; // negative indent widens the available space
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
  defaultCellMar: DocxTableCell['cellPaddingPx'],
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

  const cellMar = readCellMargins(child(tcPr, 'tcMar')) ?? defaultCellMar;
  if (cellMar) cell.cellPaddingPx = cellMar;

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

function readCellMargins(mar: XmlNode | undefined): DocxTableCell['cellPaddingPx'] {
  if (!mar) return undefined;
  const side = (name: string) => {
    const w = attrNum(child(mar, name), 'w');
    return w !== undefined ? twipToPx(w) : 0;
  };
  return { top: side('top'), right: side('right'), bottom: side('bottom'), left: side('left') };
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
