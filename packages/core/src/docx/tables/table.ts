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
import { parseParagraph, borderToStroke } from '../paragraphs/paragraph.js';
import type { DocxBlock, DocxTable, DocxTableCell, Stroke } from '../model.js';
import type { RawBorder } from '../styles/styles.js';
import type { ParseContext } from '../document/context.js';

export function parseTable(tbl: XmlNode, ctx: ParseContext): DocxTable {
  const grid = children(child(tbl, 'tblGrid'), 'gridCol').map((c) => twipToPx(attrNum(c, 'w') ?? 0));

  const tblPr = child(tbl, 'tblPr');
  const defaultCellMar = readCellMargins(child(tblPr, 'tblCellMar'));

  const rows: (DocxTableCell | null)[][] = [];
  /** colIndex -> the origin cell currently being vertically merged. */
  const vmergeOrigin = new Map<number, DocxTableCell>();

  const trList = logicalChildrenNamed(tbl, 'tr');
  for (const tr of trList) {
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

      const cell = buildCell(tc, tcPr, colSpan, defaultCellMar, ctx);
      row.push(cell);
      for (let i = 1; i < colSpan; i++) row.push(null);

      if (vMergeVal === 'restart') vmergeOrigin.set(col, cell);
      else vmergeOrigin.delete(col);

      col += colSpan;
    }
    rows.push(row);
  }

  const widthPx = grid.reduce((a, b) => a + b, 0) || undefined;
  return { kind: 'table', colWidths: grid, rows, ...(widthPx ? { widthPx } : {}) };
}

function buildCell(
  tc: XmlNode,
  tcPr: XmlNode | undefined,
  colSpan: number,
  defaultCellMar: DocxTableCell['cellPaddingPx'],
  ctx: ParseContext,
): DocxTableCell {
  const content: DocxBlock[] = [];
  for (const node of logicalChildren(tc)) {
    const name = localName(node.name);
    if (name === 'p') content.push(...parseParagraph(node, ctx));
    else if (name === 'tbl') content.push(parseTable(node, ctx));
  }

  const cell: DocxTableCell = { content, rowSpan: 1, colSpan };

  const shdFill = attr(child(tcPr, 'shd'), 'fill');
  if (shdFill && shdFill !== 'auto') cell.fillHex = shdFill.replace(/^#/, '').toLowerCase();

  const borders = readCellBorders(child(tcPr, 'tcBorders'));
  if (borders) cell.borders = borders;

  const vAlign = attr(child(tcPr, 'vAlign'), 'val');
  if (vAlign === 'center' || vAlign === 'bottom') cell.vAlign = vAlign;
  else if (vAlign === 'top') cell.vAlign = 'top';

  const cellMar = readCellMargins(child(tcPr, 'tcMar')) ?? defaultCellMar;
  if (cellMar) cell.cellPaddingPx = cellMar;

  return cell;
}

function readCellMargins(mar: XmlNode | undefined): DocxTableCell['cellPaddingPx'] {
  if (!mar) return undefined;
  const side = (name: string) => {
    const w = attrNum(child(mar, name), 'w');
    return w !== undefined ? twipToPx(w) : 0;
  };
  return { top: side('top'), right: side('right'), bottom: side('bottom'), left: side('left') };
}

function readCellBorders(b: XmlNode | undefined): DocxTableCell['borders'] {
  if (!b) return undefined;
  const out: NonNullable<DocxTableCell['borders']> = {};
  let any = false;
  const sides: [keyof NonNullable<DocxTableCell['borders']>, string][] = [
    ['l', 'left'],
    ['t', 'top'],
    ['r', 'right'],
    ['b', 'bottom'],
  ];
  for (const [key, name] of sides) {
    const e = child(b, name);
    if (!e) continue;
    const raw: RawBorder = {
      sz: attrNum(e, 'sz'),
      colorHex: hexOrUndef(attr(e, 'color')),
      val: attr(e, 'val'),
    };
    const stroke: Stroke | undefined = borderToStroke(raw);
    if (stroke) {
      out[key] = stroke;
      any = true;
    }
  }
  return any ? out : undefined;
}

function hexOrUndef(v: string | undefined): string | undefined {
  return v && v !== 'auto' ? v.replace(/^#/, '').toLowerCase() : undefined;
}
