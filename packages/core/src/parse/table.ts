/**
 * Table parser: `<a:tbl>` -> {@link Table}.
 *
 * Builds a grid of cells with widths/heights, per-cell fill, borders and text.
 * Cells covered by a horizontal/vertical merge are emitted as `null` so the
 * renderer can skip them while the spanning cell uses colSpan/rowSpan.
 */
import { child, children, attrNum, attrBool, type XmlNode } from '../xml.js';
import { emuToPx } from '../units.js';
import type { Table, TableCell, Stroke } from '../model.js';
import type { ColorContext } from '../resolve/color.js';
import { type ParseScope, parseFill, strokeFromLn } from '../resolve/fill.js';
import { parseTextBody } from './text.js';
import { TextStyleChain } from '../resolve/textStyles.js';

export function parseTable(tbl: XmlNode, ctx: ColorContext, scope: ParseScope): Table {
  const colWidths = children(child(tbl, 'tblGrid'), 'gridCol').map(
    (c) => emuToPx(attrNum(c, 'w') ?? 0),
  );

  const rows: (TableCell | null)[][] = [];
  const rowHeights: number[] = [];

  for (const tr of children(tbl, 'tr')) {
    rowHeights.push(emuToPx(attrNum(tr, 'h') ?? 0));
    const row: (TableCell | null)[] = [];
    for (const tc of children(tr, 'tc')) {
      // Cells continuing a merge carry no content of their own.
      if (attrBool(tc, 'hMerge') || attrBool(tc, 'vMerge')) {
        row.push(null);
        continue;
      }
      row.push(parseCell(tc, ctx, scope));
    }
    rows.push(row);
  }

  return { colWidths, rowHeights, rows };
}

function parseCell(tc: XmlNode, ctx: ColorContext, scope: ParseScope): TableCell {
  const tcPr = child(tc, 'tcPr');
  const fill = parseFill(tcPr, scope) ?? { type: 'none' };

  const borders: TableCell['borders'] = {};
  const l = border(child(tcPr, 'lnL'), scope);
  const r = border(child(tcPr, 'lnR'), scope);
  const t = border(child(tcPr, 'lnT'), scope);
  const b = border(child(tcPr, 'lnB'), scope);
  if (l) borders.l = l;
  if (r) borders.r = r;
  if (t) borders.t = t;
  if (b) borders.b = b;

  const cell: TableCell = {
    fill,
    colSpan: attrNum(tc, 'gridSpan') ?? 1,
    rowSpan: attrNum(tc, 'rowSpan') ?? 1,
    borders,
  };

  const txBody = child(tc, 'txBody');
  if (txBody && children(txBody, 'p').length > 0) {
    const chain = new TextStyleChain(
      [child(txBody, 'lstStyle')].filter((n): n is XmlNode => Boolean(n)),
      ctx,
    );
    cell.text = parseTextBody(txBody, chain, ctx, scope);
  }
  return cell;
}

function border(ln: XmlNode | undefined, scope: ParseScope): Stroke | undefined {
  return strokeFromLn(ln, scope);
}
