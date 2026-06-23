/**
 * Table renderer: {@link Table} -> an HTML `<table>`.
 *
 * Column widths come from the grid; row heights are treated as a minimum (the
 * row grows to fit its content, matching PowerPoint). Cell text renders in
 * normal flow so the cell expands vertically.
 */
import type { Table, Stroke } from '../model.js';
import { colorToCss } from '../color.js';
import { applyFillBackground, type RenderDeps } from '../render/primitives.js';
import { renderTextBody, anchorToValign } from '../text/render.js';

export function renderTable(
  table: Table,
  deps: RenderDeps,
  sx = 1,
  sy = 1,
): HTMLTableElement {
  const tbl = document.createElement('table');
  tbl.style.borderCollapse = 'collapse';
  tbl.style.tableLayout = 'fixed';
  tbl.style.width = '100%';

  const colgroup = document.createElement('colgroup');
  for (const w of table.colWidths) {
    const col = document.createElement('col');
    col.style.width = `${w * sx}px`;
    colgroup.appendChild(col);
  }
  tbl.appendChild(colgroup);

  table.rows.forEach((row, r) => {
    const tr = document.createElement('tr');
    // The XML row height is a minimum; the row grows to fit its content,
    // matching PowerPoint (cell text is rendered in normal flow below).
    tr.style.height = `${(table.rowHeights[r] ?? 0) * sy}px`;
    for (const cell of row) {
      if (cell === null) continue;
      const td = document.createElement('td');
      td.style.padding = '0';
      td.style.verticalAlign = cell.text ? anchorToValign(cell.text.anchor) : 'top';
      if (cell.colSpan > 1) td.colSpan = cell.colSpan;
      if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
      applyFillBackground(td, cell.fill, deps);
      const b = cell.borders;
      if (b.l) td.style.borderLeft = cssBorder(b.l, sx, sy);
      if (b.r) td.style.borderRight = cssBorder(b.r, sx, sy);
      if (b.t) td.style.borderTop = cssBorder(b.t, sx, sy);
      if (b.b) td.style.borderBottom = cssBorder(b.b, sx, sy);
      if (cell.text) td.appendChild(renderTextBody(cell.text, deps, true, sx, sy));
      tr.appendChild(td);
    }
    tbl.appendChild(tr);
  });
  return tbl;
}

function cssBorder(s: Stroke, sx = 1, sy = 1): string {
  const scale = (sx + sy) / 2;
  return `${s.width * scale}px solid ${colorToCss(s.color)}`;
}
