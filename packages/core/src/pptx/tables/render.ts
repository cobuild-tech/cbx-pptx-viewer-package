/**
 * Table renderer: {@link Table} -> an HTML `<table>`.
 *
 * Column widths come from the grid; row heights are treated as a minimum (the
 * row grows to fit its content, matching PowerPoint). Cell text renders in
 * normal flow so the cell expands vertically.
 */
import type { Table, TableCell, Stroke, TextBody } from '../model.js';
import { colorToCss } from '../color.js';
import { applyFillBackground, type RenderDeps } from '../render/primitives.js';
import { renderTextBody, anchorToValign } from '../text/render.js';

const SVG_NS = 'http://www.w3.org/2000/svg';

export function renderTable(table: Table, deps: RenderDeps): HTMLTableElement {
  const tbl = document.createElement('table');
  tbl.style.borderCollapse = 'collapse';
  tbl.style.tableLayout = 'fixed';
  tbl.style.width = '100%';

  const colgroup = document.createElement('colgroup');
  for (const w of table.colWidths) {
    const col = document.createElement('col');
    col.style.width = `${w}px`;
    colgroup.appendChild(col);
  }
  tbl.appendChild(colgroup);

  table.rows.forEach((row, r) => {
    const tr = document.createElement('tr');
    // The XML row height is a minimum; the row grows to fit its content,
    // matching PowerPoint (cell text is rendered in normal flow below).
    tr.style.height = `${table.rowHeights[r] ?? 0}px`;
    for (const cell of row) {
      if (cell === null) continue;
      const td = document.createElement('td');
      td.style.padding = '0';
      td.style.verticalAlign = cell.text && !cell.text.vert ? anchorToValign(cell.text.anchor) : 'top';
      if (cell.colSpan > 1) td.colSpan = cell.colSpan;
      if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
      applyFillBackground(td, cell.fill, deps);
      const b = cell.borders;
      if (b.l) td.style.borderLeft = cssBorder(b.l);
      if (b.r) td.style.borderRight = cssBorder(b.r);
      if (b.t) td.style.borderTop = cssBorder(b.t);
      if (b.b) td.style.borderBottom = cssBorder(b.b);
      if (cell.diagonals) {
        td.style.position = 'relative';
        td.appendChild(diagonals(cell.diagonals));
      }
      if (cell.text) {
        const box = renderTextBody(cell.text, deps, true);
        if (cell.text.vert) {
          td.style.verticalAlign = 'middle';
          placeVertical(box, cell.text, spannedHeight(table, r, cell));
        }
        td.appendChild(box);
      }
      tr.appendChild(td);
    }
    tbl.appendChild(tr);
  });
  return tbl;
}

/**
 * Vertical text in a cell wraps at the cell's height, which a table cell does
 * not know before layout — so the box is given the row's height outright
 * (centred, should the row grow). Its anchor runs across the cell: the top of
 * text turned clockwise faces right, of text turned anticlockwise left.
 */
function placeVertical(box: HTMLElement, text: TextBody, height: number): void {
  box.style.height = `${height}px`;
  const toRight = text.vert === 'vert270' ? text.anchor === 'bottom' : text.anchor === 'top';
  const toLeft = text.vert === 'vert270' ? text.anchor === 'top' : text.anchor === 'bottom';
  box.style.marginLeft = toLeft ? '0' : 'auto';
  box.style.marginRight = toRight ? '0' : 'auto';
}

/** Total stated height of the rows a cell spans. */
function spannedHeight(table: Table, r: number, cell: TableCell): number {
  let h = 0;
  for (let i = r; i < r + cell.rowSpan; i++) h += table.rowHeights[i] ?? 0;
  return h;
}

/** The cell's diagonal rules, as an SVG stretched over it. */
function diagonals(d: NonNullable<TableCell['diagonals']>): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('preserveAspectRatio', 'none');
  svg.setAttribute('viewBox', '0 0 100 100');
  Object.assign(svg.style, { position: 'absolute', inset: '0', width: '100%', height: '100%', overflow: 'visible' });
  svg.style.pointerEvents = 'none';
  const line = (s: Stroke, y1: number, y2: number) => {
    const l = document.createElementNS(SVG_NS, 'line');
    for (const [k, v] of Object.entries({ x1: '0', y1: `${y1}`, x2: '100', y2: `${y2}` })) l.setAttribute(k, v);
    l.setAttribute('stroke', colorToCss(s.color));
    l.setAttribute('stroke-width', `${s.width}`);
    // Keep the width in px however the box stretches the viewBox.
    l.setAttribute('vector-effect', 'non-scaling-stroke');
    if (s.dash) l.setAttribute('stroke-dasharray', s.dash.join(','));
    svg.appendChild(l);
  };
  if (d.down) line(d.down, 0, 100);
  if (d.up) line(d.up, 100, 0);
  return svg;
}

function cssBorder(s: Stroke): string {
  // CSS has only two dash styles: a dash no longer than the line is wide reads
  // as dots, anything longer as dashes.
  const style = !s.dash ? 'solid' : (s.dash[0] ?? 0) <= s.width * 1.01 ? 'dotted' : 'dashed';
  return `${s.width}px ${style} ${colorToCss(s.color)}`;
}
