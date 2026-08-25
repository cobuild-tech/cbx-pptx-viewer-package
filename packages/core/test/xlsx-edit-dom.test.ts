/**
 * @vitest-environment jsdom
 *
 * The DOM half of XLSX editing: grid markers, selection, keyboard navigation
 * and the viewer's commit / undo cycle.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { Workbook } from '../src/xlsx/workbook/workbook.js';
import { XlsxRelType } from '../src/xlsx/relTypes.js';
import { XlsxViewer } from '../src/xlsx/viewer/viewer.js';
import { XLSX_CELL_ATTR } from '../src/xlsx/render/dom.js';

const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

function buildXlsx(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
      </Types>`,
    ),
    '_rels/.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${XlsxRelType.OfficeDocument}" Target="xl/workbook.xml"/>`),
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${XlsxRelType.Worksheet}" Target="worksheets/sheet1.xml"/>`),
    ),
    'xl/workbook.xml': strToU8(
      `<workbook xmlns="${SML}" xmlns:r="${OFFICE}">
        <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
      </workbook>`,
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      `<worksheet xmlns="${SML}"><dimension ref="A1:B2"/><sheetData>
        <row r="1"><c r="A1" t="inlineStr"><is><t>Quarter</t></is></c><c r="B1"><v>15000</v></c></row>
        <row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><f>B1*2</f><v>30000</v></c></row>
      </sheetData></worksheet>`,
    ),
  });
}

function td(container: HTMLElement, ref: string): HTMLTableCellElement {
  return container.querySelector(`[${XLSX_CELL_ATTR}="${ref}"]`) as HTMLTableCellElement;
}

function grid(container: HTMLElement): HTMLElement {
  return container.querySelector('.cbx-xlsx-grid-container') as HTMLElement;
}

function key(el: HTMLElement, k: string, init: KeyboardEventInit = {}): void {
  el.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, ...init }));
}

function click(el: HTMLElement, init: MouseEventInit = {}): void {
  el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, ...init }));
}

describe('editable grid', () => {
  let container: HTMLElement;
  let workbook: Workbook;
  let viewer: XlsxViewer;

  beforeEach(() => {
    container = document.createElement('div');
    document.body.replaceChildren(container);
    workbook = Workbook.load(buildXlsx());
    viewer = new XlsxViewer(workbook, container, { editable: true });
  });

  it('marks every cell with its reference and renders the values', () => {
    expect(td(container, 'A1').textContent).toBe('Quarter');
    expect(td(container, 'B1').textContent).toBe('15000');
    // Blank headroom past the data is rendered so the sheet can be grown.
    expect(td(container, 'A20')).toBeTruthy();
  });

  it('shows the formula, not the cached result, in the formula bar', () => {
    click(td(container, 'B2'));
    expect(container.querySelector('.cbx-xlsx-name-box')?.textContent).toBe('B2');
    expect(container.querySelector('.cbx-xlsx-fx-input')?.textContent).toBe('=B1*2');
  });

  it('starts editing on a printable key and commits on Enter', () => {
    click(td(container, 'A2'));
    key(grid(container), 'x');
    const cell = td(container, 'A2');
    expect(cell.getAttribute('contenteditable')).toBe('true');
    expect(cell.textContent).toBe('x');

    cell.textContent = 'Q2';
    key(grid(container), 'Enter');

    expect(td(container, 'A2').textContent).toBe('Q2');
    expect(td(container, 'A2').getAttribute('contenteditable')).toBeNull();
    expect(workbook.getSheet(0)?.rows.get(1)?.cells.get(0)?.formattedValue).toBe('Q2');
  });

  it('moves the selection down after committing with Enter', () => {
    click(td(container, 'A2'));
    key(grid(container), 'F2');
    td(container, 'A2').textContent = 'Q2';
    key(grid(container), 'Enter');
    expect(container.querySelector('.cbx-xlsx-name-box')?.textContent).toBe('A3');
  });

  it('abandons an edit on Escape', () => {
    click(td(container, 'A2'));
    key(grid(container), 'F2');
    td(container, 'A2').textContent = 'nope';
    key(grid(container), 'Escape');
    expect(td(container, 'A2').textContent).toBe('Q1');
    expect(workbook.getSheet(0)?.rows.get(1)?.cells.get(0)?.formattedValue).toBe('Q1');
  });

  it('navigates with the arrow keys and extends with shift', () => {
    click(td(container, 'A1'));
    key(grid(container), 'ArrowDown');
    expect(container.querySelector('.cbx-xlsx-name-box')?.textContent).toBe('A2');
    key(grid(container), 'ArrowRight', { shiftKey: true });
    expect(container.querySelector('.cbx-xlsx-name-box')?.textContent).toBe('B2');
  });

  it('clears the selected cells with Delete', () => {
    click(td(container, 'A1'));
    click(td(container, 'A2'), { shiftKey: true });
    key(grid(container), 'Delete');
    expect(workbook.getSheet(0)?.rows.get(0)?.cells.get(0)?.formattedValue).toBe('');
    expect(workbook.getSheet(0)?.rows.get(1)?.cells.get(0)?.formattedValue).toBe('');
  });

  it('undoes and redoes a committed edit', () => {
    click(td(container, 'A2'));
    key(grid(container), 'F2');
    td(container, 'A2').textContent = 'Q2';
    key(grid(container), 'Enter');
    expect(viewer.canUndo).toBe(true);

    viewer.undo();
    expect(td(container, 'A2').textContent).toBe('Q1');
    viewer.redo();
    expect(td(container, 'A2').textContent).toBe('Q2');
  });

  it('formats the selected range through the viewer', () => {
    click(td(container, 'A1'));
    click(td(container, 'A2'), { shiftKey: true });
    expect(viewer.applyFormat({ bold: true })).toBe(true);

    expect(td(container, 'A1').style.fontWeight).toBe('bold');
    expect(td(container, 'A2').style.fontWeight).toBe('bold');
    expect(workbook.getSheet(0)?.rows.get(0)?.cells.get(0)?.style?.font?.bold).toBe(true);
  });

  it('exports a workbook that re-loads with the edit applied', () => {
    click(td(container, 'A2'));
    key(grid(container), 'F2');
    td(container, 'A2').textContent = 'Q2';
    key(grid(container), 'Enter');

    const reloaded = Workbook.load(workbook.toBytes());
    expect(reloaded.getSheet(0)?.rows.get(1)?.cells.get(0)?.formattedValue).toBe('Q2');
  });
});

describe('read-only grid', () => {
  it('does not make cells editable without `editable`', () => {
    const container = document.createElement('div');
    document.body.replaceChildren(container);
    const workbook = Workbook.load(buildXlsx());
    new XlsxViewer(workbook, container);

    const cell = td(container, 'A2');
    click(cell);
    cell.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    expect(cell.getAttribute('contenteditable')).toBeNull();
    expect(workbook.hasEdits).toBe(false);
  });
});
