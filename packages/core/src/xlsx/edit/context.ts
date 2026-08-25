/**
 * XlsxEditContext — what the grid renderer needs in order to make cells
 * editable, and the rules for what may not be.
 *
 * A spreadsheet has no layout/master inheritance to protect (as PPTX does) and
 * no separate header part (as DOCX does), but it does have three things that
 * must stay read-only, and each is answerable straight from the XML the cell
 * was parsed from:
 *
 *  - a protected sheet (`<sheetProtection sheet="1">`), which is the file
 *    saying so explicitly;
 *  - a cell covered by — but not the origin of — a merged range, which has no
 *    value of its own;
 *  - a cell hosting an array or shared formula, since overwriting it would
 *    silently break every cell that refers back to it.
 */
import { attr, attrBool, child, type XmlNode } from '../../oxml/xml.js';
import type { XlsxCell, XlsxSheet } from '../model.js';
import { indexToColAlpha, parseCellRef } from '../sheets/sheet.js';
import type { Workbook } from '../workbook/workbook.js';

/** What the grid renderer needs to know. Absent for a read-only render. */
export interface XlsxEditRenderContext {
  /** True if the cell at this reference may be typed into. */
  editable(ref: string): boolean;
}

export class XlsxEditContext implements XlsxEditRenderContext {
  private readonly workbook: Workbook;
  private sheet: XlsxSheet;
  private covered = new Set<string>();
  private protectedSheet = false;

  constructor(workbook: Workbook, sheet: XlsxSheet) {
    this.workbook = workbook;
    this.sheet = sheet;
    this.retarget(sheet);
  }

  /** Point the context at a freshly parsed sheet (after a commit or undo). */
  retarget(sheet: XlsxSheet): void {
    this.sheet = sheet;
    this.covered = new Set();
    for (const m of sheet.mergeCells) {
      for (let r = m.startRow; r <= m.endRow; r++) {
        for (let c = m.startCol; c <= m.endCol; c++) {
          if (r === m.startRow && c === m.startCol) continue;
          this.covered.add(`${indexToColAlpha(c)}${r + 1}`);
        }
      }
    }
    const xml = this.workbook.sheetXml(sheet.id);
    const protection = xml ? child(xml, 'sheetProtection') : undefined;
    this.protectedSheet = protection ? attrBool(protection, 'sheet', true) : false;
  }

  /** True if the whole sheet is locked by `<sheetProtection>`. */
  get isProtected(): boolean {
    return this.protectedSheet;
  }

  editable(ref: string): boolean {
    if (this.protectedSheet) return false;
    if (this.covered.has(ref)) return false;
    const cell = this.cellAt(ref);
    return !cell || !isGeneratedFormula(this.workbook.sourceOf(cell)?.node);
  }

  /** The parsed cell at a reference, if the sheet has one. */
  cellAt(ref: string): XlsxCell | undefined {
    const { rowIndex, colIndex } = parseCellRef(ref);
    return this.sheet.rows.get(rowIndex)?.cells.get(colIndex);
  }
}

/**
 * True for the *host* of an array or shared formula — the cell whose `<f>`
 * carries the range other cells derive their formula from.
 */
function isGeneratedFormula(cellNode: XmlNode | undefined): boolean {
  const f = cellNode ? child(cellNode, 'f') : undefined;
  if (!f) return false;
  const type = attr(f, 't');
  if (type === 'array' || type === 'dataTable') return true;
  return type === 'shared' && attr(f, 'ref') !== undefined;
}
