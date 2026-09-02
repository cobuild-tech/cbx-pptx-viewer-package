/**
 * Write-back: put an edited value into the worksheet XML.
 *
 * Follows the same reuse-over-recreation rule as the other formats — an
 * existing `<c>` element is mutated in place, so its style index and any
 * property this parser never read survive the edit untouched. Only when a cell
 * (or its row) does not exist yet is a new element created, and then it is
 * spliced in at the right position: SpreadsheetML requires rows and cells in
 * ascending order, and Excel refuses a file where they are not.
 */
import {
  attr,
  child,
  children,
  createElement,
  insertInOrder,
  localName,
  type XmlNode,
} from '../../oxml/xml.js';

export { insertInOrder };
import { indexToColAlpha, parseCellRef } from '../sheets/sheet.js';
import type { CellInput } from './values.js';
import { isBlank } from './values.js';

/** CT_Worksheet's element order — new containers must respect it. */
const WORKSHEET_ORDER = [
  'sheetPr',
  'dimension',
  'sheetViews',
  'sheetFormatPr',
  'cols',
  'sheetData',
  'sheetCalcPr',
  'sheetProtection',
  'autoFilter',
  'mergeCells',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
];

/** Namespace prefix of a node's tag name ("x:c" -> "x:", "c" -> ""). */
function prefixOf(node: XmlNode): string {
  const i = node.name.indexOf(':');
  return i === -1 ? '' : node.name.slice(0, i + 1);
}

/** The `<sheetData>` element, created (in the right place) if absent. */
export function sheetDataOf(sheetXml: XmlNode): XmlNode {
  const existing = child(sheetXml, 'sheetData');
  if (existing) return existing;
  const created = createElement(`${prefixOf(sheetXml)}sheetData`);
  insertInOrder(sheetXml, created, WORKSHEET_ORDER);
  return created;
}

/** The `<row r="n">` for a 0-indexed row, created in ascending order if absent. */
function rowFor(sheetData: XmlNode, rowIndex: number, prefix: string): XmlNode {
  const want = rowIndex + 1;
  let insertAt = sheetData.children.length;
  for (let i = 0; i < sheetData.children.length; i++) {
    const node = sheetData.children[i]!;
    if (localName(node.name) !== 'row') continue;
    const r = Number(attr(node, 'r'));
    if (r === want) return node;
    if (Number.isFinite(r) && r > want) {
      insertAt = i;
      break;
    }
    insertAt = i + 1;
  }
  const created = createElement(`${prefix}row`, { r: String(want) });
  sheetData.children.splice(insertAt, 0, created);
  return created;
}

/** The `<c r="A1">` in a row, created in ascending column order if absent. */
function cellFor(row: XmlNode, ref: string, colIndex: number, prefix: string): XmlNode {
  let insertAt = row.children.length;
  for (let i = 0; i < row.children.length; i++) {
    const node = row.children[i]!;
    if (localName(node.name) !== 'c') continue;
    const nodeRef = attr(node, 'r');
    if (nodeRef === ref) return node;
    const col = nodeRef ? parseCellRef(nodeRef).colIndex : -1;
    if (col > colIndex) {
      insertAt = i;
      break;
    }
    insertAt = i + 1;
  }
  const created = createElement(`${prefix}c`, { r: ref });
  row.children.splice(insertAt, 0, created);
  return created;
}

/** Find a cell's `<c>` element without creating anything. */
export function findCell(sheetXml: XmlNode, ref: string): XmlNode | undefined {
  const sheetData = child(sheetXml, 'sheetData');
  if (!sheetData) return undefined;
  const { rowIndex } = parseCellRef(ref);
  for (const row of children(sheetData, 'row')) {
    if (Number(attr(row, 'r')) !== rowIndex + 1) continue;
    return children(row, 'c').find((c) => attr(c, 'r') === ref);
  }
  return undefined;
}

/** Find, or create, a cell's `<c>` element. */
export function ensureCell(sheetXml: XmlNode, ref: string): XmlNode {
  const prefix = prefixOf(sheetXml);
  const { rowIndex, colIndex } = parseCellRef(ref);
  const sheetData = sheetDataOf(sheetXml);
  const row = rowFor(sheetData, rowIndex, prefix);
  return cellFor(row, ref, colIndex, prefix);
}

/**
 * Write a value into a cell, creating the cell (and its row) if needed.
 *
 * The cell's style index and unread attributes are preserved; only the type
 * attribute and the value children (`<f>`, `<v>`, `<is>`) are rewritten.
 */
export function writeCell(sheetXml: XmlNode, ref: string, input: CellInput): XmlNode {
  const prefix = prefixOf(sheetXml);
  const cell = ensureCell(sheetXml, ref);

  // Drop the old value; everything else on the element stays as authored.
  cell.children = cell.children.filter((c) => !['f', 'v', 'is'].includes(localName(c.name)));
  cell.text = '';
  delete cell.attrs['t'];
  for (const key of Object.keys(cell.attrs)) {
    if (localName(key) === 't') delete cell.attrs[key];
  }

  if (input.formula !== undefined) {
    // No cached <v>: the workbook is flagged for a full recalculation on load,
    // so Excel fills the result in itself rather than trusting a stale one.
    cell.children.push(createElement(`${prefix}f`, {}, [], input.formula));
    return cell;
  }

  if (isBlank(input)) return cell;

  if (input.type === 'inlineStr') {
    cell.attrs['t'] = 'inlineStr';
    const text = createElement(`${prefix}t`, {}, [], input.rawValue);
    if (input.rawValue !== input.rawValue.trim()) text.attrs['xml:space'] = 'preserve';
    cell.children.push(createElement(`${prefix}is`, {}, [text]));
    return cell;
  }

  if (input.type !== 'n') cell.attrs['t'] = input.type;
  cell.children.push(createElement(`${prefix}v`, {}, [], input.rawValue));
  return cell;
}

/** Point a cell at a style index, creating the cell if needed. */
export function writeCellStyle(sheetXml: XmlNode, ref: string, styleId: number): XmlNode {
  const cell = ensureCell(sheetXml, ref);
  cell.attrs['s'] = String(styleId);
  return cell;
}

/**
 * Widen `<dimension ref>` so it covers `ref`. Excel tolerates a stale
 * dimension, but keeping it honest keeps the file identical to one Excel
 * would have written.
 */
export function growDimension(sheetXml: XmlNode, ref: string): void {
  const dim = child(sheetXml, 'dimension');
  const current = dim && attr(dim, 'ref');
  if (!dim || !current) return;

  const [from, to] = current.split(':');
  const start = parseCellRef(from ?? 'A1');
  const end = parseCellRef(to ?? from ?? 'A1');
  const cell = parseCellRef(ref);

  const minRow = Math.min(start.rowIndex, cell.rowIndex);
  const minCol = Math.min(start.colIndex, cell.colIndex);
  const maxRow = Math.max(end.rowIndex, cell.rowIndex);
  const maxCol = Math.max(end.colIndex, cell.colIndex);
  dim.attrs['ref'] =
    `${indexToColAlpha(minCol)}${minRow + 1}:${indexToColAlpha(maxCol)}${maxRow + 1}`;
}

/**
 * Ensure the workbook recalculates when Excel opens it. Any edit can invalidate
 * a cached formula result elsewhere in the book, and we deliberately do not
 * evaluate formulas ourselves — Excel does it on load instead.
 *
 * Returns true if the workbook XML was changed.
 */
export function setFullCalcOnLoad(workbookXml: XmlNode): boolean {
  const prefix = prefixOf(workbookXml);
  let calcPr = child(workbookXml, 'calcPr');
  if (!calcPr) {
    calcPr = createElement(`${prefix}calcPr`, { calcId: '0' });
    // calcPr sits after sheets/definedNames in CT_Workbook; appending before
    // any extLst is close enough for a schema that only orders the leading
    // elements strictly.
    const extAt = workbookXml.children.findIndex((c) => localName(c.name) === 'extLst');
    if (extAt === -1) workbookXml.children.push(calcPr);
    else workbookXml.children.splice(extAt, 0, calcPr);
  } else if (attr(calcPr, 'fullCalcOnLoad') === '1') {
    return false;
  }
  calcPr.attrs['fullCalcOnLoad'] = '1';
  return true;
}
