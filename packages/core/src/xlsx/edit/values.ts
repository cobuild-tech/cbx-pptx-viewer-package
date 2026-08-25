/**
 * The value layer: what the user typed <-> what SpreadsheetML stores.
 *
 * Excel decides a cell's type from what was typed, and so do we: a leading '='
 * is a formula, a numeric literal is a number, TRUE/FALSE is a boolean, an
 * error literal is an error, and anything else is text. Text is written as an
 * *inline* string (`t="inlineStr"`) rather than pushed into `xl/sharedStrings.xml`
 * — inline strings are valid SpreadsheetML everywhere, and it keeps an edit to
 * one cell from rewriting a part every other sheet shares.
 */
import type { CellType, XlsxCell } from '../model.js';

/** A cell value ready to be written into the XML. */
export interface CellInput {
  type: CellType;
  /** Stored value; empty for a blank cell or an uncached formula result. */
  rawValue: string;
  /** Formula text without the leading '='. */
  formula?: string;
}

const NUMBER_RE = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;
const ERRORS = new Set(['#N/A', '#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#NULL!', '#NUM!']);

/** Interpret a raw user-entered string the way Excel's formula bar does. */
export function parseInput(text: string): CellInput {
  if (text === '') return { type: 'n', rawValue: '' };

  if (text.startsWith('=') && text.length > 1) {
    return { type: 'n', rawValue: '', formula: text.slice(1) };
  }

  const trimmed = text.trim();
  if (NUMBER_RE.test(trimmed)) {
    return { type: 'n', rawValue: String(Number(trimmed)) };
  }

  const upper = trimmed.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') {
    return { type: 'b', rawValue: upper === 'TRUE' ? '1' : '0' };
  }
  if (ERRORS.has(upper)) {
    return { type: 'e', rawValue: upper };
  }

  return { type: 'inlineStr', rawValue: text };
}

/** True if the input leaves the cell empty. */
export function isBlank(input: CellInput): boolean {
  return input.formula === undefined && input.rawValue === '';
}

/**
 * The text to put in front of the user when they edit a cell: the formula if
 * there is one, otherwise the underlying value rather than its formatted form,
 * so re-committing an untouched cell is a no-op.
 */
export function editableText(cell: XlsxCell | undefined): string {
  if (!cell) return '';
  if (cell.formula) return `=${cell.formula}`;
  if (cell.type === 's' || cell.type === 'inlineStr' || cell.type === 'str') {
    return cell.formattedValue;
  }
  return cell.rawValue;
}
