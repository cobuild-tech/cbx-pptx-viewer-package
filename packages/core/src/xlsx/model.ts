/**
 * Render-agnostic intermediate model for SpreadsheetML (.xlsx) documents.
 */

export interface XlsxFont {
  name?: string;
  sizePt?: number;
  colorHex?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
}

export interface XlsxFill {
  patternType?: string;
  fgColorHex?: string;
  bgColorHex?: string;
}

export type BorderStyle = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'none';

export interface XlsxBorderSide {
  style?: BorderStyle;
  colorHex?: string;
}

export interface XlsxBorder {
  top?: XlsxBorderSide;
  bottom?: XlsxBorderSide;
  left?: XlsxBorderSide;
  right?: XlsxBorderSide;
}

export interface XlsxAlignment {
  horizontal?: 'left' | 'center' | 'right' | 'justify' | 'fill';
  vertical?: 'top' | 'center' | 'bottom';
  wrapText?: boolean;
}

export interface XlsxCellStyle {
  font?: XlsxFont;
  fill?: XlsxFill;
  border?: XlsxBorder;
  alignment?: XlsxAlignment;
  numFmtId?: number;
  numFmtCode?: string;
}

export type CellType = 's' | 'n' | 'b' | 'e' | 'str' | 'inlineStr';

export interface XlsxCell {
  /** Cell reference, e.g. "A1" */
  ref: string;
  /** 0-indexed column index */
  colIndex: number;
  /** 0-indexed row index */
  rowIndex: number;
  /** Value type */
  type: CellType;
  /** Raw string value stored in XML or shared strings */
  rawValue: string;
  /** Renderable formatted text value */
  formattedValue: string;
  /** Cell formula if present (without leading '=') */
  formula?: string;
  /** Cell style index */
  styleId?: number;
  /** Resolved cell style properties */
  style?: XlsxCellStyle;
}

export interface XlsxMergeCell {
  /** Cell range reference, e.g. "A1:C3" */
  ref: string;
  /** 0-indexed top row */
  startRow: number;
  /** 0-indexed left col */
  startCol: number;
  /** 0-indexed bottom row */
  endRow: number;
  /** 0-indexed right col */
  endCol: number;
}

export interface XlsxColumn {
  /** 1-indexed start column */
  min: number;
  /** 1-indexed end column */
  max: number;
  /** Explicit column width in CSS pixels if defined */
  widthPx?: number;
  /** True if column is hidden */
  hidden?: boolean;
}

export interface XlsxRow {
  /** 0-indexed row index */
  rowIndex: number;
  /** Explicit row height in CSS pixels if defined */
  heightPx?: number;
  /** True if row is hidden */
  hidden?: boolean;
  /** Map of colIndex -> Cell */
  cells: Map<number, XlsxCell>;
}

export interface XlsxSheetSummary {
  id: string;
  name: string;
  rId: string;
  targetPath?: string;
}

export interface XlsxSheet {
  id: string;
  name: string;
  maxRow: number;
  maxCol: number;
  rows: Map<number, XlsxRow>;
  cols: XlsxColumn[];
  mergeCells: XlsxMergeCell[];
}
