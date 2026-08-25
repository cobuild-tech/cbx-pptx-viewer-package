/**
 * @cobuildx.ai/office-viewer — framework-agnostic .pptx, .docx and .xlsx parser
 * and DOM renderer. React bindings are available from
 * '@cobuildx.ai/office-viewer/react'.
 *
 * PPTX pipeline: read (OPC) -> parse (XML -> model) -> resolve (inheritance) -> render (DOM)
 * DOCX pipeline: read (OPC) -> parse (XML -> model) -> paginate (sections)    -> render (DOM)
 *
 * PPTX usage:
 *   const deck = loadPptx(arrayBuffer);
 *   const viewer = createViewer(deck, containerEl, { fit: 'contain' });
 *   viewer.next(); viewer.goTo(3);
 *   // when done: viewer.destroy(); deck.dispose();
 *
 * PPTX editing (text):
 *   const viewer = createViewer(deck, containerEl, { editable: true });
 *   viewer.applyFormat({ bold: true });   // formats the current selection
 *   viewer.undo(); viewer.redo();
 *   const blob = viewer.exportBlob();     // a valid .pptx with the edits
 *
 * DOCX usage:
 *   const doc = loadDocx(arrayBuffer);
 *   const viewer = createDocxViewer(doc, containerEl, { fit: 'width' });
 *   // when done: viewer.destroy(); doc.dispose();
 *
 * XLSX usage:
 *   const wb = loadXlsx(arrayBuffer);
 *   const viewer = createXlsxViewer(wb, containerEl, { editable: true });
 *   viewer.applyFormat({ bold: true });   // formats the selected cells
 *   const blob = viewer.exportBlob();     // a valid .xlsx with the edits
 *
 * DOCX editing (text): pass `editable`, which renders the document as one
 * continuous column instead of fixed pages (see docx/edit/flow.ts).
 *   const viewer = createDocxViewer(doc, containerEl, { editable: true });
 *   viewer.applyFormat({ bold: true });
 *   const blob = viewer.exportBlob();
 */

// ─── PPTX ────────────────────────────────────────────────────────────────────
export { Deck, Deck as loadDeck } from './pptx/deck/deck.js';
export { Viewer, createViewer, type ViewerOptions } from './pptx/viewer/viewer.js';
export { renderSlide, type RenderDeps } from './pptx/render/dom.js';
export { installDeckFonts, type FontInstallation } from './pptx/render/fonts.js';
export { installWebFonts, collectFontFamilies, type WebFontOptions } from './pptx/render/webfonts.js';

// PPTX text editing (opt in with `editable` on the viewer).
export { EditSession, type EditSessionOptions } from './pptx/edit/session.js';
export { EditContext } from './pptx/edit/context.js';
export { reconcileTextBody } from './pptx/edit/reconcile.js';
export { writeTextBody, type ParaEdit, type Segment } from './pptx/edit/xmlWrite.js';
export { readFormat } from './pptx/edit/format.js';
export { EDIT_ATTR } from './oxml/edit/attrs.js';

// Format-agnostic editing primitives, shared by every format slice.
export {
  applyFormatToSelection,
  formatAtSelection,
  bodyElementOf,
  type Resolver,
  type ReadRunFormat,
} from './oxml/edit/selection.js';
export { mergeFormat, isEmptyFormat, type RunFormat } from './oxml/edit/format.js';
export { installEditStyles, type TextBoxOutline } from './oxml/edit/styles.js';
export { History, type Snapshot } from './oxml/edit/history.js';
export type { ModelSource } from './pptx/deck/deck.js';
export type { EditRenderContext } from './pptx/render/primitives.js';

import { Deck } from './pptx/deck/deck.js';
/** Load a .pptx from raw bytes into a renderable {@link Deck}. */
export function loadPptx(data: ArrayBuffer | Uint8Array): Deck {
  return Deck.load(data);
}

// ─── DOCX ────────────────────────────────────────────────────────────────────
export { DocxDocument } from './docx/document/document.js';
export {
  DocxViewer,
  createDocxViewer,
  type DocxViewerOptions,
} from './docx/viewer/viewer.js';
export { renderPage as renderDocxPage, type RenderDeps as DocxRenderDeps } from './docx/render/dom.js';
export { DocxRelType } from './docx/relTypes.js';
export type {
  DocxSection,
  DocxPage,
  DocxBlock,
  DocxParagraph,
  DocxTable,
  DocxTableCell,
  DocxInlineImage,
  DocxRun,
  DocxPageSize,
  DocxPageMargins,
} from './docx/model.js';

// DOCX text editing (opt in with `editable` on the viewer).
export { DocxEditSession, type DocxEditSessionOptions } from './docx/edit/session.js';
export { DocxEditContext, type DocxEditRenderContext } from './docx/edit/context.js';
export { reconcileParagraph } from './docx/edit/reconcile.js';
export {
  writeParagraphs,
  type DocxParaEdit,
  type DocxSegment,
} from './docx/edit/xmlWrite.js';
export { renderFlow, FLOW_CLASS } from './docx/edit/flow.js';
export type { DocxSource } from './docx/document/context.js';

import { DocxDocument } from './docx/document/document.js';
/** Load a .docx from raw bytes into a renderable {@link DocxDocument}. */
export function loadDocx(data: ArrayBuffer | Uint8Array): DocxDocument {
  return DocxDocument.load(data);
}

// ─── XLSX ────────────────────────────────────────────────────────────────────
export { Workbook } from './xlsx/workbook/workbook.js';
export {
  XlsxViewer,
  createXlsxViewer,
  type XlsxViewerOptions,
} from './xlsx/viewer/viewer.js';
export {
  renderXlsxSheet,
  createSheetView,
  XLSX_CELL_ATTR,
  type RenderXlsxOptions,
  type XlsxSheetView,
} from './xlsx/render/dom.js';
export { XlsxRelType } from './xlsx/relTypes.js';
export type {
  XlsxSheet,
  XlsxSheetSummary,
  XlsxRow,
  XlsxColumn,
  XlsxCell,
  XlsxMergeCell,
  XlsxCellStyle,
} from './xlsx/model.js';

// XLSX cell editing (opt in with `editable` on the viewer).
export { XlsxEditSession, type XlsxEditSessionOptions } from './xlsx/edit/session.js';
export { XlsxEditContext, type XlsxEditRenderContext } from './xlsx/edit/context.js';
export {
  parseInput,
  editableText,
  isBlank,
  type CellInput,
} from './xlsx/edit/values.js';
export {
  writeCell,
  writeCellStyle,
  ensureCell,
  findCell,
  growDimension,
  setFullCalcOnLoad,
} from './xlsx/edit/xmlWrite.js';
export { StyleWriter, isEmptyPatch, type CellFormatPatch } from './xlsx/edit/styleWrite.js';
export { readCellFormat, readCellPatch } from './xlsx/edit/format.js';
export type { XlsxSource } from './xlsx/workbook/workbook.js';

import { Workbook } from './xlsx/workbook/workbook.js';
/** Load a .xlsx from raw bytes into a renderable {@link Workbook}. */
export function loadXlsx(data: ArrayBuffer | Uint8Array): Workbook {
  return Workbook.load(data);
}

// ─── Shared low-level building blocks ────────────────────────────────────────
export { OpcPackage, type Relationship } from './oxml/package.js';
export { RelType } from './pptx/relTypes.js';
export * as units from './oxml/units.js';
export {
  parseXml,
  serializeXml,
  serializeNode,
  localName,
  child,
  children,
  path,
  attr,
  attrNum,
  attrBool,
  createElement,
  cloneNode,
  resolveIndexPath,
  type XmlNode,
} from './oxml/xml.js';
export * from './pptx/model.js';

