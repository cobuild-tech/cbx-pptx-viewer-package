/**
 * @cobuild-tech/pptx-viewer-core — framework-agnostic .pptx and .docx parser and DOM renderer.
 *
 * PPTX pipeline: read (OPC) -> parse (XML -> model) -> resolve (inheritance) -> render (DOM)
 * DOCX pipeline: read (OPC) -> parse (XML -> model) -> paginate (sections) -> render (DOM)
 *
 * PPTX usage:
 *   const deck = loadPptx(arrayBuffer);
 *   const viewer = createViewer(deck, containerEl, { fit: 'contain' });
 *   viewer.next(); viewer.goTo(3);
 *   // when done: viewer.destroy(); deck.dispose();
 *
 * DOCX usage:
 *   const doc = loadDocx(arrayBuffer);
 *   const viewer = createDocxViewer(doc, containerEl, { fit: 'width' });
 *   viewer.next(); viewer.goTo(2);
 *   // when done: viewer.destroy(); doc.dispose();
 */

// ─── PPTX ────────────────────────────────────────────────────────────────────
export { Deck, Deck as loadDeck } from './pptx/deck/deck.js';
export { Viewer, createViewer, type ViewerOptions } from './pptx/viewer/viewer.js';
export { renderSlide, type RenderDeps } from './pptx/render/dom.js';
export { installDeckFonts, type FontInstallation } from './pptx/render/fonts.js';
export { installWebFonts, collectFontFamilies, type WebFontOptions } from './pptx/render/webfonts.js';

import { Deck } from './pptx/deck/deck.js';
/** Load a .pptx from raw bytes into a renderable {@link Deck}. */
export function loadPptx(data: ArrayBuffer | Uint8Array): Deck {
  return Deck.load(data);
}

// ─── DOCX ────────────────────────────────────────────────────────────────────
export { DocxDocument } from './docx/document/document.js';
export { DocxViewer, createDocxViewer, type DocxViewerOptions } from './docx/viewer/viewer.js';
export { renderPage as renderDocxPage } from './docx/render/dom.js';
export { DocxRelType } from './docx/relTypes.js';

import { DocxDocument } from './docx/document/document.js';
/** Load a .docx from raw bytes into a renderable {@link DocxDocument}. */
export function loadDocx(data: ArrayBuffer | Uint8Array): DocxDocument {
  return DocxDocument.load(data);
}

export type {
  DocxPage,
  DocxBlock,
  DocxParagraph,
  DocxTable,
  DocxTableCell,
  DocxInlineImage,
  DocxPageSize,
  DocxPageMargins,
} from './docx/model.js';

// ─── Shared low-level building blocks ────────────────────────────────────────
export { OpcPackage, type Relationship } from './oxml/package.js';
export { RelType } from './pptx/relTypes.js';
export * as units from './oxml/units.js';
export {
  parseXml,
  localName,
  child,
  children,
  path,
  attr,
  attrNum,
  attrBool,
  type XmlNode,
} from './oxml/xml.js';
export * from './pptx/model.js';
