/**
 * @pptx-viewer/core — framework-agnostic .pptx parser and DOM renderer.
 *
 * Pipeline: read (OPC) -> parse (XML -> model) -> resolve (inheritance)
 * -> render (DOM). Typical usage:
 *
 *   const deck = loadPptx(arrayBuffer);
 *   const viewer = createViewer(deck, containerEl, { fit: 'contain' });
 *   viewer.next(); viewer.goTo(3);
 *   // when done: viewer.destroy(); deck.dispose();
 */
export { Deck, Deck as loadDeck } from './pptx/deck/deck.js';
export { Viewer, createViewer, type ViewerOptions } from './pptx/viewer/viewer.js';
export { renderSlide, type RenderDeps } from './pptx/render/dom.js';
export { installDeckFonts, type FontInstallation } from './pptx/render/fonts.js';

import { Deck } from './pptx/deck/deck.js';
/** Load a .pptx from raw bytes into a renderable {@link Deck}. */
export function loadPptx(data: ArrayBuffer | Uint8Array): Deck {
  return Deck.load(data);
}

// Lower-level building blocks for advanced users / tooling.
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
