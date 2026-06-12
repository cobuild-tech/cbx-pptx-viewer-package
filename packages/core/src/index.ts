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
export { Deck, Deck as loadDeck } from './parse/deck.js';
export { Viewer, createViewer, type ViewerOptions } from './viewer/viewer.js';
export { renderSlide, type RenderDeps } from './render/dom.js';
export { installDeckFonts, type FontInstallation } from './render/fonts.js';

import { Deck } from './parse/deck.js';
/** Load a .pptx from raw bytes into a renderable {@link Deck}. */
export function loadPptx(data: ArrayBuffer | Uint8Array): Deck {
  return Deck.load(data);
}

// Lower-level building blocks for advanced users / tooling.
export { OpcPackage, type Relationship } from './opc/package.js';
export { RelType } from './opc/relTypes.js';
export * as units from './units.js';
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
} from './xml.js';
export * from './model.js';
