/**
 * Web-font resolution for decks. The fetching itself (Google Fonts, the
 * metric-compatible clone map, availability probing) is shared with the other
 * formats and lives in oxml/webfonts.ts; this module only knows which families
 * a deck uses.
 */
import type { Deck } from '../deck/deck.js';
import type { Shape, TextBody } from '../model.js';
import { splitFontWeight } from '../text/render.js';
import { installWebFontFamilies, type FontInstallation, type WebFontOptions } from '../../oxml/webfonts.js';

export type { WebFontOptions } from '../../oxml/webfonts.js';

/**
 * Fetch and register every font the deck uses but that isn't already available,
 * resolving once the needed faces have loaded so the viewer can re-render with
 * correct metrics.
 */
export function installWebFonts(deck: Deck, opts: WebFontOptions = {}): FontInstallation {
  // Embedded faces are installed by installDeckFonts, under their own names.
  const embedded = new Set(deck.embeddedFonts.map((f) => f.typeface.toLowerCase()));
  return installWebFontFamilies(collectFontFamilies(deck), opts, embedded);
}

/** Every distinct base font family referenced by text anywhere in the deck. */
export function collectFontFamilies(deck: Deck): Set<string> {
  const out = new Set<string>();
  const add = (font: string | undefined) => {
    if (!font) return;
    const family = splitFontWeight(font).family.trim();
    if (family) out.add(family);
  };
  const visitText = (body: TextBody | undefined) => {
    if (!body) return;
    for (const para of body.paragraphs) {
      if (para.bullet && 'font' in para.bullet) add(para.bullet.font);
      for (const run of para.runs) add(run.font);
    }
  };
  const visitShape = (shape: Shape) => {
    switch (shape.kind) {
      case 'shape':
        visitText(shape.text);
        break;
      case 'group':
        shape.children.forEach(visitShape);
        break;
      case 'frame':
        shape.diagram?.forEach(visitShape);
        if (shape.table) {
          for (const row of shape.table.rows) {
            for (const cell of row) if (cell) visitText(cell.text);
          }
        }
        break;
    }
  };
  for (const slide of deck.slides) slide.shapes.forEach(visitShape);
  return out;
}
