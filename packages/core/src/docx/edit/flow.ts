/**
 * Continuous-flow renderer for edit mode.
 *
 * The paginator splits a paragraph mid-text when it straddles a page boundary,
 * producing two cloned paragraphs on two different pages. A contentEditable
 * region cannot span that, and the clones have no source identity — so editing
 * the paginated view would be fighting the layout the whole way.
 *
 * Editing therefore happens in one continuous column, like Word's Web Layout.
 * This costs nothing in fidelity: **pagination is not stored in a .docx** —
 * Word reflows on open — so our pages are purely a display concern and the
 * saved file is identical either way. Explicit page breaks still render, as a
 * visible rule rather than a real page boundary.
 */
import { renderBlock, SHEET_FONT, type RenderDeps } from '../render/dom.js';
import type { DocxSection } from '../model.js';

/** Class on the flow container, so the viewer can find and restyle it. */
export const FLOW_CLASS = 'cbx-docx-flow';

/**
 * Render every section as one scrollable column at the section's content width
 * (page width minus margins), so line breaking still matches the real page.
 */
export function renderFlow(sections: DocxSection[], deps: RenderDeps): HTMLElement {
  const root = document.createElement('div');
  root.className = FLOW_CLASS;
  root.style.display = 'flex';
  root.style.flexDirection = 'column';
  root.style.alignItems = 'center';
  root.style.width = '100%';

  for (const section of sections) {
    const contentW = section.size.wPx - section.margins.leftPx - section.margins.rightPx;

    const sheet = document.createElement('div');
    sheet.style.boxSizing = 'content-box';
    sheet.style.width = `${contentW}px`;
    sheet.style.padding = `${section.margins.topPx}px ${section.margins.rightPx}px ${section.margins.bottomPx}px ${section.margins.leftPx}px`;
    sheet.style.background = '#fff';
    sheet.style.color = '#000';
    sheet.style.fontFamily = SHEET_FONT.fontFamily;
    sheet.style.fontSize = SHEET_FONT.fontSize;
    sheet.style.lineHeight = SHEET_FONT.lineHeight;
    // Headers and footers belong to the page furniture, which continuous flow
    // has none of; they stay read-only in the paginated view.
    for (const block of section.blocks) {
      sheet.appendChild(renderBlock(block, deps));
    }
    root.appendChild(sheet);
  }

  return root;
}
