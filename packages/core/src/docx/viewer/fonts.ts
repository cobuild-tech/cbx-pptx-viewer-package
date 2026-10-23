/**
 * The font families a document's text is drawn in — what the viewer needs to
 * have available (or stood in for) before pages are measured.
 */
import type { DocxBlock, DocxDrawing, DocxSection } from '../model.js';

/** Every distinct font family named anywhere in the sections, headers and footers included. */
export function collectDocxFonts(sections: DocxSection[]): Set<string> {
  const out = new Set<string>();
  const visitDrawing = (d: DocxDrawing | undefined): void => {
    if (d?.kind === 'shape') visit(d.content);
    else if (d?.kind === 'group') d.items.forEach((i) => visitDrawing(i.drawing));
  };
  const visit = (blocks: DocxBlock[] | undefined): void => {
    for (const b of blocks ?? []) {
      if (b.kind === 'paragraph') {
        if (b.baseFontFamily) out.add(b.baseFontFamily);
        if (b.listMarkerFormat?.font) out.add(b.listMarkerFormat.font);
        for (const r of b.runs) {
          if (r.font) out.add(r.font);
          visitDrawing(r.drawing);
        }
        for (const a of b.anchors ?? []) visitDrawing(a.drawing);
      } else if (b.kind === 'table') {
        for (const row of b.rows) for (const cell of row) if (cell) visit(cell.content);
      }
    }
  };
  for (const s of sections) {
    visit(s.blocks);
    visit(s.header);
    visit(s.footer);
    visit(s.firstHeader);
    visit(s.firstFooter);
    visit(s.evenHeader);
    visit(s.evenFooter);
  }
  return out;
}
