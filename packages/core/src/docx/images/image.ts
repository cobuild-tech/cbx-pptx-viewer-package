/**
 * Inline image parsing: <w:drawing> (DrawingML wp:inline / wp:anchor) ->
 * DocxInlineImage. Extent is in EMU; the media part is resolved from the blip's
 * r:embed relationship on the document part.
 */
import { child, children, attr, attrNum, path, type XmlNode } from '../../oxml/xml.js';
import { emuToPx } from '../../oxml/units.js';
import type { DocxInlineImage } from '../model.js';
import type { ParseContext } from '../document/context.js';

/** Extract all inline images contained in a paragraph's runs, in order. */
export function findImages(p: XmlNode, ctx: ParseContext): DocxInlineImage[] {
  const out: DocxInlineImage[] = [];
  for (const r of children(p, 'r')) {
    for (const drawing of children(r, 'drawing')) {
      const img = fromDrawing(drawing, ctx);
      if (img) out.push(img);
    }
    // Legacy VML fallback (<w:pict>) is not handled in v1.
  }
  return out;
}

function fromDrawing(drawing: XmlNode, ctx: ParseContext): DocxInlineImage | undefined {
  const anchor = child(drawing, 'inline') ?? child(drawing, 'anchor');
  if (!anchor) return undefined;

  const ext = child(anchor, 'extent');
  const cx = attrNum(ext, 'cx');
  const cy = attrNum(ext, 'cy');

  // graphic -> graphicData -> pic:pic -> blipFill -> blip@r:embed
  const blip = path(anchor, 'graphic/graphicData/pic/blipFill/blip');
  const embed = attr(blip, 'embed') ?? attr(blip, 'link');
  const rel = ctx.rel(embed);
  if (!rel || !embed) return undefined;

  const docPr = child(anchor, 'docPr');
  const alt = attr(docPr, 'descr') ?? attr(docPr, 'name');

  return {
    kind: 'image',
    part: rel.target,
    widthPx: cx ? emuToPx(cx) : 0,
    heightPx: cy ? emuToPx(cy) : 0,
    ...(alt ? { alt } : {}),
  };
}
