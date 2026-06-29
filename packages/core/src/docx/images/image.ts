/**
 * Inline image parser for WordprocessingML.
 *
 * Inline images live inside <w:drawing><wp:inline>...</wp:inline></w:drawing>.
 * Floating images (<wp:anchor>) are not yet supported.
 */
import { child, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { emuToPx } from '../../oxml/units.js';
import type { DocxInlineImage } from '../model.js';

/**
 * Parse a <w:drawing> element into a DocxInlineImage.
 * Returns null if it is not an inline image or cannot be resolved.
 */
export function parseDrawing(
  drawingEl: XmlNode,
  resolveImage: (relId: string) => string | undefined,
): DocxInlineImage | null {
  // Support both inline (<wp:inline>) and floating/anchored (<wp:anchor>) images.
  const inline = child(drawingEl, 'inline') ?? child(drawingEl, 'anchor');
  if (!inline) return null;

  const extent = child(inline, 'extent');
  const cx = attrNum(extent, 'cx') ?? 0;
  const cy = attrNum(extent, 'cy') ?? 0;

  // Navigate: <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="..."/>
  const graphic = child(inline, 'graphic');
  const graphicData = child(graphic, 'graphicData');
  if (!graphicData) return null;

  // pic:pic may be nested directly or inside another wrapper.
  const picPic = findDescendantByLocalName(graphicData, 'pic');
  if (!picPic) return null;

  const blipFill = child(picPic, 'blipFill');
  const blip = child(blipFill, 'blip');
  if (!blip) return null;

  const relId = attr(blip, 'r:embed') ?? attr(blip, 'embed');
  if (!relId) return null;

  const part = resolveImage(relId);
  if (!part) return null;

  const docDescr = child(inline, 'docPr');
  const alt = attr(docDescr, 'descr') ?? attr(docDescr, 'title');

  return {
    kind: 'image',
    part,
    widthPx: emuToPx(cx),
    heightPx: emuToPx(cy),
    alt: alt ?? undefined,
  };
}

function findDescendantByLocalName(node: XmlNode, name: string): XmlNode | undefined {
  for (const c of node.children) {
    if (localName(c.name) === name) return c;
    const found = findDescendantByLocalName(c, name);
    if (found) return found;
  }
  return undefined;
}
