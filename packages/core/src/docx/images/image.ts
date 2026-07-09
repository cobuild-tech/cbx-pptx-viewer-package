/**
 * Inline image parsing: <w:drawing> (DrawingML wp:inline / wp:anchor) ->
 * DocxInlineImage. Extent is in EMU; the media part is resolved from the blip's
 * r:embed relationship on the document part.
 */
import { child, attr, attrNum, path, localName, type XmlNode } from '../../oxml/xml.js';
import { emuToPx } from '../../oxml/units.js';
import { logicalChildrenNamed } from '../content.js';
import type { DocxInlineImage } from '../model.js';
import type { ParseContext } from '../document/context.js';

/** Extract all inline images contained in a paragraph's runs, in order. */
export function findImages(p: XmlNode, ctx: ParseContext): DocxInlineImage[] {
  const out: DocxInlineImage[] = [];
  for (const r of logicalChildrenNamed(p, 'r')) {
    // <w:drawing> (DrawingML) and legacy <w:pict>/<v:shape> (VML) image runs.
    for (const node of r.children) {
      const name = localName(node.name);
      if (name === 'drawing') {
        const img = fromDrawing(node, ctx);
        if (img) out.push(img);
      } else if (name === 'pict') {
        const img = fromPict(node, ctx);
        if (img) out.push(img);
      }
    }
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

/** Legacy VML image: <w:pict><v:shape style="width:..pt;height:..pt"><v:imagedata r:id=".."/>. */
function fromPict(pict: XmlNode, ctx: ParseContext): DocxInlineImage | undefined {
  const shape = findFirst(pict, (n) => localName(n.name) === 'imagedata');
  if (!shape) return undefined;
  const rel = ctx.rel(attr(shape, 'id'));
  if (!rel) return undefined;

  // Size comes from the parent shape's CSS-ish `style` (in points).
  const styled = findFirst(pict, (n) => attr(n, 'style') !== undefined);
  const style = attr(styled, 'style') ?? '';
  const w = ptStyle(style, 'width');
  const h = ptStyle(style, 'height');
  const alt = attr(shape, 'title') ?? attr(shape, 'alt');

  return {
    kind: 'image',
    part: rel.target,
    widthPx: w ?? 0,
    heightPx: h ?? 0,
    ...(alt ? { alt } : {}),
  };
}

/** Parse `width:123pt` (or px) out of a VML style string, returning px. */
function ptStyle(style: string, prop: string): number | undefined {
  const m = new RegExp(`${prop}\\s*:\\s*([0-9.]+)(pt|px)?`, 'i').exec(style);
  if (!m) return undefined;
  const val = parseFloat(m[1]!);
  return m[2]?.toLowerCase() === 'px' ? val : val * (96 / 72);
}

function findFirst(node: XmlNode, pred: (n: XmlNode) => boolean): XmlNode | undefined {
  for (const c of node.children) {
    if (pred(c)) return c;
    const nested = findFirst(c, pred);
    if (nested) return nested;
  }
  return undefined;
}
