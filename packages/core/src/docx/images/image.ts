/**
 * Inline image parsing: <w:drawing> (DrawingML wp:inline / wp:anchor) ->
 * DocxInlineImage. Extent is in EMU; the media part is resolved from the blip's
 * r:embed relationship on the document part.
 */
import { child, attr, attrNum, attrBool, path, localName, type XmlNode } from '../../oxml/xml.js';
import { emuToPx } from '../../oxml/units.js';
import { logicalChildrenNamed } from '../content.js';
import type { DocxInlineImage, DocxFloat, DocxPageSize, DocxPageMargins } from '../model.js';
import type { ParseContext } from '../document/context.js';

/** Extract inline images from a paragraph's runs, in order. */
export function findImages(p: XmlNode, ctx: ParseContext): DocxInlineImage[] {
  const out: DocxInlineImage[] = [];
  for (const r of logicalChildrenNamed(p, 'r')) {
    // <w:drawing> (DrawingML) and legacy <w:pict>/<v:shape> (VML) image runs.
    for (const node of r.children) {
      const name = localName(node.name);
      if (name === 'drawing') {
        // Anchored (floating) images are hoisted to page floats in header/footer
        // parsing; skip them here so they don't also render inline.
        if (ctx.hoistAnchors && child(node, 'anchor')) continue;
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

/**
 * Collect anchored (floating) images anywhere under `root`, resolved to
 * absolute page coordinates using the section's size/margins.
 */
export function collectFloats(
  root: XmlNode,
  ctx: ParseContext,
  size: DocxPageSize,
  margins: DocxPageMargins,
): DocxFloat[] {
  const out: DocxFloat[] = [];
  const tag = localName(root.name);
  const context = tag === 'hdr' ? 'header' : tag === 'ftr' ? 'footer' : 'body';

  const walk = (n: XmlNode) => {
    if (localName(n.name) === 'drawing') {
      const anchor = child(n, 'anchor');
      if (anchor) {
        const f = fromAnchor(anchor, ctx, size, margins, context);
        if (f) out.push(f);
      }
    }
    for (const c of n.children) walk(c);
  };
  walk(root);
  return out;
}

function fromDrawing(drawing: XmlNode, ctx: ParseContext): DocxInlineImage | undefined {
  const anchor = child(drawing, 'inline') ?? child(drawing, 'anchor');
  if (!anchor) return undefined;

  const ext = child(anchor, 'extent');
  const cx = attrNum(ext, 'cx');
  const cy = attrNum(ext, 'cy');

  const rel = ctx.rel(blipEmbed(anchor));
  if (!rel) return undefined;

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

function fromAnchor(
  anchor: XmlNode,
  ctx: ParseContext,
  size: DocxPageSize,
  margins: DocxPageMargins,
  context: 'header' | 'footer' | 'body',
): DocxFloat | undefined {
  const rel = ctx.rel(blipEmbed(anchor));
  if (!rel) return undefined;

  const ext = child(anchor, 'extent');
  const wPx = emuToPx(attrNum(ext, 'cx') ?? 0);
  const hPx = emuToPx(attrNum(ext, 'cy') ?? 0);

  const xPx = resolvePos(child(anchor, 'positionH'), 'h', wPx, size, margins, context);
  const yPx = resolvePos(child(anchor, 'positionV'), 'v', hPx, size, margins, context);

  const docPr = child(anchor, 'docPr');
  const alt = attr(docPr, 'descr') ?? attr(docPr, 'name');

  return {
    part: rel.target,
    xPx,
    yPx,
    wPx,
    hPx,
    behindDoc: attrBool(anchor, 'behindDoc', false),
    ...(alt ? { alt } : {}),
  };
}

/** Resolve <wp:positionH>/<wp:positionV> to a page-coordinate offset in px. */
function resolvePos(
  pos: XmlNode | undefined,
  axis: 'h' | 'v',
  sizePx: number,
  size: DocxPageSize,
  margins: DocxPageMargins,
  context: 'header' | 'footer' | 'body',
): number {
  const relFrom = attr(pos, 'relativeFrom') ?? (axis === 'h' ? 'column' : 'paragraph');
  const pageExtent = axis === 'h' ? size.wPx : size.hPx;
  const startMargin = axis === 'h' ? margins.leftPx : margins.topPx;
  const endMargin = axis === 'h' ? margins.rightPx : margins.bottomPx;
  
  // Everything except 'page' is relative to the margin/text area start.
  let base = relFrom === 'page' ? 0 : startMargin;

  if (axis === 'v' && relFrom !== 'page') {
    if (context === 'header') {
      base = margins.headerPx;
    } else if (context === 'footer') {
      base = size.hPx - margins.footerPx - sizePx;
    }
  }

  const offset = numText(child(pos, 'posOffset'));
  if (offset !== undefined) return base + emuToPx(offset);

  const align = child(pos, 'align')?.text?.trim();
  switch (align) {
    case 'center':
      return (pageExtent - sizePx) / 2;
    case 'right':
    case 'bottom':
      return pageExtent - endMargin - sizePx;
    case 'left':
    case 'top':
      return base;
    default:
      return base;
  }
}

function blipEmbed(anchor: XmlNode): string | undefined {
  const blip = path(anchor, 'graphic/graphicData/pic/blipFill/blip');
  return attr(blip, 'embed') ?? attr(blip, 'link');
}

function numText(node: XmlNode | undefined): number | undefined {
  if (!node) return undefined;
  const n = Number(node.text);
  return Number.isFinite(n) ? n : undefined;
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
