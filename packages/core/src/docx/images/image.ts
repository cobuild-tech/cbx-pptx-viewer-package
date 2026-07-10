/**
 * Drawing parsing: <w:drawing> (DrawingML) and legacy <w:pict> (VML) →
 * inline {@link DocxDrawing} (in the run stream) or a floating
 * {@link DocxAnchor} (attached to its containing paragraph).
 *
 * A drawing's payload is either a picture (<pic:pic>, a raster/vector image
 * resolved from the blip's r:embed) or a DrawingML shape (<wps:wsp>, a
 * filled/outlined box that may hold a text box). Extents are in EMU.
 */
import { child, attr, attrNum, attrBool, path, localName, type XmlNode } from '../../oxml/xml.js';
import { emuToPx } from '../../oxml/units.js';
import type {
  DocxDrawing,
  DocxInlineImage,
  DocxShape,
  DocxAnchor,
  DocxWrap,
  DocxGeom,
  DocxCrop,
} from '../model.js';
import type { ParseContext } from '../document/context.js';

/**
 * Scale an inline image down so it fits the available width, preserving aspect
 * ratio (Word shrinks an inline image to the text column). Smaller images are
 * left untouched.
 */
export function fitImageWidth(img: DocxInlineImage, contentW: number): void {
  if (img.widthPx > contentW && img.widthPx > 0) {
    const scale = contentW / img.widthPx;
    img.widthPx = contentW;
    img.heightPx = img.heightPx * scale;
  }
}

/** Result of parsing one drawing-bearing run child. */
export interface DrawingResult {
  /** An inline drawing to place in the run stream. */
  inline?: DocxDrawing;
  /** A floating drawing to attach to the containing paragraph. */
  anchor?: DocxAnchor;
}

/**
 * Parse a run child that carries a drawing: <w:drawing> (DrawingML),
 * <w:pict> (VML), or <mc:AlternateContent> (prefer the DrawingML Choice, fall
 * back to the VML Fallback). Returns whichever of inline/anchor applies.
 */
export function parseRunDrawing(node: XmlNode, ctx: ParseContext): DrawingResult | undefined {
  const name = localName(node.name);
  if (name === 'drawing') return fromDrawing(node, ctx);
  if (name === 'pict') return fromPict(node, ctx);
  if (name === 'AlternateContent') {
    const choice = child(node, 'Choice');
    for (const c of choice?.children ?? []) {
      const r = parseRunDrawing(c, ctx);
      if (r) return r;
    }
    const fallback = child(node, 'Fallback');
    for (const c of fallback?.children ?? []) {
      const r = parseRunDrawing(c, ctx);
      if (r) return r;
    }
  }
  return undefined;
}

function fromDrawing(drawing: XmlNode, ctx: ParseContext): DrawingResult | undefined {
  const inline = child(drawing, 'inline');
  if (inline) {
    const d = payload(inline, ctx);
    return d ? { inline: d } : undefined;
  }
  const anchor = child(drawing, 'anchor');
  if (anchor) {
    const a = fromAnchor(anchor, ctx);
    return a ? { anchor: a } : undefined;
  }
  return undefined;
}

/** Build the drawing payload (picture or shape) from a wp:inline / wp:anchor. */
function payload(container: XmlNode, ctx: ParseContext): DocxDrawing | undefined {
  const ext = child(container, 'extent');
  const wPx = emuToPx(attrNum(ext, 'cx') ?? 0);
  const hPx = emuToPx(attrNum(ext, 'cy') ?? 0);

  const docPr = child(container, 'docPr');
  const alt = attr(docPr, 'descr') ?? attr(docPr, 'name');

  const gd = path(container, 'graphic/graphicData');
  // DrawingML shape (text box / banner box).
  const wsp = gd && child(gd, 'wsp');
  if (wsp) return shapeFrom(wsp, wPx, hPx, alt, ctx);

  // Picture.
  const rel = ctx.rel(blipEmbed(container));
  if (rel) {
    const img: DocxInlineImage = { kind: 'image', part: rel.target, widthPx: wPx, heightPx: hPx };
    const crop = cropOf(path(container, 'graphic/graphicData/pic/blipFill'));
    if (crop) img.crop = crop;
    if (alt) img.alt = alt;
    return img;
  }
  return undefined;
}

/**
 * Parse an <a:srcRect> crop (child of a blipFill) into fractional insets. Its
 * l/t/r/b are in thousandths of a percent (60000 = 60%). Returns undefined when
 * absent or all-zero (no crop).
 */
function cropOf(blipFill: XmlNode | undefined): DocxCrop | undefined {
  const sr = child(blipFill, 'srcRect');
  if (!sr) return undefined;
  const f = (name: string) => (attrNum(sr, name) ?? 0) / 100000;
  const l = f('l');
  const t = f('t');
  const r = f('r');
  const b = f('b');
  if (!l && !t && !r && !b) return undefined;
  return { l, t, r, b };
}

const GEOM_MAP: Record<string, DocxGeom> = {
  rect: 'rect',
  roundRect: 'roundRect',
  ellipse: 'ellipse',
  line: 'line',
  straightConnector1: 'line',
};

/** Build a shape payload from a <wps:wsp>. */
function shapeFrom(
  wsp: XmlNode,
  wPx: number,
  hPx: number,
  alt: string | undefined,
  ctx: ParseContext,
): DocxShape {
  const spPr = child(wsp, 'spPr');
  const prst = attr(child(spPr, 'prstGeom'), 'prst');
  const geom: DocxGeom = prst ? (GEOM_MAP[prst] ?? 'other') : 'rect';

  const fillHex = solidHex(child(spPr, 'solidFill'));
  const fillImage = pictureFill(child(spPr, 'blipFill'), wPx, hPx, ctx);

  const ln = child(spPr, 'ln');
  const lineHex = solidHex(child(ln, 'solidFill'));
  const lnW = attrNum(ln, 'w');

  const txbx = child(wsp, 'txbx');
  const content = txbx ? ctx.parseBlocks(child(txbx, 'txbxContent')) : [];

  const bodyPr = child(wsp, 'bodyPr');
  const anchorAttr = attr(bodyPr, 'anchor');
  const vAnchor = anchorAttr === 't' ? 'top' : anchorAttr === 'b' ? 'bottom' : anchorAttr === 'ctr' ? 'ctr' : undefined;

  const shape: DocxShape = { kind: 'shape', geom, widthPx: wPx, heightPx: hPx, content };
  if (fillHex) shape.fillHex = fillHex;
  if (fillImage) shape.fillImage = fillImage;
  if (lineHex) shape.lineHex = lineHex;
  if (lnW !== undefined) shape.lineWidthPx = emuToPx(lnW);
  if (vAnchor) shape.vAnchor = vAnchor;
  if (alt) shape.alt = alt;
  return shape;
}

/**
 * Resolve an <a:blipFill> used as a shape's fill (a picture-filled banner/box)
 * to an image sized to the shape's extent, honoring any <a:srcRect> crop.
 */
function pictureFill(
  blipFill: XmlNode | undefined,
  wPx: number,
  hPx: number,
  ctx: ParseContext,
): DocxInlineImage | undefined {
  if (!blipFill) return undefined;
  const blip = child(blipFill, 'blip');
  const rel = ctx.rel(attr(blip, 'embed') ?? attr(blip, 'link'));
  if (!rel) return undefined;
  const img: DocxInlineImage = { kind: 'image', part: rel.target, widthPx: wPx, heightPx: hPx };
  const crop = cropOf(blipFill);
  if (crop) img.crop = crop;
  return img;
}

/** Resolve an <a:solidFill> to a hex string (srgbClr only; scheme colors skipped). */
function solidHex(fill: XmlNode | undefined): string | undefined {
  if (!fill) return undefined;
  const srgb = child(fill, 'srgbClr');
  const val = attr(srgb, 'val');
  return val ? val.replace(/^#/, '') : undefined;
}

function fromAnchor(anchor: XmlNode, ctx: ParseContext): DocxAnchor | undefined {
  const drawing = payload(anchor, ctx);
  if (!drawing) return undefined;

  const ext = child(anchor, 'extent');
  const wPx = emuToPx(attrNum(ext, 'cx') ?? 0);
  const hPx = emuToPx(attrNum(ext, 'cy') ?? 0);

  const posH = child(anchor, 'positionH');
  const posV = child(anchor, 'positionV');

  const a: DocxAnchor = {
    drawing,
    wPx,
    hPx,
    behindDoc: attrBool(anchor, 'behindDoc', false),
    wrap: wrapMode(anchor),
    relH: attr(posH, 'relativeFrom') ?? 'column',
    relV: attr(posV, 'relativeFrom') ?? 'paragraph',
  };

  const zOrder = attrNum(anchor, 'relativeHeight');
  if (zOrder !== undefined) a.zOrder = zOrder;

  const hOff = numText(child(posH, 'posOffset'));
  if (hOff !== undefined) a.hOffsetPx = emuToPx(hOff);
  const hAlign = child(posH, 'align')?.text?.trim();
  if (hAlign) a.hAlign = hAlign as DocxAnchor['hAlign'];

  const vOff = numText(child(posV, 'posOffset'));
  if (vOff !== undefined) a.vOffsetPx = emuToPx(vOff);
  const vAlign = child(posV, 'align')?.text?.trim();
  if (vAlign) a.vAlign = vAlign as DocxAnchor['vAlign'];

  return a;
}

function wrapMode(anchor: XmlNode): DocxWrap {
  for (const c of anchor.children) {
    switch (localName(c.name)) {
      case 'wrapNone':
        return 'none';
      case 'wrapSquare':
        return 'square';
      case 'wrapTight':
        return 'tight';
      case 'wrapThrough':
        return 'through';
      case 'wrapTopAndBottom':
        return 'topAndBottom';
    }
  }
  return 'none';
}

function blipEmbed(container: XmlNode): string | undefined {
  const blip = path(container, 'graphic/graphicData/pic/blipFill/blip');
  return attr(blip, 'embed') ?? attr(blip, 'link');
}

function numText(node: XmlNode | undefined): number | undefined {
  if (!node) return undefined;
  const n = Number(node.text);
  return Number.isFinite(n) ? n : undefined;
}

// ─── Legacy VML (<w:pict>) ───────────────────────────────────────────────────

/** Legacy VML image: <w:pict><v:shape style="width:..pt;height:..pt"><v:imagedata r:id=".."/>. */
function fromPict(pict: XmlNode, ctx: ParseContext): DrawingResult | undefined {
  const imagedata = findFirst(pict, (n) => localName(n.name) === 'imagedata');
  if (!imagedata) return undefined;
  const rel = ctx.rel(attr(imagedata, 'id'));
  if (!rel) return undefined;

  // Size comes from the parent shape's CSS-ish `style` (in points).
  const styled = findFirst(pict, (n) => attr(n, 'style') !== undefined);
  const style = attr(styled, 'style') ?? '';
  const w = ptStyle(style, 'width');
  const h = ptStyle(style, 'height');
  const alt = attr(imagedata, 'title') ?? attr(imagedata, 'alt');

  const img: DocxInlineImage = {
    kind: 'image',
    part: rel.target,
    widthPx: w ?? 0,
    heightPx: h ?? 0,
  };
  if (alt) img.alt = alt;
  return { inline: img };
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
