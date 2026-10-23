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
} from '../model.js';
import type { ParseContext } from '../document/context.js';
import { resolveContainerColor, resolveColorEl, type ColorContext } from '../../oxml/drawingml/color.js';
import { hasPreset, parseAdjust } from '../../oxml/drawingml/geometry/presets.js';
import { parseCustomGeometry } from '../../oxml/drawingml/geometry/custom.js';
import { dmlTextBlocks, graphicFrame } from './dml.js';

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
  // An embedded OLE object shows its VML preview picture.
  if (name === 'pict' || name === 'object') return fromPict(node, ctx);
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

  // Room Word leaves around the drawing for its outline, shadow or glow — part
  // of the layout box of an inline drawing, though not of its extent.
  const ee = child(container, 'effectExtent');
  const effect = {
    top: emuToPx(Math.max(0, attrNum(ee, 't') ?? 0)),
    right: emuToPx(Math.max(0, attrNum(ee, 'r') ?? 0)),
    bottom: emuToPx(Math.max(0, attrNum(ee, 'b') ?? 0)),
    left: emuToPx(Math.max(0, attrNum(ee, 'l') ?? 0)),
  };
  const hasEffect = effect.top || effect.right || effect.bottom || effect.left;

  const gd = path(container, 'graphic/graphicData');
  // DrawingML shape (text box / banner box).
  const wsp = gd && child(gd, 'wsp');
  if (wsp) {
    const shape = shapeFrom(wsp, wPx, hPx, alt, ctx);
    if (hasEffect) shape.effectPx = effect;
    return shape;
  }
  // A chart, a shape group, or a SmartArt diagram.
  const other = gd ? graphicFrame(gd, wPx, hPx, ctx) : undefined;
  if (other) {
    if (hasEffect) other.effectPx = effect;
    return other;
  }

  // Picture.
  const rel = ctx.rel(blipEmbed(container));
  if (rel) {
    const img: DocxInlineImage = { kind: 'image', part: rel.target, widthPx: wPx, heightPx: hPx };
    if (alt) img.alt = alt;
    if (hasEffect) img.effectPx = effect;
    const duotone = child(path(container, 'graphic/graphicData/pic/blipFill/blip'), 'duotone');
    if (duotone) {
      // Its two colour children, in order: darks, then lights.
      const [dark, light] = duotone.children
        .map((c) => resolveColorEl(c, ctx.colors)?.hex)
        .filter((hex): hex is string => hex !== undefined);
      if (dark && light) img.duotone = [dark, light];
    }
    return img;
  }
  return undefined;
}

const GEOM_MAP: Record<string, DocxGeom> = {
  rect: 'rect',
  roundRect: 'roundRect',
  ellipse: 'ellipse',
  line: 'line',
  straightConnector1: 'line',
};

/**
 * Build a shape payload from a DrawingML shape: a Word shape (<wps:wsp>, whose
 * text is a Word text box) or a plain DrawingML one (<dsp:sp> in a SmartArt
 * drawing, a group member), whose text is a DrawingML text body.
 */
export function shapeFrom(
  wsp: XmlNode,
  wPx: number,
  hPx: number,
  alt: string | undefined,
  ctx: ParseContext,
): DocxShape {
  const spPr = child(wsp, 'spPr');
  const prstGeom = child(spPr, 'prstGeom');
  const prst = attr(prstGeom, 'prst');
  const custGeom = child(spPr, 'custGeom');
  const geom: DocxGeom = prst ? (GEOM_MAP[prst] ?? 'other') : custGeom ? 'other' : 'rect';

  // Fill and outline: stated in spPr, or else taken from the shape style's
  // theme references (<wps:style> fillRef/lnRef), which is how a shape drawn
  // in Word gets its default blue fill and darker outline.
  const style = child(wsp, 'style');
  const colors = ctx.colors;
  const fillHex = fillColor(spPr, child(style, 'fillRef'), colors);

  const ln = child(spPr, 'ln');
  const lnRef = child(style, 'lnRef');
  const lineHex = child(ln, 'noFill')
    ? undefined
    : (resolveContainerColor(child(ln, 'solidFill'), colors)?.hex ??
      (child(ln, 'solidFill') || (attrNum(lnRef, 'idx') ?? 0) === 0 ? undefined : resolveContainerColor(lnRef, colors)?.hex));
  const lnW = attrNum(ln, 'w') ?? (lineHex && attrNum(lnRef, 'idx') ? 12700 : undefined);
  const fontRef = child(style, 'fontRef');

  const txbx = child(wsp, 'txbx');
  const txBody = child(wsp, 'txBody');
  const content = txbx
    ? ctx.parseBlocks(child(txbx, 'txbxContent'))
    : txBody
      ? dmlTextBlocks(txBody, ctx, resolveContainerColor(child(child(wsp, 'style'), 'fontRef'), ctx.colors)?.hex)
      : [];

  const bodyPr = child(wsp, 'bodyPr') ?? child(txBody, 'bodyPr');
  const anchorAttr = attr(bodyPr, 'anchor');
  const vert = attr(bodyPr, 'vert');
  const vAnchor = anchorAttr === 't' ? 'top' : anchorAttr === 'b' ? 'bottom' : anchorAttr === 'ctr' ? 'ctr' : undefined;

  const shape: DocxShape = { kind: 'shape', geom, widthPx: wPx, heightPx: hPx, content };
  // Insets default to 0.1" left/right and 0.05" top/bottom.
  const inset = (name: string, dflt: number) => emuToPx(attrNum(bodyPr, name) ?? dflt);
  shape.insetsPx = {
    top: inset('tIns', 45720),
    right: inset('rIns', 91440),
    bottom: inset('bIns', 45720),
    left: inset('lIns', 91440),
  };
  // A SmartArt shape's text sits in its own rectangle (<dsp:txXfrm>), stated
  // in the same space as the shape's: fold the difference into the insets.
  const txXfrm = child(wsp, 'txXfrm');
  const shapeXfrm = child(spPr, 'xfrm');
  if (txXfrm && shapeXfrm) {
    const so = child(shapeXfrm, 'off');
    const se = child(shapeXfrm, 'ext');
    const to = child(txXfrm, 'off');
    const te = child(txXfrm, 'ext');
    const sx = attrNum(so, 'x') ?? 0;
    const sy = attrNum(so, 'y') ?? 0;
    const tx = attrNum(to, 'x') ?? sx;
    const ty = attrNum(to, 'y') ?? sy;
    const i = shape.insetsPx;
    i.left += emuToPx(tx - sx);
    i.top += emuToPx(ty - sy);
    i.right += emuToPx(sx + (attrNum(se, 'cx') ?? 0) - tx - (attrNum(te, 'cx') ?? 0));
    i.bottom += emuToPx(sy + (attrNum(se, 'cy') ?? 0) - ty - (attrNum(te, 'cy') ?? 0));
  }
  // Geometry beyond what CSS can draw goes to an SVG outline.
  if (prst && prst !== 'rect' && hasPreset(prst)) {
    shape.preset = prst;
    const adjust = parseAdjust(child(prstGeom, 'avLst'));
    if (Object.keys(adjust).length) shape.adjust = adjust;
  } else if (custGeom) {
    const paths = parseCustomGeometry(custGeom);
    if (paths.length) shape.paths = paths;
  }
  const rot = attrNum(shapeXfrm, 'rot');
  if (rot) shape.rot = rot / 60000;
  if (fillHex) shape.fillHex = fillHex;
  // A picture fill: the shape is a frame for an image (often with text on it).
  const blipFill = child(spPr, 'blipFill');
  const blip = child(blipFill, 'blip');
  const fillRel = ctx.rel(attr(blip, 'embed') ?? attr(blip, 'link'));
  if (fillRel) shape.fillImagePart = fillRel.target;
  if (lineHex) shape.lineHex = lineHex;
  const textHex = resolveContainerColor(fontRef, colors)?.hex;
  if (textHex) shape.textColorHex = textHex;
  if (lnW !== undefined) shape.lineWidthPx = emuToPx(lnW);
  if (vAnchor) shape.vAnchor = vAnchor;
  // vert/eaVert run text top-to-bottom; vert270 bottom-to-top. (mongolianVert
  // and wordArtVert are approximated as eaVert.)
  if (vert === 'vert' || vert === 'vert270') shape.vert = vert;
  else if (vert === 'eaVert' || vert === 'mongolianVert' || vert === 'wordArtVert' || vert === 'wordArtVertRtl') shape.vert = 'eaVert';
  if (alt) shape.alt = alt;
  return shape;
}

/**
 * A shape's fill colour: an explicit solid fill, the first stop of a gradient
 * (an approximation — CSS could draw it, but a flat fill keeps text legible),
 * nothing for noFill, else the style's fillRef colour.
 */
function fillColor(spPr: XmlNode | undefined, fillRef: XmlNode | undefined, colors: ColorContext): string | undefined {
  if (child(spPr, 'noFill')) return undefined;
  const solid = child(spPr, 'solidFill');
  if (solid) return resolveContainerColor(solid, colors)?.hex;
  const grad = child(spPr, 'gradFill');
  if (grad) return resolveContainerColor(child(child(grad, 'gsLst'), 'gs'), colors)?.hex;
  if (child(spPr, 'blipFill') || child(spPr, 'pattFill')) return undefined;
  if (!fillRef || (attrNum(fillRef, 'idx') ?? 0) === 0) return undefined;
  return resolveContainerColor(fillRef, colors)?.hex;
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
    ...(attrNum(anchor, 'relativeHeight') !== undefined ? { zOrder: attrNum(anchor, 'relativeHeight')! } : {}),
    wrap: wrapMode(anchor),
    distPx: {
      top: emuToPx(attrNum(anchor, 'distT') ?? 0),
      right: emuToPx(attrNum(anchor, 'distR') ?? 0),
      bottom: emuToPx(attrNum(anchor, 'distB') ?? 0),
      left: emuToPx(attrNum(anchor, 'distL') ?? 0),
    },
    relH: attr(posH, 'relativeFrom') ?? 'column',
    relV: attr(posV, 'relativeFrom') ?? 'paragraph',
  };

  const wrapEl = anchor.children.find((c) => /^wrap(Square|Tight|Through)$/.test(localName(c.name)));
  const wrapText = attr(wrapEl, 'wrapText');
  if (wrapText === 'left' || wrapText === 'right' || wrapText === 'largest' || wrapText === 'bothSides') a.wrapText = wrapText;

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
