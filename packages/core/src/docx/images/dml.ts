/**
 * DrawingML content inside a Word document beyond pictures and Word shapes:
 * charts, shape groups (<wpg:wgp>), SmartArt diagrams, and the DrawingML text
 * bodies (<a:txBody>) their shapes carry.
 *
 * A SmartArt diagram is stored twice: as a data model plus layout definition,
 * and as Office's own pre-drawn result (a `diagramDrawing` part). As in the
 * PPTX slice we draw the pre-drawn shapes — exactly what Word shows — rather
 * than re-running the layout algorithm.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { emuToPx } from '../../oxml/units.js';
import { resolveContainerColor } from '../../oxml/drawingml/color.js';
import { parseChart } from '../../oxml/drawingml/chart/parse.js';
import type { DocxBlock, DocxChart, DocxDrawing, DocxGroup, DocxParagraph, DocxRun, TextAlign } from '../model.js';
import type { ParseContext } from '../document/context.js';
import { shapeFrom } from './image.js';

const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const GROUP_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
const DIAGRAM_DRAWING_REL = 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing';

/** A chart, shape group or SmartArt diagram in a <a:graphicData>, sized wPx × hPx. */
export function graphicFrame(gd: XmlNode, wPx: number, hPx: number, ctx: ParseContext): DocxChart | DocxGroup | undefined {
  const uri = attr(gd, 'uri');
  if (uri === CHART_URI) {
    const rel = ctx.rel(attr(child(gd, 'chart'), 'id'));
    const xml = rel ? ctx.getPartXml(rel.target) : undefined;
    const chart = xml ? parseChart(xml, ctx.colors) : undefined;
    return chart ? { kind: 'chart', chart, widthPx: wPx, heightPx: hPx } : undefined;
  }
  if (uri === GROUP_URI) {
    const wgp = child(gd, 'wgp');
    return wgp ? group(wgp, wPx, hPx, ctx) : undefined;
  }
  if (uri === DIAGRAM_URI) return diagram(child(gd, 'relIds'), wPx, hPx, ctx);
  return undefined;
}

/**
 * A group (<wpg:wgp> or <a:grpSp>): its children are stated in the group's
 * child coordinate space (chOff/chExt) and scaled into its wPx × hPx box.
 */
function group(node: XmlNode, wPx: number, hPx: number, ctx: ParseContext): DocxGroup {
  const xfrm = child(child(node, 'grpSpPr'), 'xfrm');
  const chOff = child(xfrm, 'chOff');
  const chExt = child(xfrm, 'chExt');
  const ext = child(xfrm, 'ext');
  const cw = attrNum(chExt, 'cx') ?? attrNum(ext, 'cx') ?? 1;
  const ch = attrNum(chExt, 'cy') ?? attrNum(ext, 'cy') ?? 1;
  const map = {
    x0: attrNum(chOff, 'x') ?? 0,
    y0: attrNum(chOff, 'y') ?? 0,
    sx: wPx / (emuToPx(cw) || 1),
    sy: hPx / (emuToPx(ch) || 1),
  };
  return { kind: 'group', widthPx: wPx, heightPx: hPx, items: groupItems(node, map, ctx) };
}

interface Mapping {
  x0: number;
  y0: number;
  sx: number;
  sy: number;
}

/** The drawable children of a group or shape tree, placed through `map`. */
function groupItems(container: XmlNode, map: Mapping, ctx: ParseContext): DocxGroup['items'] {
  const items: DocxGroup['items'] = [];
  for (const c of container.children) {
    const name = localName(c.name);
    if (name !== 'wsp' && name !== 'sp' && name !== 'grpSp' && name !== 'pic' && name !== 'cxnSp') continue;
    const xfrm = child(child(c, name === 'grpSp' ? 'grpSpPr' : 'spPr'), 'xfrm');
    const off = child(xfrm, 'off');
    const ext = child(xfrm, 'ext');
    const xPx = emuToPx((attrNum(off, 'x') ?? 0) - map.x0) * map.sx;
    const yPx = emuToPx((attrNum(off, 'y') ?? 0) - map.y0) * map.sy;
    const wPx = emuToPx(attrNum(ext, 'cx') ?? 0) * map.sx;
    const hPx = emuToPx(attrNum(ext, 'cy') ?? 0) * map.sy;
    let drawing: DocxDrawing | undefined;
    if (name === 'grpSp') drawing = group(c, wPx, hPx, ctx);
    else if (name === 'pic') {
      const blip = child(child(c, 'blipFill'), 'blip');
      const rel = ctx.rel(attr(blip, 'embed') ?? attr(blip, 'link'));
      if (rel) drawing = { kind: 'image', part: rel.target, widthPx: wPx, heightPx: hPx };
    } else {
      const shape = shapeFrom(c, wPx, hPx, undefined, ctx);
      // The group item carries the rotation; the shape draws upright inside it.
      const rot = shape.rot;
      delete shape.rot;
      drawing = shape;
      if (rot) {
        items.push({ xPx, yPx, wPx, hPx, rot, drawing });
        continue;
      }
    }
    if (drawing) items.push({ xPx, yPx, wPx, hPx, drawing });
  }
  return items;
}

/**
 * A SmartArt diagram: find its pre-drawn shape tree — through the data model's
 * <dsp:dataModelExt relId>, else the diagramDrawing relationship whose part
 * number matches the data model's — and lay it out in the frame box.
 */
function diagram(relIds: XmlNode | undefined, wPx: number, hPx: number, ctx: ParseContext): DocxGroup | undefined {
  const dataPart = ctx.rel(attr(relIds, 'dm'))?.target;
  let drawingPart: string | undefined;
  const dataXml = dataPart ? ctx.getPartXml(dataPart) : undefined;
  const ext = dataXml ? findFirst(dataXml, 'dataModelExt') : undefined;
  const viaExt = ctx.rel(attr(ext, 'relId'));
  if (viaExt && ctx.getPartXml(viaExt.target)) drawingPart = viaExt.target;
  if (!drawingPart) {
    const drawings = ctx.relsByType(DIAGRAM_DRAWING_REL).map((r) => r.target);
    const n = dataPart?.match(/(\d+)\.xml$/i)?.[1];
    drawingPart = drawings.find((d) => d.match(/(\d+)\.xml$/i)?.[1] === n) ?? drawings[0];
  }
  const spTree = drawingPart ? child(ctx.getPartXml(drawingPart), 'spTree') : undefined;
  if (!spTree || !drawingPart) return undefined;
  // The drawing is authored in the frame's own EMU space; its pictures resolve
  // against the drawing part's relationships.
  const partCtx = ctx.forPart(drawingPart);
  return { kind: 'group', widthPx: wPx, heightPx: hPx, items: groupItems(spTree, { x0: 0, y0: 0, sx: 1, sy: 1 }, partCtx) };
}

function findFirst(node: XmlNode, name: string): XmlNode | undefined {
  for (const c of node.children) {
    if (localName(c.name) === name) return c;
    const hit = findFirst(c, name);
    if (hit) return hit;
  }
  return undefined;
}

// ─── DrawingML text ────────────────────────────────────────────────────────────

const ALIGN: Record<string, TextAlign> = { l: 'l', ctr: 'ctr', r: 'r', just: 'just', dist: 'just' };

/**
 * A DrawingML text body (<a:txBody>) as paragraphs. `textHex` is the shape
 * style's font colour, for runs that state none.
 */
export function dmlTextBlocks(txBody: XmlNode, ctx: ParseContext, textHex?: string): DocxBlock[] {
  // Autofit shrinks the text by a stated factor (<a:normAutofit fontScale>).
  const fontScale = (attrNum(child(child(txBody, 'bodyPr'), 'normAutofit'), 'fontScale') ?? 100000) / 100000;
  const defRPr = child(child(child(txBody, 'lstStyle'), 'lvl1pPr'), 'defRPr');
  const out: DocxBlock[] = [];
  for (const p of children(txBody, 'p')) {
    const pPr = child(p, 'pPr');
    const runs: DocxRun[] = [];
    let pendingBreak = false;
    for (const c of p.children) {
      const name = localName(c.name);
      if (name === 'br') {
        pendingBreak = true;
        continue;
      }
      if (name !== 'r' && name !== 'fld') continue;
      const run = dmlRun(child(c, 'rPr') ?? defRPr, child(c, 't')?.text ?? '', ctx, fontScale, textHex);
      if (pendingBreak) run.breakBefore = true;
      pendingBreak = false;
      runs.push(run);
    }
    // An empty paragraph is as tall as its end mark.
    const endRun = dmlRun(child(p, 'endParaRPr') ?? defRPr, '', ctx, fontScale, textHex);
    const first = runs[0] ?? endRun;
    const para: DocxParagraph = { kind: 'paragraph', runs, styleName: 'DrawingML' };
    para.baseFontSizePt = first.sizePt ?? 18 * fontScale;
    if (first.font) para.baseFontFamily = first.font;
    if (first.colorHex) para.baseColorHex = first.colorHex;
    const algn = attr(pPr, 'algn');
    if (algn && ALIGN[algn]) para.align = ALIGN[algn];
    const lnSpcPct = attrNum(child(child(pPr, 'lnSpc'), 'spcPct'), 'val');
    if (lnSpcPct !== undefined) para.lineSpacingPct = lnSpcPct / 100000;
    const lnSpcPts = attrNum(child(child(pPr, 'lnSpc'), 'spcPts'), 'val');
    if (lnSpcPts !== undefined) {
      para.lineSpacingPt = lnSpcPts / 100;
      para.lineSpacingRule = 'exact';
    }
    const before = attrNum(child(child(pPr, 'spcBef'), 'spcPts'), 'val');
    if (before !== undefined) para.spaceBeforePt = before / 100;
    const after = attrNum(child(child(pPr, 'spcAft'), 'spcPts'), 'val');
    if (after !== undefined) para.spaceAfterPt = after / 100;
    out.push(para);
  }
  return out;
}

function dmlRun(rPr: XmlNode | undefined, text: string, ctx: ParseContext, scale: number, textHex?: string): DocxRun {
  const run: DocxRun = { text };
  const sz = attrNum(rPr, 'sz');
  run.sizePt = (sz !== undefined ? sz / 100 : 18) * scale;
  if (attr(rPr, 'b') === '1' || attr(rPr, 'b') === 'true') run.bold = true;
  if (attr(rPr, 'i') === '1' || attr(rPr, 'i') === 'true') run.italic = true;
  const u = attr(rPr, 'u');
  if (u && u !== 'none') run.underline = true;
  const strike = attr(rPr, 'strike');
  if (strike && strike !== 'noStrike') run.strike = true;
  const color = resolveContainerColor(child(rPr, 'solidFill'), ctx.colors)?.hex ?? textHex;
  if (color) run.colorHex = color;
  // Theme font tokens: +mj-lt / +mn-lt (major/minor latin).
  const face = attr(child(rPr, 'latin'), 'typeface');
  const font = face === '+mj-lt' ? ctx.styles.fontName('+major') : face === '+mn-lt' || !face ? ctx.styles.fontName('+minor') : face;
  if (font) run.font = font;
  return run;
}
