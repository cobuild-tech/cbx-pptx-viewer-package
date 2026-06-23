/**
 * SmartArt data-model fallback layout.
 *
 * When a diagram frame has no cached `dsp:drawing` (the usual fast path handled
 * by {@link resolveDiagramDrawing}), we instead read the *data model*
 * (`data1.xml`, root `<dgm:dataModel>`) and lay the nodes out ourselves. This is
 * not the full DrawingML layout-algorithm spec — it reads the node text and the
 * layout *category* (`loCatId` on the doc node: cycle/process/list/hierarchy/
 * pyramid/…) and applies a generic arrangement per family. The result is a set
 * of model {@link Shape}s in the frame's pixel box, rendered exactly like any
 * other shapes, so the diagram shows its real content instead of a placeholder.
 */
import { child, children, attr, attrNum, type XmlNode } from '../../oxml/xml.js';
import type { Shape, PresetShape, ConnectorShape, TextBody, Paragraph, Color } from '../model.js';
import type { SlideBuildCtx } from '../shapes/props.js';

interface Node {
  id: string;
  paras: Paragraph[];
  children: Node[];
}

export function buildDiagramFallback(
  graphicData: XmlNode | undefined,
  ctx: SlideBuildCtx,
  w: number,
  h: number,
): Shape[] | undefined {
  const relIds = child(graphicData, 'relIds');
  const dmPart = relIds ? ctx.parts.partForRel(attr(relIds, 'r:dm') ?? '') : undefined;
  const dataModel = dmPart ? ctx.parts.xml(dmPart) : undefined;
  if (!dataModel) return undefined;

  const tree = parseDataModel(dataModel);
  if (!tree || !tree.roots.length) return undefined;

  const palette = accents(ctx);
  const family = tree.category;
  if (family === 'cycle') return layoutCycle(tree.roots, w, h, palette);
  if (family === 'pyramid') return layoutPyramid(tree.roots, w, h, palette);
  if (family === 'hierarchy') return layoutHierarchy(tree.roots, w, h, palette);
  if (family === 'process') return layoutSequence(tree.roots, w, h, palette, true);
  if (family === 'list' && isWide(w, h, tree.roots.length))
    return layoutSequence(tree.roots, w, h, palette, false);
  return layoutList(tree.roots, w, h, palette);
}

// ---------------------------------------------------------------------------
// Data-model parsing
// ---------------------------------------------------------------------------

function parseDataModel(dataModel: XmlNode): { roots: Node[]; category: string } | undefined {
  const ptLst = child(dataModel, 'ptLst');
  const cxnLst = child(dataModel, 'cxnLst');
  if (!ptLst) return undefined;

  const byId = new Map<string, XmlNode>();
  let docId: string | undefined;
  let category = 'list';
  for (const pt of children(ptLst, 'pt')) {
    const id = attr(pt, 'modelId');
    if (!id) continue;
    byId.set(id, pt);
    const type = attr(pt, 'type');
    if (type === 'doc') {
      docId = id;
      category = attr(child(pt, 'prSet'), 'loCatId') ?? category;
    }
  }
  if (!docId) return undefined;

  // Parent -> ordered child ids, from hierarchy (untyped "parOf") connections.
  const kids = new Map<string, Array<{ ord: number; id: string }>>();
  for (const cxn of children(cxnLst, 'cxn')) {
    if (attr(cxn, 'type')) continue; // skip presOf / presParOf transitions
    const src = attr(cxn, 'srcId');
    const dest = attr(cxn, 'destId');
    if (!src || !dest || !byId.has(dest)) continue;
    const ord = attrNum(cxn, 'srcOrd') ?? 0;
    (kids.get(src) ?? kids.set(src, []).get(src)!).push({ ord, id: dest });
  }

  const isData = (pt: XmlNode | undefined) => {
    const t = attr(pt, 'type');
    return !t || (t !== 'parTrans' && t !== 'sibTrans' && t !== 'pres' && t !== 'doc');
  };
  const build = (id: string): Node | undefined => {
    const pt = byId.get(id);
    if (!pt || !isData(pt)) return undefined;
    const childNodes = (kids.get(id) ?? [])
      .sort((a, b) => a.ord - b.ord)
      .map((c) => build(c.id))
      .filter((n): n is Node => !!n);
    return { id, paras: nodeParas(pt), children: childNodes };
  };

  const roots = (kids.get(docId) ?? [])
    .sort((a, b) => a.ord - b.ord)
    .map((c) => build(c.id))
    .filter((n): n is Node => !!n);

  return { roots, category };
}

/** Text paragraphs of a data point (`<dgm:t>`), preserving line breaks. */
function nodeParas(pt: XmlNode): Paragraph[] {
  const t = child(pt, 't');
  const out: Paragraph[] = [];
  for (const p of children(t, 'p')) {
    const text = runText(p);
    out.push({ runs: text ? [{ text }] : [], level: 0, align: 'ctr' });
  }
  return out.length ? out : [{ runs: [], level: 0, align: 'ctr' }];
}

function runText(p: XmlNode): string {
  return children(p, 'r')
    .map((r) => child(r, 't')?.text ?? '')
    .join('');
}

function hasText(n: Node): boolean {
  return n.paras.some((p) => p.runs.some((r) => r.text.trim()));
}

// ---------------------------------------------------------------------------
// Layout families
// ---------------------------------------------------------------------------

const PAD = 10;

function isWide(w: number, h: number, _n: number): boolean {
  return w > h * 1.4;
}

/** A rounded-rect node box with centered white text on an accent fill. */
function box(x: number, y: number, w: number, h: number, fill: Color, node: Node): PresetShape {
  const sizePt = Math.max(8, Math.min(13, Math.round(h * 0.16)));
  const text: TextBody = {
    paragraphs: node.paras.map((p) => ({
      ...p,
      runs: p.runs.map((r) => ({ ...r, color: { hex: 'FFFFFF' }, sizePt })),
    })),
    anchor: 'ctr',
    wrap: true,
    insets: { l: 6, t: 4, r: 6, b: 4 },
  };
  return {
    kind: 'shape',
    transform: { x, y, w, h },
    geom: { type: 'preset', preset: 'roundRect', adjust: { adj: 0.12 } },
    fill: { type: 'solid', color: fill },
    text,
  };
}

function arrow(x: number, y: number, w: number, h: number, color: Color): ConnectorShape {
  return {
    kind: 'connector',
    transform: { x, y, w, h },
    geom: { type: 'preset', preset: 'rightArrow', adjust: {} },
    stroke: { color, width: 1 },
  } as ConnectorShape;
}

function layoutList(nodes: Node[], w: number, h: number, pal: Color[]): Shape[] {
  const n = nodes.length || 1;
  const bh = (h - PAD * (n + 1)) / n;
  const bw = w - PAD * 2;
  return nodes.map((node, i) => box(PAD, PAD + i * (bh + PAD), bw, bh, pal[i % pal.length]!, node));
}

function layoutSequence(nodes: Node[], w: number, h: number, pal: Color[], arrows: boolean): Shape[] {
  const n = nodes.length || 1;
  const aw = arrows ? Math.min(w * 0.06, 34) : 0;
  const gap = PAD;
  const bw = (w - PAD * 2 - aw * (n - 1) - gap * (n - 1) * (arrows ? 0 : 1)) / n;
  const bh = Math.min(h - PAD * 2, h * 0.6);
  const y = (h - bh) / 2;
  const out: Shape[] = [];
  let x = PAD;
  nodes.forEach((node, i) => {
    out.push(box(x, y, bw, bh, pal[i % pal.length]!, node));
    x += bw;
    if (i < n - 1) {
      if (arrows) {
        out.push(arrow(x, h / 2 - aw * 0.3, aw, aw * 0.6, pal[i % pal.length]!));
        x += aw;
      } else {
        x += gap;
      }
    }
  });
  return out;
}

function layoutCycle(nodes: Node[], w: number, h: number, pal: Color[]): Shape[] {
  const n = nodes.length || 1;
  const cx = w / 2;
  const cy = h / 2;
  const ring = Math.min(w, h) * 0.34;
  const bw = Math.min(w, h) * 0.34;
  const bh = bw * 0.55;
  return nodes.map((node, i) => {
    const ang = (-90 + (360 / n) * i) * (Math.PI / 180);
    const x = cx + ring * Math.cos(ang) - bw / 2;
    const y = cy + ring * Math.sin(ang) - bh / 2;
    return box(x, y, bw, bh, pal[i % pal.length]!, node);
  });
}

function layoutPyramid(nodes: Node[], w: number, h: number, pal: Color[]): Shape[] {
  const n = nodes.length || 1;
  const bh = (h - PAD * (n + 1)) / n;
  return nodes.map((node, i) => {
    // Bands widen toward the base.
    const frac = 0.35 + 0.65 * ((i + 1) / n);
    const bw = w * frac;
    return box((w - bw) / 2, PAD + i * (bh + PAD), bw, bh, pal[i % pal.length]!, node);
  });
}

function layoutHierarchy(nodes: Node[], w: number, h: number, pal: Color[]): Shape[] {
  // Two-level tree if there's nesting, else a single row.
  const hasNesting = nodes.some((n) => n.children.length);
  if (!hasNesting) return layoutSequence(nodes, w, h, pal, false);

  const out: Shape[] = [];
  const topH = h * 0.32;
  const bottomY = topH + PAD * 2;
  const bottomH = h - bottomY - PAD;
  const colW = (w - PAD * 2) / nodes.length;
  nodes.forEach((node, i) => {
    const cx = PAD + colW * i + colW / 2;
    const bw = colW - PAD;
    out.push(box(cx - bw / 2, PAD, bw, topH, pal[i % pal.length]!, node));
    const kids = node.children.filter(hasText);
    const kh = kids.length ? (bottomH - PAD * (kids.length + 1)) / kids.length : 0;
    kids.forEach((k, j) => {
      out.push(box(cx - bw / 2, bottomY + PAD + j * (kh + PAD), bw, kh, pal[(i + 1) % pal.length]!, k));
    });
  });
  return out;
}

function accents(ctx: SlideBuildCtx): Color[] {
  const c = ctx.colorCtx.theme.colors;
  const list = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
    .map((k) => c[k])
    .filter((hex): hex is string => !!hex)
    .map((hex) => ({ hex }));
  return list.length ? list : [{ hex: '4472C4' }];
}
