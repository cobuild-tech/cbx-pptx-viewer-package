/**
 * Custom geometry (`<a:custGeom>`) -> SVG paths.
 *
 * Converts DrawingML path commands (moveTo/lnTo/cubicBezTo/quadBezTo/arcTo/close)
 * into SVG path data. Coordinates stay in the path's own EMU space (`<a:path w h>`);
 * the renderer scales them to the shape box via an SVG viewBox.
 */
import { children, child, attr, attrNum, localName, type XmlNode } from '../../../oxml/xml.js';
import type { CustomPath } from '../../model.js';

interface Pt {
  x: number;
  y: number;
}

function ptOf(node: XmlNode | undefined): Pt | undefined {
  if (!node) return undefined;
  const x = attrNum(node, 'x');
  const y = attrNum(node, 'y');
  if (x === undefined || y === undefined) return undefined;
  return { x, y };
}

export function parseCustomGeometry(custGeom: XmlNode): CustomPath[] {
  const out: CustomPath[] = [];
  for (const pathEl of children(child(custGeom, 'pathLst'), 'path')) {
    const w = attrNum(pathEl, 'w') ?? 0;
    const h = attrNum(pathEl, 'h') ?? 0;
    const segs: string[] = [];
    let cur: Pt = { x: 0, y: 0 };

    for (const cmd of pathEl.children) {
      switch (localName(cmd.name)) {
        case 'moveTo': {
          const p = ptOf(child(cmd, 'pt'));
          if (p) {
            segs.push(`M${p.x},${p.y}`);
            cur = p;
          }
          break;
        }
        case 'lnTo': {
          const p = ptOf(child(cmd, 'pt'));
          if (p) {
            segs.push(`L${p.x},${p.y}`);
            cur = p;
          }
          break;
        }
        case 'cubicBezTo': {
          const pts = children(cmd, 'pt').map(ptOf);
          const [c1, c2, e] = pts;
          if (pts.length === 3 && c1 && c2 && e) {
            segs.push(`C${c1.x},${c1.y} ${c2.x},${c2.y} ${e.x},${e.y}`);
            cur = e;
          }
          break;
        }
        case 'quadBezTo': {
          const pts = children(cmd, 'pt').map(ptOf);
          const [c1, e] = pts;
          if (pts.length === 2 && c1 && e) {
            segs.push(`Q${c1.x},${c1.y} ${e.x},${e.y}`);
            cur = e;
          }
          break;
        }
        case 'arcTo': {
          const seg = arcTo(cmd, cur);
          if (seg) {
            segs.push(seg.d);
            cur = seg.end;
          }
          break;
        }
        case 'close':
          segs.push('Z');
          break;
      }
    }
    if (segs.length) out.push({ d: segs.join(' '), w, h });
  }
  return out;
}

/**
 * DrawingML arcTo: an elliptical arc starting at the current point, with radii
 * wR/hR, start angle stAng and swept angle swAng (both in 60000ths of a degree).
 * The ellipse center is derived from the current point and start angle.
 */
function arcTo(cmd: XmlNode, cur: Pt): { d: string; end: Pt } | undefined {
  const wR = attrNum(cmd, 'wR');
  const hR = attrNum(cmd, 'hR');
  const stAng = attrNum(cmd, 'stAng');
  const swAng = attrNum(cmd, 'swAng');
  if (wR === undefined || hR === undefined || stAng === undefined || swAng === undefined) {
    return undefined;
  }
  const st = (stAng / 60000) * (Math.PI / 180);
  const sw = (swAng / 60000) * (Math.PI / 180);
  const cx = cur.x - wR * Math.cos(st);
  const cy = cur.y - hR * Math.sin(st);
  const end: Pt = {
    x: cx + wR * Math.cos(st + sw),
    y: cy + hR * Math.sin(st + sw),
  };
  const largeArc = Math.abs(swAng / 60000) > 180 ? 1 : 0;
  const sweep = swAng > 0 ? 1 : 0;
  return {
    d: `A${Math.abs(wR)},${Math.abs(hR)} 0 ${largeArc} ${sweep} ${end.x},${end.y}`,
    end,
  };
}
