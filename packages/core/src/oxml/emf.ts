/**
 * EMF (Enhanced Metafile) -> SVG.
 *
 * Office stores vector clip-art, logos and decorative artwork as EMF, which no
 * browser can decode, so a picture whose blip is an `.emf` would otherwise draw
 * nothing. This replays the metafile's GDI records — the same playback Windows
 * does — and emits the result as SVG markup a browser can show through `<img>`.
 *
 * Coverage is the record set real Office files use: map modes, window/viewport
 * and world transforms, saved DC states, pens/brushes/fonts (incl. stock
 * objects), paths (fill, stroke, clip), polygons/polylines/Béziers in their 16-
 * and 32-bit forms, rectangles/ellipses/arcs, clip rectangles and regions,
 * device-independent bitmaps and text. EMF+ records (in GDI comments) are
 * skipped: a dual EMF+ file carries the same drawing as GDI records too.
 *
 * Every point is mapped to device units on the way in (world transform, then
 * the window->viewport page transform), so the SVG needs no nested transforms.
 * The viewBox is the header's `rclFrame` in device units — the rectangle Office
 * maps onto the picture box.
 */

// Record types (MS-EMF 2.1.1).
const EMR = {
  HEADER: 1,
  POLYBEZIER: 2,
  POLYGON: 3,
  POLYLINE: 4,
  POLYBEZIERTO: 5,
  POLYLINETO: 6,
  POLYPOLYLINE: 7,
  POLYPOLYGON: 8,
  SETWINDOWEXTEX: 9,
  SETWINDOWORGEX: 10,
  SETVIEWPORTEXTEX: 11,
  SETVIEWPORTORGEX: 12,
  EOF: 14,
  SETMAPMODE: 17,
  SETPOLYFILLMODE: 19,
  SETTEXTALIGN: 22,
  SETTEXTCOLOR: 24,
  MOVETOEX: 27,
  INTERSECTCLIPRECT: 30,
  SCALEVIEWPORTEXTEX: 31,
  SCALEWINDOWEXTEX: 32,
  SAVEDC: 33,
  RESTOREDC: 34,
  SETWORLDTRANSFORM: 35,
  MODIFYWORLDTRANSFORM: 36,
  SELECTOBJECT: 37,
  CREATEPEN: 38,
  CREATEBRUSHINDIRECT: 39,
  DELETEOBJECT: 40,
  ANGLEARC: 41,
  ELLIPSE: 42,
  RECTANGLE: 43,
  ROUNDRECT: 44,
  ARC: 45,
  CHORD: 46,
  PIE: 47,
  LINETO: 54,
  ARCTO: 55,
  BEGINPATH: 59,
  ENDPATH: 60,
  CLOSEFIGURE: 61,
  FILLPATH: 62,
  STROKEANDFILLPATH: 63,
  STROKEPATH: 64,
  SELECTCLIPPATH: 67,
  ABORTPATH: 68,
  EXTSELECTCLIPRGN: 75,
  BITBLT: 76,
  STRETCHBLT: 77,
  STRETCHDIBITS: 81,
  EXTCREATEFONTINDIRECTW: 82,
  EXTTEXTOUTW: 84,
  POLYBEZIER16: 85,
  POLYGON16: 86,
  POLYLINE16: 87,
  POLYBEZIERTO16: 88,
  POLYLINETO16: 89,
  POLYPOLYLINE16: 90,
  POLYPOLYGON16: 91,
  CREATEDIBPATTERNBRUSHPT: 94,
  EXTCREATEPEN: 95,
  ALPHABLEND: 114,
} as const;

const MM_TEXT = 1;
const MM_ISOTROPIC = 7;
const MM_ANISOTROPIC = 8;
/** Logical-unit size in mm for the fixed map modes MM_LOMETRIC..MM_TWIPS. */
const MAP_MODE_MM: Record<number, number> = { 2: 0.1, 3: 0.01, 4: 0.254, 5: 0.0254, 6: 25.4 / 1440 };

const RGN_AND = 1;
const RGN_COPY = 5;

type Matrix = [number, number, number, number, number, number]; // a b c d e f
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

/** `m` applied after `n` (x' = m(n(x))). */
function multiply(m: Matrix, n: Matrix): Matrix {
  return [
    m[0] * n[0] + m[2] * n[1],
    m[1] * n[0] + m[3] * n[1],
    m[0] * n[2] + m[2] * n[3],
    m[1] * n[2] + m[3] * n[3],
    m[0] * n[4] + m[2] * n[5] + m[4],
    m[1] * n[4] + m[3] * n[5] + m[5],
  ];
}

interface Pen {
  kind: 'pen';
  none: boolean;
  width: number;
  /** Cosmetic pens are always one device pixel wide. */
  cosmetic: boolean;
  color: string;
  dash?: number[];
  cap: 'round' | 'square' | 'butt';
  join: 'round' | 'bevel' | 'miter';
}

interface Brush {
  kind: 'brush';
  none: boolean;
  color: string;
  /** A bitmap pattern brush, as an image data URL. */
  pattern?: { url: string; w: number; h: number };
}

interface Font {
  kind: 'font';
  height: number;
  weight: number;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  escapement: number;
  face: string;
}

type GdiObject = Pen | Brush | Font;

interface DcState {
  pen: Pen;
  brush: Brush;
  font: Font;
  fillRule: 'evenodd' | 'nonzero';
  textColor: string;
  textAlign: number;
  mapMode: number;
  winOrg: [number, number];
  winExt: [number, number];
  vpOrg: [number, number];
  vpExt: [number, number];
  world: Matrix;
  miterLimit: number;
  /** Id of the active SVG clipPath, if any. */
  clip?: string;
}

const BLACK_PEN: Pen = { kind: 'pen', none: false, width: 0, cosmetic: true, color: '#000', cap: 'round', join: 'round' };
const NULL_PEN: Pen = { ...BLACK_PEN, none: true };
const WHITE_BRUSH: Brush = { kind: 'brush', none: false, color: '#fff' };
const NULL_BRUSH: Brush = { kind: 'brush', none: true, color: '#000' };
const DEFAULT_FONT: Font = {
  kind: 'font',
  height: 12,
  weight: 400,
  italic: false,
  underline: false,
  strike: false,
  escapement: 0,
  face: 'Arial',
};

/** GDI stock objects, selected with the high bit of the handle set. */
const STOCK: Record<number, GdiObject> = {
  0: WHITE_BRUSH,
  1: { kind: 'brush', none: false, color: '#c0c0c0' },
  2: { kind: 'brush', none: false, color: '#808080' },
  3: { kind: 'brush', none: false, color: '#404040' },
  4: { kind: 'brush', none: false, color: '#000' },
  5: NULL_BRUSH,
  6: { ...BLACK_PEN, color: '#fff' },
  7: BLACK_PEN,
  8: NULL_PEN,
  10: DEFAULT_FONT,
  11: DEFAULT_FONT,
  12: DEFAULT_FONT,
  13: DEFAULT_FONT,
  14: DEFAULT_FONT,
  16: DEFAULT_FONT,
  17: DEFAULT_FONT,
  18: WHITE_BRUSH,
  19: BLACK_PEN,
};

/** True when the bytes start with an EMF header record. */
export function isEmf(bytes: Uint8Array): boolean {
  if (bytes.length < 88) return false;
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return dv.getUint32(0, true) === EMR.HEADER && dv.getUint32(40, true) === 0x464d4520; // ' EMF'
}

/**
 * Convert EMF bytes to an SVG document. Returns undefined for something that
 * is not an EMF or has no drawable extent.
 */
export function emfToSvg(bytes: Uint8Array): string | undefined {
  if (!isEmf(bytes)) return undefined;
  return new EmfPlayer(bytes).play();
}

class EmfPlayer {
  private dv: DataView;
  private objects = new Map<number, GdiObject>();
  private dc: DcState;
  private saved: DcState[] = [];
  private matrix: Matrix = IDENTITY;
  /** Device units per millimetre, from the header's reference device. */
  private devPerMm = 96 / 25.4;
  private out: string[] = [];
  private defs: string[] = [];
  private nextId = 0;
  private readonly idPrefix = `emf${Math.random().toString(36).slice(2, 8)}`;
  /** Current position, in device units. */
  private cp: [number, number] = [0, 0];
  /** Path being recorded between BEGINPATH and ENDPATH. */
  private recording = false;
  /** The last closed path (ENDPATH), consumed by FILL/STROKE/SELECTCLIPPATH. */
  private path = '';

  constructor(private bytes: Uint8Array) {
    this.dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    this.dc = {
      pen: BLACK_PEN,
      brush: WHITE_BRUSH,
      font: DEFAULT_FONT,
      fillRule: 'evenodd',
      textColor: '#000',
      textAlign: 0,
      mapMode: MM_TEXT,
      winOrg: [0, 0],
      winExt: [1, 1],
      vpOrg: [0, 0],
      vpExt: [1, 1],
      world: IDENTITY,
      miterLimit: 10,
    };
  }

  play(): string | undefined {
    const dv = this.dv;
    const frame = [24, 28, 32, 36].map((o) => dv.getInt32(o, true));
    const bounds = [8, 12, 16, 20].map((o) => dv.getInt32(o, true));
    // The header extensions (pixel format @88, szlMicrometers @100) exist only
    // where the description string does not already occupy those bytes.
    const headerSize = dv.getUint32(4, true);
    const nDesc = dv.getUint32(60, true);
    const offDesc = dv.getUint32(64, true);
    const headerEnd = nDesc && offDesc ? Math.min(headerSize, offDesc) : headerSize;
    const devW = dv.getInt32(72, true);
    const mmW = dv.getInt32(80, true);
    const umW = headerEnd >= 108 ? dv.getInt32(100, true) : 0;
    if (devW > 0 && umW > 0) this.devPerMm = devW / (umW / 1000);
    else if (devW > 0 && mmW > 0) this.devPerMm = devW / mmW;

    // rclFrame is inclusive, in 0.01 mm; Office maps it onto the picture box.
    const k = this.devPerMm / 100;
    let vx = frame[0]! * k;
    let vy = frame[1]! * k;
    let vw = (frame[2]! - frame[0]!) * k;
    let vh = (frame[3]! - frame[1]!) * k;
    if (!(vw > 0 && vh > 0)) {
      vx = bounds[0]!;
      vy = bounds[1]!;
      vw = bounds[2]! - bounds[0]! + 1;
      vh = bounds[3]! - bounds[1]! + 1;
    }
    if (!(vw > 0 && vh > 0)) return undefined;

    let off = 0;
    while (off + 8 <= this.bytes.length) {
      const type = dv.getUint32(off, true);
      const size = dv.getUint32(off + 4, true);
      if (size < 8 || off + size > this.bytes.length) break;
      if (type === EMR.EOF) break;
      try {
        this.record(type, off, size);
      } catch {
        // A malformed record only loses itself, never the whole picture.
      }
      off += size;
    }

    const defs = this.defs.length ? `<defs>${this.defs.join('')}</defs>` : '';
    return (
      `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" ` +
      `viewBox="${n(vx)} ${n(vy)} ${n(vw)} ${n(vh)}" width="${n(vw)}" height="${n(vh)}" ` +
      `preserveAspectRatio="none">${defs}${this.out.join('')}</svg>`
    );
  }

  // ─── Records ────────────────────────────────────────────────────────────────

  private record(type: number, off: number, size: number): void {
    const dv = this.dv;
    const i32 = (o: number) => dv.getInt32(off + o, true);
    const u32 = (o: number) => dv.getUint32(off + o, true);
    switch (type) {
      case EMR.SETMAPMODE:
        this.dc.mapMode = u32(8);
        this.updateMatrix();
        break;
      case EMR.SETWINDOWEXTEX:
        this.dc.winExt = [i32(8), i32(12)];
        this.updateMatrix();
        break;
      case EMR.SETWINDOWORGEX:
        this.dc.winOrg = [i32(8), i32(12)];
        this.updateMatrix();
        break;
      case EMR.SETVIEWPORTEXTEX:
        this.dc.vpExt = [i32(8), i32(12)];
        this.updateMatrix();
        break;
      case EMR.SETVIEWPORTORGEX:
        this.dc.vpOrg = [i32(8), i32(12)];
        this.updateMatrix();
        break;
      case EMR.SCALEVIEWPORTEXTEX:
      case EMR.SCALEWINDOWEXTEX: {
        const ext = type === EMR.SCALEWINDOWEXTEX ? this.dc.winExt : this.dc.vpExt;
        const [xNum, xDen, yNum, yDen] = [i32(8), i32(12), i32(16), i32(20)];
        if (xDen && yDen) {
          const scaled: [number, number] = [(ext[0] * xNum) / xDen, (ext[1] * yNum) / yDen];
          if (type === EMR.SCALEWINDOWEXTEX) this.dc.winExt = scaled;
          else this.dc.vpExt = scaled;
          this.updateMatrix();
        }
        break;
      }
      case EMR.SETWORLDTRANSFORM:
        this.dc.world = this.readXform(off + 8);
        this.updateMatrix();
        break;
      case EMR.MODIFYWORLDTRANSFORM: {
        const x = this.readXform(off + 8);
        const mode = u32(32);
        if (mode === 1) this.dc.world = IDENTITY;
        else if (mode === 2) this.dc.world = multiply(this.dc.world, x);
        else if (mode === 3) this.dc.world = multiply(x, this.dc.world);
        else if (mode === 4) this.dc.world = x;
        this.updateMatrix();
        break;
      }
      case EMR.SAVEDC:
        this.saved.push({ ...this.dc });
        break;
      case EMR.RESTOREDC: {
        const rel = i32(8);
        const index = rel < 0 ? this.saved.length + rel : rel - 1;
        if (index >= 0 && index < this.saved.length) {
          this.dc = { ...this.saved[index]! };
          this.saved.length = index;
          this.updateMatrix();
        }
        break;
      }
      case EMR.SETPOLYFILLMODE:
        this.dc.fillRule = u32(8) === 2 ? 'nonzero' : 'evenodd';
        break;
      case EMR.SETTEXTCOLOR:
        this.dc.textColor = this.colorRef(off + 8);
        break;
      case EMR.SETTEXTALIGN:
        this.dc.textAlign = u32(8);
        break;
      case 58: // EMR_SETMITERLIMIT
        this.dc.miterLimit = u32(8) || 10;
        break;
      case EMR.CREATEPEN: {
        const style = u32(12);
        const width = i32(16);
        this.objects.set(u32(8), this.makePen(style, width, this.colorRef(off + 24), false));
        break;
      }
      case EMR.EXTCREATEPEN: {
        const style = u32(28);
        const width = u32(32);
        const brushStyle = u32(36);
        const pen = this.makePen(style, width, this.colorRef(off + 40), true);
        if (brushStyle === 1) pen.none = true; // BS_NULL
        this.objects.set(u32(8), pen);
        break;
      }
      case EMR.CREATEBRUSHINDIRECT: {
        const style = u32(12);
        // BS_SOLID and BS_HATCHED draw in the brush colour; BS_NULL draws nothing.
        this.objects.set(u32(8), { kind: 'brush', none: style === 1, color: this.colorRef(off + 16) });
        break;
      }
      case EMR.CREATEDIBPATTERNBRUSHPT: {
        const dib = this.dibUrl(off, u32(16), u32(20), u32(24), u32(28));
        this.objects.set(
          u32(8),
          dib ? { kind: 'brush', none: false, color: '#808080', pattern: dib } : { ...NULL_BRUSH },
        );
        break;
      }
      case EMR.EXTCREATEFONTINDIRECTW:
        this.objects.set(u32(8), this.readFont(off + 12));
        break;
      case EMR.SELECTOBJECT: {
        const h = u32(8);
        const obj = h & 0x80000000 ? STOCK[h & 0x7fffffff] : this.objects.get(h);
        if (obj?.kind === 'pen') this.dc.pen = obj;
        else if (obj?.kind === 'brush') this.dc.brush = obj;
        else if (obj?.kind === 'font') this.dc.font = obj;
        break;
      }
      case EMR.DELETEOBJECT:
        this.objects.delete(u32(8));
        break;

      // Paths.
      case EMR.BEGINPATH:
        this.recording = true;
        this.path = '';
        break;
      case EMR.ENDPATH:
        this.recording = false;
        break;
      case EMR.ABORTPATH:
        this.recording = false;
        this.path = '';
        break;
      case EMR.CLOSEFIGURE:
        if (this.recording) this.path += 'Z';
        break;
      case EMR.FILLPATH:
        this.emitPath(this.path, true, false);
        this.path = '';
        break;
      case EMR.STROKEPATH:
        this.emitPath(this.path, false, true);
        this.path = '';
        break;
      case EMR.STROKEANDFILLPATH:
        this.emitPath(this.path, true, true);
        this.path = '';
        break;
      case EMR.SELECTCLIPPATH:
        this.setClip(this.path, u32(8));
        this.path = '';
        break;

      // Lines and curves that move the current position.
      case EMR.MOVETOEX: {
        this.cp = this.map(i32(8), i32(12));
        if (this.recording) this.path += `M${pt(this.cp)}`;
        break;
      }
      case EMR.LINETO: {
        const p = this.map(i32(8), i32(12));
        this.drawOrRecord(`M${pt(this.cp)}L${pt(p)}`, `L${pt(p)}`, false);
        this.cp = p;
        break;
      }
      case EMR.POLYLINETO16:
      case EMR.POLYLINETO:
      case EMR.POLYBEZIERTO16:
      case EMR.POLYBEZIERTO: {
        const pts = this.readPoints(off + 28, u32(24), type === EMR.POLYLINETO16 || type === EMR.POLYBEZIERTO16);
        const bez = type === EMR.POLYBEZIERTO16 || type === EMR.POLYBEZIERTO;
        const seg = bez ? bezierSegs(pts) : pts.map((p) => `L${pt(p)}`).join('');
        this.drawOrRecord(`M${pt(this.cp)}${seg}`, seg, false);
        if (pts.length) this.cp = pts[pts.length - 1]!;
        break;
      }

      // Self-contained figures.
      case EMR.POLYLINE16:
      case EMR.POLYLINE:
      case EMR.POLYGON16:
      case EMR.POLYGON:
      case EMR.POLYBEZIER16:
      case EMR.POLYBEZIER: {
        const short = type === EMR.POLYLINE16 || type === EMR.POLYGON16 || type === EMR.POLYBEZIER16;
        const pts = this.readPoints(off + 28, u32(24), short);
        if (!pts.length) break;
        const closed = type === EMR.POLYGON16 || type === EMR.POLYGON;
        const bez = type === EMR.POLYBEZIER16 || type === EMR.POLYBEZIER;
        const d = `M${pt(pts[0]!)}${bez ? bezierSegs(pts.slice(1)) : pts.slice(1).map((p) => `L${pt(p)}`).join('')}${closed ? 'Z' : ''}`;
        this.drawOrRecord(d, d, closed);
        break;
      }
      case EMR.POLYPOLYLINE16:
      case EMR.POLYPOLYLINE:
      case EMR.POLYPOLYGON16:
      case EMR.POLYPOLYGON: {
        const short = type === EMR.POLYPOLYLINE16 || type === EMR.POLYPOLYGON16;
        const closed = type === EMR.POLYPOLYGON16 || type === EMR.POLYPOLYGON;
        const nPolys = u32(24);
        const total = u32(28);
        const pts = this.readPoints(off + 32 + nPolys * 4, total, short);
        let d = '';
        let at = 0;
        for (let i = 0; i < nPolys; i++) {
          const count = u32(32 + i * 4);
          const poly = pts.slice(at, at + count);
          at += count;
          if (!poly.length) continue;
          d += `M${pt(poly[0]!)}${poly.slice(1).map((p) => `L${pt(p)}`).join('')}${closed ? 'Z' : ''}`;
        }
        this.drawOrRecord(d, d, closed);
        break;
      }
      case EMR.RECTANGLE: {
        const d = this.rectPath(i32(8), i32(12), i32(16), i32(20));
        this.drawOrRecord(d, d, true);
        break;
      }
      case EMR.ROUNDRECT: {
        const d = this.roundRectPath(i32(8), i32(12), i32(16), i32(20), i32(24), i32(28));
        this.drawOrRecord(d, d, true);
        break;
      }
      case EMR.ELLIPSE: {
        const d = this.ellipsePath(i32(8), i32(12), i32(16), i32(20));
        this.drawOrRecord(d, d, true);
        break;
      }
      case EMR.ARC:
      case EMR.ARCTO:
      case EMR.CHORD:
      case EMR.PIE: {
        const box = [i32(8), i32(12), i32(16), i32(20)] as const;
        const arc = this.arcPoints(box, [i32(24), i32(28)], [i32(32), i32(36)]);
        if (!arc.length) break;
        const cx = (box[0] + box[2]) / 2;
        const cy = (box[1] + box[3]) / 2;
        let d: string;
        if (type === EMR.ARCTO) {
          d = `L${pt(arc[0]!)}${bezierSegs(arc.slice(1))}`;
          this.drawOrRecord(`M${pt(this.cp)}${d}`, d, false);
          this.cp = arc[arc.length - 1]!;
          break;
        }
        d = `M${pt(arc[0]!)}${bezierSegs(arc.slice(1))}`;
        if (type === EMR.CHORD) d += 'Z';
        if (type === EMR.PIE) d += `L${pt(this.map(cx, cy))}Z`;
        this.drawOrRecord(d, d, type !== EMR.ARC);
        break;
      }

      // Clipping.
      case EMR.INTERSECTCLIPRECT:
        this.setClip(this.rectPath(i32(8), i32(12), i32(16), i32(20)), RGN_AND);
        break;
      case EMR.EXTSELECTCLIPRGN: {
        const cb = u32(8);
        const mode = u32(12);
        if (mode === RGN_COPY && cb === 0) {
          delete this.dc.clip;
          break;
        }
        // RGNDATA: a 32-byte header then nCount device-space rectangles.
        const count = u32(16 + 8);
        let d = '';
        for (let i = 0; i < count; i++) {
          const o = 16 + 32 + i * 16;
          const [l, t, r, b] = [i32(o), i32(o + 4), i32(o + 8), i32(o + 12)];
          d += `M${l},${t}H${r}V${b}H${l}Z`;
        }
        this.setClip(d, mode);
        break;
      }

      // Bitmaps.
      case EMR.STRETCHDIBITS: {
        const url = this.dibUrl(off, u32(48), u32(52), u32(56), u32(60));
        // xDest, yDest @24; xSrc, ySrc, cxSrc, cySrc @32; cxDest, cyDest @72.
        this.emitImage(url, i32(24), i32(28), i32(72), i32(76), i32(32), i32(36), i32(40), i32(44));
        break;
      }
      case EMR.BITBLT:
      case EMR.STRETCHBLT:
      case EMR.ALPHABLEND: {
        // Destination @24, source origin @44, bitmap offsets @84; STRETCHBLT and
        // ALPHABLEND add the source extent @100.
        const [x, y, w, h] = [i32(24), i32(28), i32(32), i32(36)];
        const cbBmi = u32(88);
        if (!cbBmi) {
          // No source bitmap: a pattern blit fills the rectangle with the brush.
          if (type === EMR.BITBLT) this.emitPath(this.rectPath(x, y, x + w, y + h), true, false);
          break;
        }
        const url = this.dibUrl(off, u32(84), cbBmi, u32(92), u32(96));
        const [sx, sy] = [i32(44), i32(48)];
        const srcW = type === EMR.BITBLT ? w : i32(100);
        const srcH = type === EMR.BITBLT ? h : i32(104);
        this.emitImage(url, x, y, w, h, sx, sy, srcW, srcH);
        break;
      }

      case EMR.EXTTEXTOUTW:
        this.text(off, size);
        break;
    }
  }

  // ─── Coordinate mapping ─────────────────────────────────────────────────────

  private updateMatrix(): void {
    const dc = this.dc;
    let sx = 1;
    let sy = 1;
    if (dc.mapMode === MM_ISOTROPIC || dc.mapMode === MM_ANISOTROPIC) {
      sx = dc.winExt[0] ? dc.vpExt[0] / dc.winExt[0] : 1;
      sy = dc.winExt[1] ? dc.vpExt[1] / dc.winExt[1] : 1;
      if (dc.mapMode === MM_ISOTROPIC) {
        const s = Math.min(Math.abs(sx), Math.abs(sy));
        sx = Math.sign(sx || 1) * s;
        sy = Math.sign(sy || 1) * s;
      }
    } else if (MAP_MODE_MM[dc.mapMode] !== undefined) {
      sx = MAP_MODE_MM[dc.mapMode]! * this.devPerMm;
      sy = -sx;
    }
    const page: Matrix = [sx, 0, 0, sy, dc.vpOrg[0] - dc.winOrg[0] * sx, dc.vpOrg[1] - dc.winOrg[1] * sy];
    this.matrix = multiply(page, dc.world);
  }

  private map(x: number, y: number): [number, number] {
    const m = this.matrix;
    return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
  }

  /** Average linear scale of the logical->device mapping (for pen widths). */
  private scale(): number {
    const m = this.matrix;
    return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
  }

  private readXform(o: number): Matrix {
    const f = (k: number) => this.dv.getFloat32(o + k * 4, true);
    return [f(0), f(1), f(2), f(3), f(4), f(5)];
  }

  private readPoints(o: number, count: number, short: boolean): [number, number][] {
    const pts: [number, number][] = [];
    const step = short ? 4 : 8;
    const max = Math.min(count, Math.floor((this.bytes.length - o) / step));
    for (let i = 0; i < max; i++) {
      const at = o + i * step;
      const x = short ? this.dv.getInt16(at, true) : this.dv.getInt32(at, true);
      const y = short ? this.dv.getInt16(at + 2, true) : this.dv.getInt32(at + 4, true);
      pts.push(this.map(x, y));
    }
    return pts;
  }

  // ─── Geometry helpers (logical in, device-space path out) ───────────────────

  private rectPath(l: number, t: number, r: number, b: number): string {
    return `M${pt(this.map(l, t))}L${pt(this.map(r, t))}L${pt(this.map(r, b))}L${pt(this.map(l, b))}Z`;
  }

  private ellipsePath(l: number, t: number, r: number, b: number): string {
    const cx = (l + r) / 2;
    const cy = (t + b) / 2;
    const pts = ellipseBeziers(cx, cy, (r - l) / 2, (b - t) / 2, 0, Math.PI * 2).map(([x, y]) => this.map(x, y));
    return `M${pt(pts[0]!)}${bezierSegs(pts.slice(1))}Z`;
  }

  private roundRectPath(l: number, t: number, r: number, b: number, cw: number, ch: number): string {
    const rx = Math.min(Math.abs(cw) / 2, Math.abs(r - l) / 2);
    const ry = Math.min(Math.abs(ch) / 2, Math.abs(b - t) / 2);
    if (!rx || !ry) return this.rectPath(l, t, r, b);
    const corner = (cx: number, cy: number, a0: number) =>
      ellipseBeziers(cx, cy, rx, ry, a0, Math.PI / 2).map(([x, y]) => this.map(x, y));
    const segs = [
      corner(r - rx, t + ry, -Math.PI / 2),
      corner(r - rx, b - ry, 0),
      corner(l + rx, b - ry, Math.PI / 2),
      corner(l + rx, t + ry, Math.PI),
    ];
    let d = `M${pt(segs[0]![0]!)}`;
    for (let i = 0; i < 4; i++) {
      if (i) d += `L${pt(segs[i]![0]!)}`;
      d += bezierSegs(segs[i]!.slice(1));
    }
    return `${d}Z`;
  }

  /** Bézier points of the counter-clockwise (GDI default) arc from start to end radial. */
  private arcPoints(
    box: readonly [number, number, number, number],
    start: [number, number],
    end: [number, number],
  ): [number, number][] {
    const cx = (box[0] + box[2]) / 2;
    const cy = (box[1] + box[3]) / 2;
    const rx = (box[2] - box[0]) / 2;
    const ry = (box[3] - box[1]) / 2;
    if (!rx || !ry) return [];
    const a0 = Math.atan2((start[1] - cy) / ry, (start[0] - cx) / rx);
    let a1 = Math.atan2((end[1] - cy) / ry, (end[0] - cx) / rx);
    // Counter-clockwise on screen (y down) means decreasing angle.
    while (a1 >= a0) a1 -= Math.PI * 2;
    return ellipseBeziers(cx, cy, rx, ry, a0, a1 - a0).map(([x, y]) => this.map(x, y));
  }

  // ─── Output ─────────────────────────────────────────────────────────────────

  /** Inside BEGINPATH..ENDPATH a figure joins the path; otherwise it is drawn now. */
  private drawOrRecord(drawD: string, pathD: string, filled: boolean): void {
    if (this.recording) {
      this.path += pathD.startsWith('M') || this.path ? pathD : `M${pt(this.cp)}${pathD}`;
      return;
    }
    this.emitPath(drawD, filled, true);
  }

  private emitPath(d: string, fill: boolean, stroke: boolean): void {
    if (!d) return;
    const dc = this.dc;
    const attrs: string[] = [`d="${d}"`];
    const brush = fill && !dc.brush.none ? dc.brush : undefined;
    if (brush?.pattern) attrs.push(`fill="url(#${this.patternId(brush.pattern)})"`);
    else attrs.push(`fill="${brush ? brush.color : 'none'}"`);
    if (brush) attrs.push(`fill-rule="${dc.fillRule}"`);
    const pen = stroke && !dc.pen.none ? dc.pen : undefined;
    if (pen) {
      const w = pen.cosmetic || pen.width === 0 ? 1 : pen.width * this.scale();
      attrs.push(
        `stroke="${pen.color}"`,
        `stroke-width="${n(w)}"`,
        `stroke-linecap="${pen.cap}"`,
        `stroke-linejoin="${pen.join}"`,
      );
      if (pen.join === 'miter') attrs.push(`stroke-miterlimit="${dc.miterLimit}"`);
      if (pen.dash) attrs.push(`stroke-dasharray="${pen.dash.map((x) => n(x * w)).join(' ')}"`);
    }
    if (!brush && !pen) return;
    if (dc.clip) attrs.push(`clip-path="url(#${dc.clip})"`);
    this.out.push(`<path ${attrs.join(' ')}/>`);
  }

  /** Replace or intersect the clip with a device-space path. */
  private setClip(d: string, mode: number): void {
    if (!d) return;
    const id = `${this.idPrefix}c${this.nextId++}`;
    const parent = mode === RGN_AND && this.dc.clip ? ` clip-path="url(#${this.dc.clip})"` : '';
    this.defs.push(`<clipPath id="${id}"${parent}><path d="${d}" clip-rule="${this.dc.fillRule}"/></clipPath>`);
    this.dc.clip = id;
  }

  private patternId(p: { url: string; w: number; h: number }): string {
    const id = `${this.idPrefix}p${this.nextId++}`;
    this.defs.push(
      `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${p.w}" height="${p.h}">` +
        `<image href="${p.url}" width="${p.w}" height="${p.h}"/></pattern>`,
    );
    return id;
  }

  /** Draw the source rectangle of a bitmap into the logical destination rectangle. */
  private emitImage(
    url: { url: string; w: number; h: number } | undefined,
    x: number,
    y: number,
    w: number,
    h: number,
    sx: number,
    sy: number,
    sw: number,
    sh: number,
  ): void {
    if (!url || !w || !h || !sw || !sh) return;
    // Map the destination box's corners; images under a rotating world
    // transform are drawn into their mapped bounding box.
    const [x0, y0] = this.map(x, y);
    const [x1, y1] = this.map(x + w, y + h);
    const dw = x1 - x0;
    const dh = y1 - y0;
    // Show only the source rectangle: scale the whole bitmap, clip to the box.
    const kx = dw / sw;
    const ky = dh / sh;
    const id = `${this.idPrefix}i${this.nextId++}`;
    this.defs.push(
      `<clipPath id="${id}"><rect x="${n(Math.min(x0, x1))}" y="${n(Math.min(y0, y1))}" ` +
        `width="${n(Math.abs(dw))}" height="${n(Math.abs(dh))}"/></clipPath>`,
    );
    const ix = x0 - sx * kx;
    const iy = y0 - sy * ky;
    const flip = kx < 0 || ky < 0 ? ` transform="scale(${kx < 0 ? -1 : 1} ${ky < 0 ? -1 : 1})"` : '';
    const sign = (v: number, k: number) => (k < 0 ? -v : v);
    const clip = this.dc.clip ? `<g clip-path="url(#${this.dc.clip})">` : '<g>';
    this.out.push(
      `${clip}<g clip-path="url(#${id})"><image href="${url.url}" preserveAspectRatio="none" ` +
        `x="${n(sign(ix, kx) - (kx < 0 ? Math.abs(url.w * kx) : 0))}" ` +
        `y="${n(sign(iy, ky) - (ky < 0 ? Math.abs(url.h * ky) : 0))}" ` +
        `width="${n(Math.abs(url.w * kx))}" height="${n(Math.abs(url.h * ky))}"${flip}/></g></g>`,
    );
  }

  private text(off: number, size: number): void {
    const dv = this.dv;
    const base = off + 36; // EmrText follows rclBounds, iGraphicsMode, exScale, eyScale
    const refX = dv.getInt32(base, true);
    const refY = dv.getInt32(base + 4, true);
    const nChars = dv.getUint32(base + 8, true);
    const offString = dv.getUint32(base + 12, true);
    const offDx = dv.getUint32(base + 36, true);
    if (!nChars || offString + nChars * 2 > size) return;
    let str = '';
    for (let i = 0; i < nChars; i++) str += String.fromCharCode(dv.getUint16(off + offString + i * 2, true));
    if (!str.trim()) return;

    const dc = this.dc;
    const font = dc.font;
    const scale = this.scale();
    // A negative height is the character height; a positive one includes the
    // internal leading, which is roughly a fifth of the cell.
    const em = Math.abs(font.height || 12) * (font.height > 0 ? 0.8 : 1) * scale;
    const [x, y] = this.map(refX, refY);
    const align = dc.textAlign;
    const hAlign = align & 6;
    const anchor = hAlign === 6 ? 'middle' : hAlign === 2 ? 'end' : 'start';
    const vAlign = align & 24;
    const baseline = vAlign === 24 ? 'alphabetic' : vAlign === 8 ? 'text-after-edge' : 'text-before-edge';

    const attrs: string[] = [
      `x="${n(x)}"`,
      `y="${n(y)}"`,
      `font-family="${escapeXml(font.face)}, sans-serif"`,
      `font-size="${n(em)}"`,
      `fill="${dc.textColor}"`,
      `text-anchor="${anchor}"`,
      `dominant-baseline="${baseline}"`,
    ];
    if (font.weight >= 600) attrs.push('font-weight="bold"');
    if (font.italic) attrs.push('font-style="italic"');
    const deco = [font.underline && 'underline', font.strike && 'line-through'].filter(Boolean);
    if (deco.length) attrs.push(`text-decoration="${deco.join(' ')}"`);

    // Per-character advances pin each glyph where GDI put it, so a substitute
    // font cannot make the run drift. Only usable when the run is not rotated.
    if (offDx && !font.escapement && anchor === 'start' && offDx + nChars * 4 <= size) {
      const xs: string[] = [];
      let cx = refX;
      for (let i = 0; i < nChars; i++) {
        xs.push(n(this.map(cx, refY)[0]));
        cx += dv.getInt32(off + offDx + i * 4, true);
      }
      attrs[0] = `x="${xs.join(' ')}"`;
    }
    if (font.escapement) attrs.push(`transform="rotate(${n(-font.escapement / 10)} ${n(x)} ${n(y)})"`);
    if (dc.clip) attrs.push(`clip-path="url(#${dc.clip})"`);
    this.out.push(`<text ${attrs.join(' ')} xml:space="preserve">${escapeXml(str)}</text>`);
  }

  // ─── Object readers ─────────────────────────────────────────────────────────

  private makePen(style: number, width: number, color: string, extended: boolean): Pen {
    const kind = style & 0x0f;
    const capBits = style & 0x0f00;
    const joinBits = style & 0xf000;
    const geometric = extended ? (style & 0x00010000) !== 0 : true;
    const pen: Pen = {
      kind: 'pen',
      none: kind === 5,
      width: Math.abs(width),
      cosmetic: !geometric,
      color,
      cap: capBits === 0x100 ? 'square' : capBits === 0x200 ? 'butt' : 'round',
      join: joinBits === 0x1000 ? 'bevel' : joinBits === 0x2000 ? 'miter' : 'round',
    };
    const dashes: Record<number, number[]> = { 1: [3, 1], 2: [1, 1], 3: [3, 1, 1, 1], 4: [3, 1, 1, 1, 1, 1] };
    if (dashes[kind]) pen.dash = dashes[kind];
    return pen;
  }

  private readFont(o: number): Font {
    const dv = this.dv;
    let face = '';
    for (let i = 0; i < 32; i++) {
      const c = dv.getUint16(o + 28 + i * 2, true);
      if (!c) break;
      face += String.fromCharCode(c);
    }
    return {
      kind: 'font',
      height: dv.getInt32(o, true),
      escapement: dv.getInt32(o + 8, true),
      weight: dv.getInt32(o + 16, true),
      italic: dv.getUint8(o + 20) !== 0,
      underline: dv.getUint8(o + 21) !== 0,
      strike: dv.getUint8(o + 22) !== 0,
      face: face || 'Arial',
    };
  }

  private colorRef(o: number): string {
    const r = this.dv.getUint8(o);
    const g = this.dv.getUint8(o + 1);
    const b = this.dv.getUint8(o + 2);
    return `#${((1 << 24) | (r << 16) | (g << 8) | b).toString(16).slice(1)}`;
  }

  /**
   * A device-independent bitmap embedded in a record, as a data URL. The DIB is
   * wrapped in a BMP file header, which every browser decodes; a DIB whose
   * compression is BI_JPEG/BI_PNG already holds a complete JPEG/PNG.
   */
  private dibUrl(
    recOff: number,
    offBmi: number,
    cbBmi: number,
    offBits: number,
    cbBits: number,
  ): { url: string; w: number; h: number } | undefined {
    if (!cbBmi || !cbBits) return undefined;
    const bmi = this.bytes.subarray(recOff + offBmi, recOff + offBmi + cbBmi);
    const bits = this.bytes.subarray(recOff + offBits, recOff + offBits + cbBits);
    const dv = new DataView(bmi.buffer, bmi.byteOffset, bmi.byteLength);
    const w = Math.abs(dv.getInt32(4, true));
    const h = Math.abs(dv.getInt32(8, true));
    const compression = cbBmi >= 20 ? dv.getUint32(16, true) : 0;
    if (!w || !h) return undefined;
    if (compression === 4 || compression === 5) {
      const mime = compression === 4 ? 'image/jpeg' : 'image/png';
      return { url: `data:${mime};base64,${toBase64(bits)}`, w, h };
    }
    const file = new Uint8Array(14 + bmi.length + bits.length);
    const fdv = new DataView(file.buffer);
    fdv.setUint8(0, 0x42); // 'B'
    fdv.setUint8(1, 0x4d); // 'M'
    fdv.setUint32(2, file.length, true);
    fdv.setUint32(10, 14 + bmi.length, true);
    file.set(bmi, 14);
    file.set(bits, 14 + bmi.length);
    return { url: `data:image/bmp;base64,${toBase64(file)}`, w, h };
  }
}

/** SVG cubic segments for points [c1, c2, end, c1, c2, end, ...]. */
function bezierSegs(pts: [number, number][]): string {
  let d = '';
  for (let i = 0; i + 2 < pts.length; i += 3) {
    d += `C${pt(pts[i]!)} ${pt(pts[i + 1]!)} ${pt(pts[i + 2]!)}`;
  }
  return d;
}

/**
 * Bézier approximation of an elliptical arc in logical space: the start point
 * followed by (c1, c2, end) triples, at most a quarter turn each.
 */
function ellipseBeziers(
  cx: number,
  cy: number,
  rx: number,
  ry: number,
  start: number,
  sweep: number,
): [number, number][] {
  const segs = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2) - 1e-9));
  const step = sweep / segs;
  const k = (4 / 3) * Math.tan(step / 4);
  const at = (a: number): [number, number] => [cx + rx * Math.cos(a), cy + ry * Math.sin(a)];
  const out: [number, number][] = [at(start)];
  for (let i = 0; i < segs; i++) {
    const a0 = start + i * step;
    const a1 = a0 + step;
    out.push(
      [cx + rx * (Math.cos(a0) - k * Math.sin(a0)), cy + ry * (Math.sin(a0) + k * Math.cos(a0))],
      [cx + rx * (Math.cos(a1) + k * Math.sin(a1)), cy + ry * (Math.sin(a1) - k * Math.cos(a1))],
      at(a1),
    );
  }
  return out;
}

function n(v: number): string {
  return Number.isInteger(v) ? String(v) : v.toFixed(2).replace(/\.?0+$/, '');
}

function pt(p: [number, number]): string {
  return `${n(p[0])},${n(p[1])}`;
}

function escapeXml(s: string): string {
  return s.replace(/[<>&"]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;' })[c]!);
}

function toBase64(bytes: Uint8Array): string {
  let bin = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(bin);
}
