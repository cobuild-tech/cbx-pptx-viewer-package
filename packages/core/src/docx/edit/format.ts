/**
 * Run formatting for DOCX: the toolbar's view of a run, and how it maps to
 * `<w:rPr>`.
 *
 * The shape mirrors pptx/edit/format.ts, but the encoding does not.
 * WordprocessingML puts run properties in child *elements* with a `w:val`
 * attribute (`<w:b/>`, `<w:sz w:val="24"/>`) rather than attributes on the
 * property element, sizes are in **half-points**, and `<w:rPr>` has a strict
 * child order that Word enforces.
 */
import { child, createElement, setAttr, localName, type XmlNode } from '../../oxml/xml.js';

/** A formatting override the toolbar can apply to a stretch of text. */
export interface DocxRunFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Font size in points (stored as half-points). */
  sizePt?: number;
  /** sRGB hex, no leading '#'. */
  colorHex?: string;
  /** Typeface name. */
  font?: string;
}

export function isEmptyFormat(f: DocxRunFormat | undefined): boolean {
  return !f || Object.values(f).every((v) => v === undefined);
}

/** Merge `over` on top of `base` (later wins; undefined does not clear). */
export function mergeFormat(
  base: DocxRunFormat | undefined,
  over: DocxRunFormat | undefined,
): DocxRunFormat {
  const out: DocxRunFormat = { ...base };
  if (over) {
    for (const [k, v] of Object.entries(over)) {
      if (v !== undefined) (out as Record<string, unknown>)[k] = v;
    }
  }
  return out;
}

/**
 * `<w:rPr>` child order (CT_RPr). Word rejects a run whose properties are out
 * of order, so new children are inserted by rank rather than appended.
 */
const RPR_ORDER = [
  'rStyle',
  'rFonts',
  'b',
  'bCs',
  'i',
  'iCs',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'noProof',
  'snapToGrid',
  'vanish',
  'webHidden',
  'color',
  'spacing',
  'w',
  'kern',
  'position',
  'sz',
  'szCs',
  'highlight',
  'u',
  'effect',
  'bdr',
  'shd',
  'fitText',
  'vertAlign',
  'rtl',
  'cs',
  'em',
  'lang',
  'eastAsianLayout',
  'specVanish',
  'oMath',
];

function rank(local: string): number {
  const i = RPR_ORDER.indexOf(local);
  return i === -1 ? RPR_ORDER.length : i;
}

/** Insert `node` into `parent` at the position its schema rank requires. */
function insertOrdered(parent: XmlNode, node: XmlNode): void {
  const r = rank(localName(node.name));
  let at = parent.children.length;
  for (let i = 0; i < parent.children.length; i++) {
    const c = parent.children[i];
    if (c && rank(localName(c.name)) > r) {
      at = i;
      break;
    }
  }
  parent.children.splice(at, 0, node);
}

/** Remove every child with this local name. */
function dropChildren(parent: XmlNode, local: string): void {
  parent.children = parent.children.filter((c) => localName(c.name) !== local);
}

/**
 * Set (or clear) a boolean toggle property. WordprocessingML toggles are
 * present-means-on, but an explicit `w:val="0"` is needed to switch one OFF
 * when a style would otherwise turn it on.
 */
function setToggle(rPr: XmlNode, local: string, on: boolean, prefix: string): void {
  dropChildren(rPr, local);
  const q = prefix ? `${prefix}:${local}` : local;
  insertOrdered(rPr, on ? createElement(q) : createElement(q, { [`${prefix || 'w'}:val`]: '0' }));
}

function setVal(rPr: XmlNode, local: string, value: string, prefix: string): void {
  const q = prefix ? `${prefix}:${local}` : local;
  const valAttr = `${prefix || 'w'}:val`;
  const existing = child(rPr, local);
  if (existing) setAttr(existing, valAttr, value);
  else insertOrdered(rPr, createElement(q, { [valAttr]: value }));
}

/**
 * Apply a {@link DocxRunFormat} to a `<w:rPr>` in place. `prefix` is the
 * WordprocessingML namespace prefix in use (normally "w"), taken from
 * surrounding nodes rather than assumed.
 */
export function applyFormat(rPr: XmlNode, format: DocxRunFormat, prefix: string): void {
  if (format.bold !== undefined) setToggle(rPr, 'b', format.bold, prefix);
  if (format.italic !== undefined) setToggle(rPr, 'i', format.italic, prefix);
  if (format.strike !== undefined) setToggle(rPr, 'strike', format.strike, prefix);

  // Underline is not a toggle — it carries a style value.
  if (format.underline !== undefined) {
    setVal(rPr, 'u', format.underline ? 'single' : 'none', prefix);
  }

  if (format.sizePt !== undefined) {
    // Half-points, and szCs (complex script) tracks it so the run doesn't
    // render at two different sizes.
    const half = String(Math.round(format.sizePt * 2));
    setVal(rPr, 'sz', half, prefix);
    setVal(rPr, 'szCs', half, prefix);
  }

  if (format.colorHex !== undefined) {
    setVal(rPr, 'color', format.colorHex.replace(/^#/, '').toUpperCase(), prefix);
  }

  if (format.font !== undefined) {
    // <w:rFonts> carries the ascii/hAnsi/cs/eastAsia faces as attributes.
    const q = prefix ? `${prefix}:rFonts` : 'rFonts';
    const existing = child(rPr, 'rFonts');
    const node = existing ?? createElement(q);
    for (const which of ['ascii', 'hAnsi', 'cs', 'eastAsia']) {
      setAttr(node, `${prefix || 'w'}:${which}`, format.font);
    }
    if (!existing) insertOrdered(rPr, node);
  }
}

/** Read the formatting explicitly set on a `<w:rPr>` (no style inheritance). */
export function readFormat(rPr: XmlNode | undefined): DocxRunFormat {
  if (!rPr) return {};
  const out: DocxRunFormat = {};

  const toggle = (local: string): boolean | undefined => {
    const el = child(rPr, local);
    if (!el) return undefined;
    const val = el.attrs['w:val'] ?? el.attrs['val'];
    return val === undefined ? true : val !== '0' && val !== 'false';
  };

  const b = toggle('b');
  if (b !== undefined) out.bold = b;
  const i = toggle('i');
  if (i !== undefined) out.italic = i;
  const strike = toggle('strike');
  if (strike !== undefined) out.strike = strike;

  const u = child(rPr, 'u');
  if (u) out.underline = (u.attrs['w:val'] ?? u.attrs['val']) !== 'none';

  const sz = child(rPr, 'sz');
  const szVal = sz && (sz.attrs['w:val'] ?? sz.attrs['val']);
  if (szVal) {
    const n = Number(szVal);
    if (Number.isFinite(n)) out.sizePt = n / 2;
  }

  const color = child(rPr, 'color');
  const colorVal = color && (color.attrs['w:val'] ?? color.attrs['val']);
  if (colorVal && colorVal !== 'auto') out.colorHex = colorVal.toUpperCase();

  const fonts = child(rPr, 'rFonts');
  const face = fonts && (fonts.attrs['w:ascii'] ?? fonts.attrs['ascii']);
  if (face) out.font = face;

  return out;
}
