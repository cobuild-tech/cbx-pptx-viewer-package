/**
 * Run formatting: the toolbar's view of a run, and how it maps to `<a:rPr>`.
 *
 * Kept separate from the DOM so both the toolbar (which reads the current
 * selection's formatting) and the XML writer (which applies it) share one
 * definition of what a "format" is.
 */
import { child, createElement, setAttr, type XmlNode } from '../../oxml/xml.js';

/** A formatting override the toolbar can apply to a stretch of text. */
export interface RunFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Font size in points. */
  sizePt?: number;
  /** sRGB hex, no leading '#'. */
  colorHex?: string;
  /** Typeface name. */
  font?: string;
}

/** True if the format carries no actual overrides. */
export function isEmptyFormat(f: RunFormat | undefined): boolean {
  return !f || Object.values(f).every((v) => v === undefined);
}

/** Merge `over` on top of `base` (later wins, undefined does not clear). */
export function mergeFormat(base: RunFormat | undefined, over: RunFormat | undefined): RunFormat {
  const out: RunFormat = { ...base };
  if (over) {
    for (const [k, v] of Object.entries(over)) {
      if (v !== undefined) (out as Record<string, unknown>)[k] = v;
    }
  }
  return out;
}

/**
 * The DrawingML child-element order for `<a:rPr>` (CT_TextCharacterProperties).
 * A fill has to sit in the right slot or PowerPoint rejects the part, so new
 * children are inserted by rank rather than appended.
 */
const RPR_ORDER = [
  'ln',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'effectLst',
  'effectDag',
  'highlight',
  'uLnTx',
  'uLn',
  'uFillTx',
  'uFill',
  'latin',
  'ea',
  'cs',
  'sym',
  'hlinkClick',
  'hlinkMouseOver',
  'rtl',
  'extLst',
];

function rank(local: string): number {
  const i = RPR_ORDER.indexOf(local);
  return i === -1 ? RPR_ORDER.length : i;
}

function localOf(name: string): string {
  const i = name.indexOf(':');
  return i === -1 ? name : name.slice(i + 1);
}

/** Insert `node` into `parent` at the position its schema rank requires. */
function insertOrdered(parent: XmlNode, node: XmlNode): void {
  const r = rank(localOf(node.name));
  let at = parent.children.length;
  for (let i = 0; i < parent.children.length; i++) {
    const c = parent.children[i];
    if (c && rank(localOf(c.name)) > r) {
      at = i;
      break;
    }
  }
  parent.children.splice(at, 0, node);
}

const FILL_TAGS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);

/**
 * Apply a {@link RunFormat} to an `<a:rPr>` node in place. `prefix` is the
 * namespace prefix in use for the DrawingML namespace (e.g. "a"), taken from
 * surrounding nodes so we never guess.
 */
export function applyFormat(rPr: XmlNode, format: RunFormat, prefix: string): void {
  const q = (n: string) => (prefix ? `${prefix}:${n}` : n);

  if (format.bold !== undefined) setAttr(rPr, 'b', format.bold ? '1' : '0');
  if (format.italic !== undefined) setAttr(rPr, 'i', format.italic ? '1' : '0');
  if (format.underline !== undefined) setAttr(rPr, 'u', format.underline ? 'sng' : 'none');
  if (format.strike !== undefined) setAttr(rPr, 'strike', format.strike ? 'sngStrike' : 'noStrike');
  if (format.sizePt !== undefined) setAttr(rPr, 'sz', String(Math.round(format.sizePt * 100)));

  if (format.colorHex !== undefined) {
    // Replace whatever fill is there — a run has at most one.
    rPr.children = rPr.children.filter((c) => !FILL_TAGS.has(localOf(c.name)));
    insertOrdered(
      rPr,
      createElement(q('solidFill'), {}, [
        createElement(q('srgbClr'), { val: format.colorHex.replace(/^#/, '').toUpperCase() }),
      ]),
    );
  }

  if (format.font !== undefined) {
    // PowerPoint tracks latin/ea/cs typefaces separately; set them together so
    // the run doesn't render in a mix of families.
    for (const tag of ['latin', 'ea', 'cs']) {
      const existing = child(rPr, tag);
      if (existing) setAttr(existing, 'typeface', format.font);
      else insertOrdered(rPr, createElement(q(tag), { typeface: format.font }));
    }
  }
}

/** Read the formatting explicitly set on an `<a:rPr>` (no inheritance). */
export function readFormat(rPr: XmlNode | undefined): RunFormat {
  if (!rPr) return {};
  const out: RunFormat = {};
  const b = rPr.attrs['b'];
  if (b !== undefined) out.bold = b === '1' || b === 'true';
  const i = rPr.attrs['i'];
  if (i !== undefined) out.italic = i === '1' || i === 'true';
  const u = rPr.attrs['u'];
  if (u !== undefined) out.underline = u !== 'none';
  const strike = rPr.attrs['strike'];
  if (strike !== undefined) out.strike = strike !== 'noStrike';
  const sz = rPr.attrs['sz'];
  if (sz !== undefined) {
    const n = Number(sz);
    if (Number.isFinite(n)) out.sizePt = n / 100;
  }
  const srgb = child(child(rPr, 'solidFill'), 'srgbClr');
  const val = srgb?.attrs['val'];
  if (val) out.colorHex = val.toUpperCase();
  const latin = child(rPr, 'latin');
  const typeface = latin?.attrs['typeface'];
  if (typeface) out.font = typeface;
  return out;
}

/** Remove a hyperlink from an rPr (used when a linked run is re-typed). */
export function stripHyperlink(rPr: XmlNode): void {
  rPr.children = rPr.children.filter((c) => localOf(c.name) !== 'hlinkClick');
}
