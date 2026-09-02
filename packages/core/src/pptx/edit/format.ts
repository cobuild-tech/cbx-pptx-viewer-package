/**
 * DrawingML encoding of a {@link RunFormat}: how a format maps to `<a:rPr>`.
 *
 * The format *value* is shared with the other formats (oxml/edit/format.ts);
 * only the encoding is PowerPoint-specific — attributes on `<a:rPr>` plus a
 * `<a:solidFill>` child, in schema order.
 */
import {
  child,
  createElement,
  insertInOrder,
  localName,
  setAttr,
  type XmlNode,
} from '../../oxml/xml.js';
import { isEmptyFormat, mergeFormat, type RunFormat } from '../../oxml/edit/format.js';
import type { TextRun } from '../model.js';

export { isEmptyFormat, mergeFormat, type RunFormat };

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
    rPr.children = rPr.children.filter((c) => !FILL_TAGS.has(localName(c.name)));
    insertInOrder(
      rPr,
      createElement(q('solidFill'), {}, [
        createElement(q('srgbClr'), { val: format.colorHex.replace(/^#/, '').toUpperCase() }),
      ]),
      RPR_ORDER,
    );
  }

  if (format.font !== undefined) {
    // PowerPoint tracks latin/ea/cs typefaces separately; set them together so
    // the run doesn't render in a mix of families.
    for (const tag of ['latin', 'ea', 'cs']) {
      const existing = child(rPr, tag);
      if (existing) setAttr(existing, 'typeface', format.font);
      else insertInOrder(rPr, createElement(q(tag), { typeface: format.font }), RPR_ORDER);
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

/**
 * The formatting of a *resolved* model run, for the toolbar's active state.
 * Reads the model rather than `<a:rPr>` so properties inherited from the
 * layout, master and theme are reflected too.
 */
export function readRunFormat(run: object): RunFormat {
  const r = run as TextRun;
  const out: RunFormat = {};
  if (r.bold !== undefined) out.bold = r.bold;
  if (r.italic !== undefined) out.italic = r.italic;
  if (r.underline !== undefined) out.underline = r.underline;
  if (r.strike !== undefined) out.strike = r.strike;
  if (r.sizePt !== undefined) out.sizePt = r.sizePt;
  if (r.color) out.colorHex = r.color.hex;
  if (r.font) out.font = r.font;
  return out;
}

/** Remove a hyperlink from an rPr (used when a linked run is re-typed). */
export function stripHyperlink(rPr: XmlNode): void {
  rPr.children = rPr.children.filter((c) => localName(c.name) !== 'hlinkClick');
}
