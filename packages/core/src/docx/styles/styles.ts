/**
 * Style resolution for WordprocessingML.
 *
 * Parses word/styles.xml and builds a resolved style map so paragraph and run
 * parsers can look up inherited properties without traversing the XML again.
 * Each style inherits from its <w:basedOn> parent, ultimately rooting at Normal.
 */
import { child, children, attr, attrNum, attrBool, type XmlNode } from '../../oxml/xml.js';
import { halfPtToPt, twipsToPx, borderSzToPx } from '../units.js';
import type { TextAlign } from '../model.js';

export interface ParaBorderSide {
  colorHex: string;
  widthPx: number;
  type: string;
  spacePt: number;
}

export interface ResolvedRunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  sizePt?: number;
  /** Hex color (no #). */
  colorHex?: string;
  fontAscii?: string;
  fontHAnsi?: string;
  caps?: 'all' | 'small';
  /** Highlight color name (yellow, green, cyan, etc.). */
  highlight?: string;
  /** Character spacing in pt (from <w:spacing> in rPr, which is in 20ths of a pt). */
  letterSpacingPt?: number;
}

export interface ResolvedParaStyle {
  align?: TextAlign;
  indentLeftPx?: number;
  indentFirstLinePx?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  lineSpacingPct?: number;
  lineSpacingPt?: number;
  /** numId + ilvl from a linked numbering definition. */
  numId?: number;
  ilvl?: number;
  outlineLevel?: number;
  shadingHex?: string;
  indentRightPx?: number;
  contextualSpacing?: boolean;
  pBdr?: Partial<Record<'top' | 'bottom' | 'left' | 'right', ParaBorderSide>>;
}

export interface ResolvedStyle {
  id: string;
  name: string;
  para: ResolvedParaStyle;
  run: ResolvedRunStyle;
}

export class StyleMap {
  private styles = new Map<string, ResolvedStyle>();

  static parse(stylesXml: XmlNode | undefined): StyleMap {
    const map = new StyleMap();
    if (!stylesXml) return map;

    // Parse docDefaults for document-level font/size/para defaults.
    const docDefaultsEl = children(stylesXml, 'docDefaults')[0];
    const rPrDefaultEl = child(child(docDefaultsEl, 'rPrDefault'), 'rPr');
    const pPrDefaultEl = child(child(docDefaultsEl, 'pPrDefault'), 'pPr');
    const docDefaultRun: ResolvedRunStyle = {};
    const docDefaultPara: ResolvedParaStyle = {};
    if (rPrDefaultEl) mergeRunProps(docDefaultRun, rPrDefaultEl);
    if (pPrDefaultEl) mergeParaProps(docDefaultPara, pPrDefaultEl);

    // First pass: collect raw styles.
    const rawStyles = new Map<string, { id: string; name: string; basedOn?: string; pPr?: XmlNode; rPr?: XmlNode }>();
    for (const styleEl of children(stylesXml, 'style')) {
      const id = attr(styleEl, 'w:styleId') ?? attr(styleEl, 'styleId');
      if (!id) continue;
      const basedOn = attr(child(styleEl, 'basedOn'), 'w:val') ?? attr(child(styleEl, 'basedOn'), 'val');
      const nameEl = child(styleEl, 'name');
      const displayName = (attr(nameEl, 'w:val') ?? attr(nameEl, 'val') ?? id).toLowerCase();
      rawStyles.set(id, {
        id,
        name: displayName,
        basedOn,
        pPr: child(styleEl, 'pPr'),
        rPr: child(styleEl, 'rPr'),
      });
    }

    // Second pass: resolve inheritance (DFS with memoisation).
    const resolved = new Map<string, ResolvedStyle>();
    function resolve(id: string, visited = new Set<string>()): ResolvedStyle {
      const memo = resolved.get(id);
      if (memo) return memo;

      const raw = rawStyles.get(id);
      const base: ResolvedStyle = { id, name: raw?.name ?? id, para: { ...docDefaultPara }, run: { ...docDefaultRun } };

      if (raw?.basedOn && !visited.has(raw.basedOn)) {
        visited.add(id);
        const parent = resolve(raw.basedOn, visited);
        base.para = { ...parent.para };
        base.run = { ...parent.run };
      }

      if (raw?.pPr) mergeParaProps(base.para, raw.pPr);
      if (raw?.rPr) mergeRunProps(base.run, raw.rPr);

      resolved.set(id, base);
      return base;
    }

    for (const id of rawStyles.keys()) {
      map.styles.set(id, resolve(id));
    }
    return map;
  }

  get(styleId: string): ResolvedStyle {
    return this.styles.get(styleId) ?? { id: styleId, name: styleId, para: {}, run: {} };
  }

  /** Return a style by its canonical name (case-insensitive). */
  getByName(name: string): ResolvedStyle | undefined {
    const lower = name.toLowerCase().replace(/\s/g, '');
    for (const s of this.styles.values()) {
      if (s.name.toLowerCase().replace(/\s/g, '') === lower) return s;
    }
    return undefined;
  }
}

export function mergeRunProps(out: ResolvedRunStyle, rPr: XmlNode): void {
  if (child(rPr, 'b')) {
    const val = attr(child(rPr, 'b'), 'w:val') ?? attr(child(rPr, 'b'), 'val');
    out.bold = val === undefined || (val !== '0' && val !== 'false');
  }
  if (child(rPr, 'i')) {
    const val = attr(child(rPr, 'i'), 'w:val') ?? attr(child(rPr, 'i'), 'val');
    out.italic = val === undefined || (val !== '0' && val !== 'false');
  }
  const uEl = child(rPr, 'u');
  if (uEl) {
    const val = attr(uEl, 'w:val') ?? attr(uEl, 'val');
    out.underline = val !== 'none' && val !== undefined ? true : val !== 'none';
    if (val === 'none') out.underline = false;
    else if (val !== undefined) out.underline = true;
  }
  if (child(rPr, 'strike')) {
    const val = attr(child(rPr, 'strike'), 'w:val') ?? attr(child(rPr, 'strike'), 'val');
    out.strike = val === undefined || (val !== '0' && val !== 'false');
  }
  const szEl = child(rPr, 'sz');
  if (szEl) {
    const sz = attrNum(szEl, 'w:val') ?? attrNum(szEl, 'val');
    if (sz !== undefined) out.sizePt = halfPtToPt(sz);
  }
  const colorEl = child(rPr, 'color');
  if (colorEl) {
    const hex = attr(colorEl, 'w:val') ?? attr(colorEl, 'val');
    const themeColor = attr(colorEl, 'w:themeColor') ?? attr(colorEl, 'themeColor');
    if (themeColor) {
      const tc = themeColor.toLowerCase();
      // Dark/text theme slots → always black text.
      if (tc === 'dark1' || tc === 'dk1' || tc === 'text1' || tc === 'tx1') {
        out.colorHex = '000000';
      } else if (tc === 'dark2' || tc === 'dk2' || tc === 'text2' || tc === 'tx2') {
        out.colorHex = '000000';
      } else if (tc === 'light1' || tc === 'lt1' || tc === 'background1' || tc === 'bg1') {
        // Light/background slot — leave unset so text inherits the page default.
      } else if (tc === 'light2' || tc === 'lt2' || tc === 'background2' || tc === 'bg2') {
        // Light/background slot — leave unset so text inherits the page default.
      } else {
        // Accent, hyperlink, etc. — trust the cached w:val fallback.
        if (hex && hex !== 'auto') out.colorHex = hex.toUpperCase();
      }
    } else if (hex && hex !== 'auto') {
      out.colorHex = hex.toUpperCase();
    }
  }
  const fontsEl = child(rPr, 'rFonts');
  if (fontsEl) {
    const ascii = attr(fontsEl, 'w:ascii') ?? attr(fontsEl, 'ascii');
    const hAnsi = attr(fontsEl, 'w:hAnsi') ?? attr(fontsEl, 'hAnsi');
    if (ascii) out.fontAscii = ascii;
    if (hAnsi) out.fontHAnsi = hAnsi;
  }
  if (child(rPr, 'caps')) {
    const val = attr(child(rPr, 'caps'), 'w:val') ?? attr(child(rPr, 'caps'), 'val');
    if (val === undefined || (val !== '0' && val !== 'false')) out.caps = 'all';
    else out.caps = undefined;
  }
  if (child(rPr, 'smallCaps')) {
    const val = attr(child(rPr, 'smallCaps'), 'w:val') ?? attr(child(rPr, 'smallCaps'), 'val');
    if (val === undefined || (val !== '0' && val !== 'false')) out.caps = 'small';
    else out.caps = undefined;
  }
  const hlEl = child(rPr, 'highlight');
  if (hlEl) {
    const hl = attr(hlEl, 'w:val') ?? attr(hlEl, 'val');
    if (hl && hl !== 'none') out.highlight = hl;
  }
  // Character spacing (w:spacing in rPr is in twentieths of a point)
  const runSpacingEl = child(rPr, 'spacing');
  if (runSpacingEl) {
    const sp = attrNum(runSpacingEl, 'w:val') ?? attrNum(runSpacingEl, 'val');
    if (sp !== undefined) out.letterSpacingPt = sp / 20;
  }
}

export function mergeParaProps(out: ResolvedParaStyle, pPr: XmlNode): void {
  const jcEl = child(pPr, 'jc');
  if (jcEl) {
    const jc = attr(jcEl, 'w:val') ?? attr(jcEl, 'val');
    out.align = jcToAlign(jc);
  }
  const indEl = child(pPr, 'ind');
  if (indEl) {
    const left = attrNum(indEl, 'w:left') ?? attrNum(indEl, 'left');
    const firstLine = attrNum(indEl, 'w:firstLine') ?? attrNum(indEl, 'firstLine');
    const hanging = attrNum(indEl, 'w:hanging') ?? attrNum(indEl, 'hanging');
    const right = attrNum(indEl, 'w:right') ?? attrNum(indEl, 'right');
    if (left !== undefined) out.indentLeftPx = twipsToPx(left);
    if (firstLine !== undefined) out.indentFirstLinePx = twipsToPx(firstLine);
    if (hanging !== undefined) out.indentFirstLinePx = -twipsToPx(hanging);
    if (right !== undefined) out.indentRightPx = twipsToPx(right);
  }
  const spcEl = child(pPr, 'spacing');
  if (spcEl) {
    const before = attrNum(spcEl, 'w:before') ?? attrNum(spcEl, 'before');
    const after = attrNum(spcEl, 'w:after') ?? attrNum(spcEl, 'after');
    const line = attrNum(spcEl, 'w:line') ?? attrNum(spcEl, 'line');
    const lineRule = attr(spcEl, 'w:lineRule') ?? attr(spcEl, 'lineRule');
    if (before !== undefined) out.spaceBeforePt = twipsToPx(before) * 0.75; // px→pt approx
    if (after !== undefined) out.spaceAfterPt = twipsToPx(after) * 0.75;
    if (line !== undefined) {
      if (lineRule === 'exact' || lineRule === 'atLeast') {
        out.lineSpacingPt = twipsToPx(line) * 0.75;
      } else {
        // auto: line is in 240ths; 240 = single spacing
        out.lineSpacingPct = line / 240;
      }
    }
  }
  const numPrEl = child(pPr, 'numPr');
  if (numPrEl) {
    const numId = attrNum(child(numPrEl, 'numId'), 'w:val') ?? attrNum(child(numPrEl, 'numId'), 'val');
    const ilvl = attrNum(child(numPrEl, 'ilvl'), 'w:val') ?? attrNum(child(numPrEl, 'ilvl'), 'val');
    if (numId !== undefined) out.numId = numId;
    if (ilvl !== undefined) out.ilvl = ilvl;
  }
  // Contextual spacing
  const ctxSpcEl = child(pPr, 'contextualSpacing');
  if (ctxSpcEl) {
    const val = attr(ctxSpcEl, 'w:val') ?? attr(ctxSpcEl, 'val');
    out.contextualSpacing = val === undefined || (val !== '0' && val !== 'false');
  }

  // Paragraph borders
  const pBdrEl = child(pPr, 'pBdr');
  if (pBdrEl) {
    const bdr: Partial<Record<'top' | 'bottom' | 'left' | 'right', ParaBorderSide>> = {};
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const sideEl = child(pBdrEl, side);
      if (!sideEl) continue;
      const val = attr(sideEl, 'w:val') ?? attr(sideEl, 'val');
      if (!val || val === 'none' || val === 'nil') continue;
      const sz = attrNum(sideEl, 'w:sz') ?? attrNum(sideEl, 'sz') ?? 4;
      const colorHex = attr(sideEl, 'w:color') ?? attr(sideEl, 'color');
      const space = attrNum(sideEl, 'w:space') ?? attrNum(sideEl, 'space') ?? 0;
      const hex = colorHex && colorHex !== 'auto' ? colorHex.toUpperCase() : '000000';
      bdr[side] = { colorHex: hex, widthPx: borderSzToPx(sz), type: val, spacePt: space };
    }
    if (Object.keys(bdr).length > 0) out.pBdr = bdr;
  }

  const outlineEl = child(pPr, 'outlineLvl');
  if (outlineEl) {
    const lvl = attrNum(outlineEl, 'w:val') ?? attrNum(outlineEl, 'val');
    if (lvl !== undefined) out.outlineLevel = lvl;
  }
  const shdEl = child(pPr, 'shd');
  if (shdEl) {
    const fill = attr(shdEl, 'w:fill') ?? attr(shdEl, 'fill');
    const val = attr(shdEl, 'w:val') ?? attr(shdEl, 'val');
    if (fill && fill !== 'auto' && val !== 'clear' && val !== 'nil') {
      out.shadingHex = fill.toUpperCase();
    }
  }
}

function jcToAlign(jc: string | undefined): TextAlign | undefined {
  switch (jc) {
    case 'center': return 'ctr';
    case 'right': return 'r';
    case 'both':
    case 'distribute': return 'just';
    case 'left': return 'l';
    default: return undefined;
  }
}
