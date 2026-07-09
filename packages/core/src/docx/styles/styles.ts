/**
 * Style cascade for WordprocessingML.
 *
 * Word resolves formatting through a chain:
 *   docDefaults -> (default style) -> named style (basedOn chain) -> direct props
 * for both paragraphs (pPr) and runs (rPr). This module parses styles.xml into
 * typed partial-property objects and resolves that cascade to concrete
 * {@link ParaProps} / {@link RunProps} the parsers consume.
 */
import { child, children, attr, attrNum, type XmlNode } from '../../oxml/xml.js';

/** Character-level formatting, in raw Word units where noted. */
export interface RunProps {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Font size in half-points (Word's <w:sz>). */
  sizeHalfPt?: number;
  /** Text color hex (no #), or undefined for 'auto'/inherit. */
  colorHex?: string;
  /** Highlight color hex (no #). */
  highlightHex?: string;
  font?: string;
  vertAlign?: 'super' | 'sub';
  caps?: 'all' | 'small';
}

/** Paragraph-level formatting, twips where noted. */
export interface ParaProps {
  styleId?: string;
  align?: 'l' | 'ctr' | 'r' | 'just';
  indentLeftTwip?: number;
  indentRightTwip?: number;
  /** First-line indent in twips (positive). */
  indentFirstLineTwip?: number;
  /** Hanging indent in twips (positive; pulls first line left). */
  hangingTwip?: number;
  spaceBeforeTwip?: number;
  spaceAfterTwip?: number;
  /** Raw <w:spacing w:line>; interpretation depends on lineRule. */
  line?: number;
  lineRule?: 'auto' | 'exact' | 'atLeast';
  contextualSpacing?: boolean;
  keepTogether?: boolean;
  pageBreakBefore?: boolean;
  shadingHex?: string;
  numId?: number;
  ilvl?: number;
  borders?: RawBorders;
}

export interface RawBorder {
  /** Border width in eighths of a point (<w:sz>). */
  sz?: number;
  colorHex?: string;
  val?: string;
}
export type RawBorders = Partial<Record<'top' | 'bottom' | 'left' | 'right', RawBorder>>;

interface StyleDef {
  id: string;
  type: string;
  name: string;
  basedOn?: string;
  linkId?: string;
  pPr: Partial<ParaProps>;
  rPr: Partial<RunProps>;
}

export class StyleTable {
  private readonly byId = new Map<string, StyleDef>();
  private defaultParaId?: string;
  private defaultCharId?: string;
  readonly docDefaultRun: Partial<RunProps>;
  readonly docDefaultPara: Partial<ParaProps>;

  private constructor(
    byId: Map<string, StyleDef>,
    docDefaultRun: Partial<RunProps>,
    docDefaultPara: Partial<ParaProps>,
    defaultParaId: string | undefined,
    defaultCharId: string | undefined,
  ) {
    this.byId = byId;
    this.docDefaultRun = docDefaultRun;
    this.docDefaultPara = docDefaultPara;
    this.defaultParaId = defaultParaId;
    this.defaultCharId = defaultCharId;
  }

  static parse(stylesXml: XmlNode | undefined): StyleTable {
    const byId = new Map<string, StyleDef>();
    let docRun: Partial<RunProps> = {};
    let docPara: Partial<ParaProps> = {};
    let defaultParaId: string | undefined;
    let defaultCharId: string | undefined;

    if (stylesXml) {
      const dd = child(stylesXml, 'docDefaults');
      docRun = rPrFrom(child(child(dd, 'rPrDefault'), 'rPr'));
      docPara = pPrFrom(child(child(dd, 'pPrDefault'), 'pPr'));

      for (const s of children(stylesXml, 'style')) {
        const id = attr(s, 'styleId');
        if (!id) continue;
        const type = attr(s, 'type') ?? 'paragraph';
        const def: StyleDef = {
          id,
          type,
          name: attr(child(s, 'name'), 'val') ?? id,
          basedOn: attr(child(s, 'basedOn'), 'val'),
          linkId: attr(child(s, 'link'), 'val'),
          pPr: pPrFrom(child(s, 'pPr')),
          rPr: rPrFrom(child(s, 'rPr')),
        };
        byId.set(id, def);
        if (attr(s, 'default') && bool(attr(s, 'default'))) {
          if (type === 'paragraph') defaultParaId = id;
          else if (type === 'character') defaultCharId = id;
        }
      }
    }
    return new StyleTable(byId, docRun, docPara, defaultParaId, defaultCharId);
  }

  styleName(id: string | undefined): string {
    if (!id) return 'Normal';
    return this.byId.get(id)?.name ?? id;
  }

  /** Style ids from a style up to its root via basedOn, ordered root→leaf. */
  private chain(id: string | undefined): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    let cur = id ? this.byId.get(id) : undefined;
    while (cur && !seen.has(cur.id)) {
      seen.add(cur.id);
      out.unshift(cur);
      cur = cur.basedOn ? this.byId.get(cur.basedOn) : undefined;
    }
    return out;
  }

  /** Resolved paragraph props from docDefaults + default style + named style chain. */
  resolveParaProps(styleId: string | undefined): ParaProps {
    let acc: ParaProps = { ...this.docDefaultPara };
    const id = styleId ?? this.defaultParaId;
    for (const s of this.chain(id)) acc = mergePara(acc, s.pPr);
    acc.styleId = styleId;
    return acc;
  }

  /**
   * Run props inherited by a paragraph's runs: docDefaults rPr, then the rPr
   * declared on the paragraph style chain. Direct run rPr is layered on top by
   * the run parser.
   */
  resolveParaRunProps(styleId: string | undefined): RunProps {
    let acc: RunProps = { ...this.docDefaultRun };
    const id = styleId ?? this.defaultParaId;
    for (const s of this.chain(id)) acc = mergeRun(acc, s.rPr);
    return acc;
  }

  /** Run props from a character style chain (rStyle). */
  resolveCharStyle(styleId: string | undefined): RunProps {
    let acc: RunProps = {};
    for (const s of this.chain(styleId)) acc = mergeRun(acc, s.rPr);
    return acc;
  }
}

// ─── Merging ───────────────────────────────────────────────────────────────────

export function mergeRun(base: RunProps, over: Partial<RunProps>): RunProps {
  return { ...base, ...definedOnly(over) };
}
export function mergePara(base: ParaProps, over: Partial<ParaProps>): ParaProps {
  const out: ParaProps = { ...base, ...definedOnly(over) };
  if (over.borders) out.borders = { ...base.borders, ...over.borders };
  return out;
}

function definedOnly<T extends object>(o: T): Partial<T> {
  const out: Partial<T> = {};
  for (const k in o) {
    const v = (o as Record<string, unknown>)[k];
    if (v !== undefined) (out as Record<string, unknown>)[k] = v;
  }
  return out;
}

// ─── Property parsers (also used for direct pPr/rPr) ─────────────────────────────

/** Word boolean toggle: element present with no val (or val on/true/1) = true. */
export function bool(val: string | undefined): boolean {
  if (val === undefined) return true;
  return val !== '0' && val !== 'false' && val !== 'off';
}

/** Parse a run-properties element (<w:rPr>) into typed partial props. */
export function rPrFrom(rPr: XmlNode | undefined): Partial<RunProps> {
  if (!rPr) return {};
  const out: Partial<RunProps> = {};

  const b = child(rPr, 'b');
  if (b) out.bold = bool(attr(b, 'val'));
  const i = child(rPr, 'i');
  if (i) out.italic = bool(attr(i, 'val'));
  const strike = child(rPr, 'strike');
  if (strike) out.strike = bool(attr(strike, 'val'));
  const u = child(rPr, 'u');
  if (u) out.underline = (attr(u, 'val') ?? 'single') !== 'none';

  const sz = attrNum(child(rPr, 'sz'), 'val');
  if (sz !== undefined) out.sizeHalfPt = sz;

  const color = attr(child(rPr, 'color'), 'val');
  if (color && color !== 'auto') out.colorHex = normHex(color);

  const rFonts = child(rPr, 'rFonts');
  const font = attr(rFonts, 'ascii') ?? attr(rFonts, 'hAnsi') ?? attr(rFonts, 'cs');
  if (font) out.font = font;

  const hl = attr(child(rPr, 'highlight'), 'val');
  if (hl && hl !== 'none') out.highlightHex = namedHighlight(hl);
  const shdFill = attr(child(rPr, 'shd'), 'fill');
  if (shdFill && shdFill !== 'auto' && !out.highlightHex) out.highlightHex = normHex(shdFill);

  const va = attr(child(rPr, 'vertAlign'), 'val');
  if (va === 'superscript') out.vertAlign = 'super';
  else if (va === 'subscript') out.vertAlign = 'sub';

  if (child(rPr, 'caps')) out.caps = 'all';
  else if (child(rPr, 'smallCaps')) out.caps = 'small';

  return out;
}

/** Parse a paragraph-properties element (<w:pPr>) into typed partial props. */
export function pPrFrom(pPr: XmlNode | undefined): Partial<ParaProps> {
  if (!pPr) return {};
  const out: Partial<ParaProps> = {};

  const style = attr(child(pPr, 'pStyle'), 'val');
  if (style) out.styleId = style;

  const jc = attr(child(pPr, 'jc'), 'val');
  if (jc) out.align = mapAlign(jc);

  const ind = child(pPr, 'ind');
  if (ind) {
    const left = attrNum(ind, 'left') ?? attrNum(ind, 'start');
    const right = attrNum(ind, 'right') ?? attrNum(ind, 'end');
    const firstLine = attrNum(ind, 'firstLine');
    const hanging = attrNum(ind, 'hanging');
    if (left !== undefined) out.indentLeftTwip = left;
    if (right !== undefined) out.indentRightTwip = right;
    if (firstLine !== undefined) out.indentFirstLineTwip = firstLine;
    if (hanging !== undefined) out.hangingTwip = hanging;
  }

  const spacing = child(pPr, 'spacing');
  if (spacing) {
    const before = attrNum(spacing, 'before');
    const after = attrNum(spacing, 'after');
    const line = attrNum(spacing, 'line');
    if (before !== undefined) out.spaceBeforeTwip = before;
    if (after !== undefined) out.spaceAfterTwip = after;
    if (line !== undefined) {
      out.line = line;
      out.lineRule = (attr(spacing, 'lineRule') as ParaProps['lineRule']) ?? 'auto';
    }
  }

  if (child(pPr, 'contextualSpacing')) out.contextualSpacing = bool(attr(child(pPr, 'contextualSpacing'), 'val'));
  if (child(pPr, 'keepNext') || child(pPr, 'keepLines')) out.keepTogether = true;
  if (child(pPr, 'pageBreakBefore')) out.pageBreakBefore = bool(attr(child(pPr, 'pageBreakBefore'), 'val'));

  const shdFill = attr(child(pPr, 'shd'), 'fill');
  if (shdFill && shdFill !== 'auto') out.shadingHex = normHex(shdFill);

  const numPr = child(pPr, 'numPr');
  if (numPr) {
    const numId = attrNum(child(numPr, 'numId'), 'val');
    const ilvl = attrNum(child(numPr, 'ilvl'), 'val');
    if (numId !== undefined) out.numId = numId;
    out.ilvl = ilvl ?? 0;
  }

  const pBdr = child(pPr, 'pBdr');
  if (pBdr) {
    const b: RawBorders = {};
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const e = child(pBdr, side);
      if (e) b[side] = { sz: attrNum(e, 'sz'), colorHex: hexOrUndef(attr(e, 'color')), val: attr(e, 'val') };
    }
    out.borders = b;
  }

  return out;
}

function mapAlign(jc: string): ParaProps['align'] {
  switch (jc) {
    case 'center':
      return 'ctr';
    case 'right':
    case 'end':
      return 'r';
    case 'both':
    case 'distribute':
      return 'just';
    default:
      return 'l';
  }
}

function normHex(v: string): string {
  return v.replace(/^#/, '').toLowerCase();
}
function hexOrUndef(v: string | undefined): string | undefined {
  return v && v !== 'auto' ? normHex(v) : undefined;
}

/** Word named highlight colors → hex. */
function namedHighlight(name: string): string {
  const map: Record<string, string> = {
    black: '000000', blue: '0000ff', cyan: '00ffff', green: '00ff00',
    magenta: 'ff00ff', red: 'ff0000', yellow: 'ffff00', white: 'ffffff',
    darkBlue: '000080', darkCyan: '008080', darkGreen: '008000',
    darkMagenta: '800080', darkRed: '800000', darkYellow: '808000',
    darkGray: '808080', lightGray: 'c0c0c0',
  };
  return map[name] ?? 'ffff00';
}
