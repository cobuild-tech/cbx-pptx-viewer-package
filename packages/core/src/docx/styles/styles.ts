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

/** Cell-level formatting a table style (or conditional format) contributes. */
export interface CellShade {
  fillHex?: string;
  vAlign?: 'top' | 'center' | 'bottom';
  borders?: RawBorders;
}

/** One layer of table formatting: whole-table defaults or a conditional format. */
export interface TableCond {
  pPr: Partial<ParaProps>;
  rPr: Partial<RunProps>;
  tc: CellShade;
}

/** A table style resolved through its basedOn chain, ready to apply per cell. */
export interface ResolvedTableStyle {
  /** Outer table borders (top/bottom/left/right). */
  tblBorders: RawBorders;
  insideH?: RawBorder;
  insideV?: RawBorder;
  /** Default cell margins in twips. */
  cellMar?: { top: number; right: number; bottom: number; left: number };
  rowBandSize: number;
  colBandSize: number;
  /** Whole-table defaults. */
  whole: TableCond;
  /** Conditional formats keyed by <w:tblStylePr w:type> (firstRow, band1Horz, …). */
  cond: Map<string, TableCond>;
}

/** Which conditional formats a table enables (from <w:tblLook>). */
export interface TableLook {
  firstRow: boolean;
  lastRow: boolean;
  firstCol: boolean;
  lastCol: boolean;
  noHBand: boolean;
  noVBand: boolean;
}

/** Per-style table data (before the basedOn chain is merged). */
interface TableStylePart {
  tblBorders?: RawBorders;
  insideH?: RawBorder;
  insideV?: RawBorder;
  cellMar?: { top: number; right: number; bottom: number; left: number };
  rowBandSize?: number;
  colBandSize?: number;
  tc: CellShade;
  cond: Map<string, TableCond>;
}

interface StyleDef {
  id: string;
  type: string;
  name: string;
  basedOn?: string;
  linkId?: string;
  pPr: Partial<ParaProps>;
  rPr: Partial<RunProps>;
  /** Table-only formatting (present when type === 'table'). */
  tbl?: TableStylePart;
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
        if (type === 'table') def.tbl = readTablePart(s);
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

  /**
   * Resolve a table style (via <w:tblStyle>) through its basedOn chain into a
   * flat {@link ResolvedTableStyle}: outer/inside borders, default cell margins,
   * whole-table defaults, and the conditional formats (firstRow, bands, …).
   */
  resolveTableStyle(styleId: string | undefined): ResolvedTableStyle | undefined {
    if (!styleId) return undefined;
    const chain = this.chain(styleId).filter((d) => d.tbl);
    if (!chain.length) return undefined;

    const res: ResolvedTableStyle = {
      tblBorders: {},
      rowBandSize: 1,
      colBandSize: 1,
      whole: { pPr: {}, rPr: {}, tc: {} },
      cond: new Map(),
    };
    for (const d of chain) {
      const t = d.tbl!;
      res.whole.pPr = mergePara(res.whole.pPr as ParaProps, d.pPr);
      res.whole.rPr = mergeRun(res.whole.rPr as RunProps, d.rPr);
      res.whole.tc = mergeCellShade(res.whole.tc, t.tc);
      if (t.tblBorders) res.tblBorders = { ...res.tblBorders, ...t.tblBorders };
      if (t.insideH) res.insideH = t.insideH;
      if (t.insideV) res.insideV = t.insideV;
      if (t.cellMar) res.cellMar = t.cellMar;
      if (t.rowBandSize !== undefined) res.rowBandSize = t.rowBandSize;
      if (t.colBandSize !== undefined) res.colBandSize = t.colBandSize;
      for (const [type, cond] of t.cond) {
        const prev = res.cond.get(type);
        res.cond.set(type, prev ? mergeTableCond(prev, cond) : cond);
      }
    }
    return res;
  }
}

/** Merge two conditional-format layers (`over` wins). */
function mergeTableCond(base: TableCond, over: TableCond): TableCond {
  return {
    pPr: mergePara(base.pPr as ParaProps, over.pPr),
    rPr: mergeRun(base.rPr as RunProps, over.rPr),
    tc: mergeCellShade(base.tc, over.tc),
  };
}

function mergeCellShade(base: CellShade, over: CellShade): CellShade {
  const out: CellShade = { ...base, ...definedOnly(over) };
  if (base.borders || over.borders) out.borders = { ...base.borders, ...over.borders };
  return out;
}

/**
 * Effective formatting for one cell: layer the whole-table defaults, then the
 * enabled conditional formats in ascending precedence (bands < first/last
 * col < first/last row < corner cells), exactly as Word applies them.
 */
export function tableCellFormat(
  ts: ResolvedTableStyle,
  row: number,
  col: number,
  rowCount: number,
  colCount: number,
  look: TableLook,
): TableCond {
  const layers: TableCond[] = [ts.whole];
  const get = (type: string): void => {
    const c = ts.cond.get(type);
    if (c) layers.push(c);
  };

  // Banding runs over the "inner" rows/cols (excluding a first/last that has its
  // own conditional). Band 1 is the first inner band; alternates 1,2,1,2,…
  if (!look.noVBand) {
    const start = look.firstCol ? 1 : 0;
    const end = colCount - (look.lastCol ? 1 : 0);
    if (col >= start && col < end) {
      const band = Math.floor((col - start) / (ts.colBandSize || 1));
      get(band % 2 === 0 ? 'band1Vert' : 'band2Vert');
    }
  }
  if (!look.noHBand) {
    const start = look.firstRow ? 1 : 0;
    const end = rowCount - (look.lastRow ? 1 : 0);
    if (row >= start && row < end) {
      const band = Math.floor((row - start) / (ts.rowBandSize || 1));
      get(band % 2 === 0 ? 'band1Horz' : 'band2Horz');
    }
  }
  if (look.firstCol && col === 0) get('firstCol');
  if (look.lastCol && col === colCount - 1) get('lastCol');
  if (look.firstRow && row === 0) get('firstRow');
  if (look.lastRow && row === rowCount - 1) get('lastRow');
  if (look.firstRow && look.firstCol && row === 0 && col === 0) get('nwCell');
  if (look.firstRow && look.lastCol && row === 0 && col === colCount - 1) get('neCell');
  if (look.lastRow && look.firstCol && row === rowCount - 1 && col === 0) get('swCell');
  if (look.lastRow && look.lastCol && row === rowCount - 1 && col === colCount - 1) get('seCell');

  let out: TableCond = { pPr: {}, rPr: {}, tc: {} };
  for (const l of layers) out = mergeTableCond(out, l);
  return out;
}

/** Parse a <w:tblLook> element (or its hex val bitmask) into typed flags. */
export function parseTableLook(el: XmlNode | undefined): TableLook {
  const flag = (name: string, bit: number): boolean => {
    const a = attr(el, name);
    if (a !== undefined) return bool(a);
    const val = attr(el, 'val');
    if (val) return (parseInt(val, 16) & bit) !== 0;
    return false;
  };
  return {
    firstRow: flag('firstRow', 0x0020),
    lastRow: flag('lastRow', 0x0040),
    firstCol: flag('firstColumn', 0x0080),
    lastCol: flag('lastColumn', 0x0100),
    // noHBand/noVBand: the val bits are inverted (set bit = banding OFF).
    noHBand: flag('noHBand', 0x0200),
    noVBand: flag('noVBand', 0x0400),
  };
}

// ─── Table style parsing ─────────────────────────────────────────────────────

function readTablePart(s: XmlNode): TableStylePart {
  const part: TableStylePart = { tc: readCellShade(child(s, 'tcPr')), cond: new Map() };

  const tblPr = child(s, 'tblPr');
  if (tblPr) {
    const bordersNode = child(tblPr, 'tblBorders');
    if (bordersNode) {
      part.tblBorders = readSideBorders(bordersNode);
      part.insideH = rawBorder(child(bordersNode, 'insideH'));
      part.insideV = rawBorder(child(bordersNode, 'insideV'));
    }
    const cm = child(tblPr, 'tblCellMar');
    if (cm) part.cellMar = readCellMarTwips(cm);
    const rb = attrNum(child(tblPr, 'tblStyleRowBandSize'), 'val');
    if (rb !== undefined) part.rowBandSize = rb;
    const cb = attrNum(child(tblPr, 'tblStyleColBandSize'), 'val');
    if (cb !== undefined) part.colBandSize = cb;
  }

  for (const sp of children(s, 'tblStylePr')) {
    const type = attr(sp, 'type');
    if (!type) continue;
    part.cond.set(type, {
      pPr: pPrFrom(child(sp, 'pPr')),
      rPr: rPrFrom(child(sp, 'rPr')),
      tc: readCellShade(child(sp, 'tcPr')),
    });
  }
  return part;
}

function readCellShade(tcPr: XmlNode | undefined): CellShade {
  if (!tcPr) return {};
  const out: CellShade = {};
  const fill = attr(child(tcPr, 'shd'), 'fill');
  if (fill && fill !== 'auto') out.fillHex = normHex(fill);
  const va = attr(child(tcPr, 'vAlign'), 'val');
  if (va === 'center' || va === 'bottom' || va === 'top') out.vAlign = va;
  const tb = child(tcPr, 'tcBorders');
  if (tb) {
    const b = readSideBorders(tb);
    if (Object.keys(b).length) out.borders = b;
  }
  return out;
}

function rawBorder(e: XmlNode | undefined): RawBorder | undefined {
  if (!e) return undefined;
  return { sz: attrNum(e, 'sz'), colorHex: hexOrUndef(attr(e, 'color')), val: attr(e, 'val') };
}

function readSideBorders(node: XmlNode): RawBorders {
  const out: RawBorders = {};
  for (const side of ['top', 'bottom', 'left', 'right'] as const) {
    const b = rawBorder(child(node, side));
    if (b) out[side] = b;
  }
  return out;
}

function readCellMarTwips(mar: XmlNode): { top: number; right: number; bottom: number; left: number } {
  const side = (name: string): number => attrNum(child(mar, name), 'w') ?? 0;
  return { top: side('top'), right: side('right'), bottom: side('bottom'), left: side('left') };
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
