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

/** Latin typefaces of the document theme's font scheme. */
export interface ThemeFonts {
  major?: string;
  minor?: string;
  /** East Asian typefaces (the scheme's <a:ea>, or its script font for the document's language). */
  majorEa?: string;
  minorEa?: string;
  /** Complex-script typefaces (<a:cs>). */
  majorCs?: string;
  minorCs?: string;
}

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
  /** <w:rFonts w:eastAsia> (or its theme token): the face for East Asian characters. */
  eastAsiaFont?: string;
  /** <w:rFonts w:cs>: the face for complex-script (e.g. Arabic, Hebrew) runs. */
  csFont?: string;
  /** <w:szCs>: complex-script size, half-points. */
  csSizeHalfPt?: number;
  csBold?: boolean;
  csItalic?: boolean;
  /** <w:rtl>: the run is right-to-left complex-script text. */
  rtl?: boolean;
  /** <w:vanish>: hidden text. */
  hidden?: boolean;
  /** <w:spacing w:val>: extra space between characters, in twips (may be negative). */
  charSpacingTwip?: number;
  /** <w:kern w:val>: kern text at or above this size, in half-points. */
  kernHalfPt?: number;
}

/** One <w:tab> in a <w:tabs> list. `clear` removes an inherited stop at the same position. */
export interface RawTab {
  posTwip: number;
  val: string;
  leader?: string;
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
  keepNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  pageBreakBefore?: boolean;
  /** Tab stops, already merged down the cascade (clears applied). */
  tabs?: RawTab[];
  /** <w:bidi>: a right-to-left paragraph. */
  bidi?: boolean;
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
  /** <w:space>: distance from the text, in points. */
  spacePt?: number;
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
  /** Rows may not break across pages (<w:trPr><w:cantSplit> in the style). */
  cantSplit?: boolean;
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
  cantSplit?: boolean;
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

  /**
   * @param themeFonts the theme's latin major (headings) and minor (body)
   *   typefaces, which `asciiTheme`/`hAnsiTheme` references resolve to.
   */
  static parse(stylesXml: XmlNode | undefined, themeFonts: ThemeFonts = {}): StyleTable {
    const byId = new Map<string, StyleDef>();
    let docRun: Partial<RunProps> = {};
    let docPara: Partial<ParaProps> = {};
    let defaultParaId: string | undefined;
    let defaultCharId: string | undefined;
    let defaultTableId: string | undefined;

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
          else if (type === 'table') defaultTableId = id;
        }
      }
    }
    // Word's own defaults when docDefaults is silent: Times New Roman, 10pt.
    docRun = { font: 'Times New Roman', sizeHalfPt: 20, ...docRun };
    const table = new StyleTable(byId, docRun, docPara, defaultParaId, defaultCharId);
    table.themeFonts = themeFonts;
    table.defaultTableId = defaultTableId;
    return table;
  }

  private themeFonts: ThemeFonts = {};
  private defaultTableId: string | undefined;

  /** Resolve a theme-font token to the theme's typeface; literal names pass through. */
  fontName(font: string | undefined): string | undefined {
    const t = this.themeFonts;
    switch (font) {
      case MAJOR_FONT:
        return t.major ?? t.minor;
      case MINOR_FONT:
        return t.minor ?? t.major;
      case `${MAJOR_FONT}-ea`:
        return t.majorEa ?? t.minorEa;
      case `${MINOR_FONT}-ea`:
        return t.minorEa ?? t.majorEa;
      case `${MAJOR_FONT}-cs`:
        return t.majorCs ?? t.minorCs;
      case `${MINOR_FONT}-cs`:
        return t.minorCs ?? t.majorCs;
      default:
        return font;
    }
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

  /**
   * A paragraph style's properties split in two layers: what comes from
   * docDefaults and the default ("Normal") style, and what the style's own
   * chain above that adds. A table style slots *between* the two: Word lets it
   * override Normal (which is why a cell loses Normal's space-after) but not a
   * style the paragraph names.
   */
  paraLayers(styleId: string | undefined): {
    basePara: ParaProps;
    ownPara: Partial<ParaProps>;
    baseRun: RunProps;
    ownRun: Partial<RunProps>;
  } {
    const chain = this.chain(styleId ?? this.defaultParaId);
    // Normal is in the base only for styles that actually inherit from it.
    const dflt = chain.find((st) => st.id === this.defaultParaId);
    let basePara: ParaProps = { ...this.docDefaultPara };
    let baseRun: RunProps = { ...this.docDefaultRun };
    if (dflt) {
      basePara = mergePara(basePara, dflt.pPr);
      baseRun = mergeRun(baseRun, dflt.rPr);
    }
    let ownPara: ParaProps = {};
    let ownRun: RunProps = {};
    for (const st of chain) {
      if (st === dflt) continue;
      ownPara = mergePara(ownPara, st.pPr);
      ownRun = mergeRun(ownRun, st.rPr);
    }
    return { basePara, ownPara, baseRun, ownRun };
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
    // A table naming no style gets the document's default table style ("Normal
    // Table"), which is where Word's 0.08" left/right cell margins come from.
    const id = styleId && this.byId.has(styleId) ? styleId : this.defaultTableId;
    if (!id) return undefined;
    const chain = this.chain(id).filter((d) => d.tbl);
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
      if (t.cantSplit !== undefined) res.cantSplit = t.cantSplit;
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

  const cantSplit = child(child(s, 'trPr'), 'cantSplit');
  if (cantSplit) part.cantSplit = bool(attr(cantSplit, 'val'));

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
  if (over.tabs) out.tabs = mergeTabs(base.tabs, over.tabs);
  return out;
}

/**
 * Tab stops accumulate down the cascade rather than replacing each other: a
 * style's stops stay unless a later layer states `clear` at the same position.
 */
function mergeTabs(base: RawTab[] | undefined, over: RawTab[]): RawTab[] {
  const byPos = new Map<number, RawTab>();
  for (const t of base ?? []) byPos.set(t.posTwip, t);
  for (const t of over) {
    if (t.val === 'clear') byPos.delete(t.posTwip);
    else byPos.set(t.posTwip, t);
  }
  return [...byPos.values()].sort((a, b) => a.posTwip - b.posTwip);
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
  const strike = child(rPr, 'strike') ?? child(rPr, 'dstrike');
  if (strike) out.strike = bool(attr(strike, 'val'));
  const vanish = child(rPr, 'vanish');
  if (vanish) out.hidden = bool(attr(vanish, 'val'));
  const charSpacing = attrNum(child(rPr, 'spacing'), 'val');
  if (charSpacing !== undefined) out.charSpacingTwip = charSpacing;
  const kern = attrNum(child(rPr, 'kern'), 'val');
  if (kern !== undefined) out.kernHalfPt = kern;
  const u = child(rPr, 'u');
  if (u) out.underline = (attr(u, 'val') ?? 'single') !== 'none';

  const sz = attrNum(child(rPr, 'sz'), 'val');
  if (sz !== undefined) out.sizeHalfPt = sz;

  const color = attr(child(rPr, 'color'), 'val');
  if (color && color !== 'auto') out.colorHex = normHex(color);

  // A theme reference wins over a literal face in the same <w:rFonts> — Word
  // writes both, and draws the theme's. Kept as a token until the style table
  // resolves it against the document theme.
  const rFonts = child(rPr, 'rFonts');
  const theme = attr(rFonts, 'asciiTheme') ?? attr(rFonts, 'hAnsiTheme');
  const font = theme ? themeToken(theme) : (attr(rFonts, 'ascii') ?? attr(rFonts, 'hAnsi'));
  if (font) out.font = font;
  const eaTheme = attr(rFonts, 'eastAsiaTheme');
  const ea = eaTheme ? themeToken(eaTheme) : attr(rFonts, 'eastAsia');
  if (ea) out.eastAsiaFont = ea;
  const csTheme = attr(rFonts, 'cstheme');
  const cs = csTheme ? themeToken(csTheme) : attr(rFonts, 'cs');
  if (cs) out.csFont = cs;
  const szCs = attrNum(child(rPr, 'szCs'), 'val');
  if (szCs !== undefined) out.csSizeHalfPt = szCs;
  const bCs = child(rPr, 'bCs');
  if (bCs) out.csBold = bool(attr(bCs, 'val'));
  const iCs = child(rPr, 'iCs');
  if (iCs) out.csItalic = bool(attr(iCs, 'val'));
  const rtl = child(rPr, 'rtl');
  if (rtl) out.rtl = bool(attr(rtl, 'val'));

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
    // HTML-style automatic spacing, which Word renders as 14pt.
    if (bool(attr(spacing, 'beforeAutospacing') ?? 'false')) out.spaceBeforeTwip = 280;
    if (bool(attr(spacing, 'afterAutospacing') ?? 'false')) out.spaceAfterTwip = 280;
    if (line !== undefined) {
      out.line = line;
      out.lineRule = (attr(spacing, 'lineRule') as ParaProps['lineRule']) ?? 'auto';
    }
  }

  if (child(pPr, 'contextualSpacing')) out.contextualSpacing = bool(attr(child(pPr, 'contextualSpacing'), 'val'));
  const flag = (name: string): boolean | undefined => {
    const el = child(pPr, name);
    return el ? bool(attr(el, 'val')) : undefined;
  };
  const keepNext = flag('keepNext');
  if (keepNext !== undefined) out.keepNext = keepNext;
  const keepLines = flag('keepLines');
  if (keepLines !== undefined) out.keepLines = keepLines;
  const widowControl = flag('widowControl');
  if (widowControl !== undefined) out.widowControl = widowControl;

  const tabs = child(pPr, 'tabs');
  if (tabs) {
    out.tabs = children(tabs, 'tab').flatMap((t) => {
      const pos = attrNum(t, 'pos');
      if (pos === undefined) return [];
      const tab: RawTab = { posTwip: pos, val: attr(t, 'val') ?? 'left' };
      const leader = attr(t, 'leader');
      if (leader && leader !== 'none') tab.leader = leader;
      return [tab];
    });
  }
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

  const bidi = child(pPr, 'bidi');
  if (bidi) out.bidi = bool(attr(bidi, 'val'));

  const pBdr = child(pPr, 'pBdr');
  if (pBdr) {
    const b: RawBorders = {};
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const e = child(pBdr, side);
      if (e) b[side] = { sz: attrNum(e, 'sz'), colorHex: hexOrUndef(attr(e, 'color')), val: attr(e, 'val'), spacePt: attrNum(e, 'space') };
    }
    out.borders = b;
  }

  return out;
}

/** Token standing in for a theme font until {@link StyleTable.fontName} resolves it. */
const MAJOR_FONT = '+major';
const MINOR_FONT = '+minor';

function themeToken(theme: string): string {
  const base = theme.startsWith('major') ? MAJOR_FONT : MINOR_FONT;
  if (/EastAsia$/.test(theme)) return `${base}-ea`;
  if (/Bidi$/.test(theme)) return `${base}-cs`;
  return base;
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
