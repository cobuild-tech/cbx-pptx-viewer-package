/**
 * List numbering (numbering.xml).
 *
 * Resolves a paragraph's (numId, ilvl) to a rendered list marker ("•", "1.",
 * "a)", "II.", …). Ordered lists need running counters, so {@link Numbering} is
 * stateful: the body walker calls {@link Numbering.marker} in document order and
 * counters advance/reset exactly as Word tracks them.
 */
import { child, children, attr, attrNum, type XmlNode } from '../../oxml/xml.js';
import { rPrFrom, type RunProps } from '../styles/styles.js';

interface LvlDef {
  numFmt: string;
  lvlText: string;
  start: number;
  /** Left indent (twips) from the level's pPr, if any. */
  indentLeftTwip?: number;
  hangingTwip?: number;
  /** Character formatting of the marker (<w:lvl><w:rPr>). */
  rPr: Partial<RunProps>;
  /** What follows the marker (<w:suff>). */
  suffix: 'tab' | 'space' | 'nothing';
  /** <w:isLgl>: render every level's counter as an arabic number. */
  legal: boolean;
}

/** A resolved list marker for one paragraph. */
export interface ListMarker {
  text: string;
  /**
   * The symbol font a bullet was drawn in before being mapped to Unicode: the
   * glyph no longer uses it, but the line still takes its height from it.
   */
  metricsFont?: string;
  /** Marker formatting from the level definition (font already a literal face). */
  rPr: Partial<RunProps>;
  suffix: 'tab' | 'space' | 'nothing';
}

/** One <w:num>: its abstract definition plus per-level overrides. */
interface NumDef {
  abstractId: number;
  /** ilvl -> replacement start value (<w:lvlOverride><w:startOverride>). */
  startOverride: Map<number, number>;
  /** ilvl -> replacement level definition (<w:lvlOverride><w:lvl>). */
  lvlOverride: Map<number, LvlDef>;
}

export class Numbering {
  /** abstractNumId -> ilvl -> level definition */
  private readonly abstract = new Map<number, Map<number, LvlDef>>();
  /** numId -> its definition */
  private readonly nums = new Map<number, NumDef>();
  /** live counters: numId -> ilvl -> current value */
  private readonly counters = new Map<number, Map<number, number>>();

  private constructor(abstract: Map<number, Map<number, LvlDef>>, nums: Map<number, NumDef>) {
    this.abstract = abstract;
    this.nums = nums;
  }

  static parse(numberingXml: XmlNode | undefined): Numbering {
    const abstract = new Map<number, Map<number, LvlDef>>();
    const nums = new Map<number, NumDef>();
    if (!numberingXml) return new Numbering(abstract, nums);

    for (const an of children(numberingXml, 'abstractNum')) {
      const id = attrNum(an, 'abstractNumId');
      if (id === undefined) continue;
      const levels = new Map<number, LvlDef>();
      for (const lvl of children(an, 'lvl')) levels.set(attrNum(lvl, 'ilvl') ?? 0, readLvl(lvl));
      abstract.set(id, levels);
    }

    for (const num of children(numberingXml, 'num')) {
      const numId = attrNum(num, 'numId');
      const absId = attrNum(child(num, 'abstractNumId'), 'val');
      if (numId === undefined || absId === undefined) continue;
      const def: NumDef = { abstractId: absId, startOverride: new Map(), lvlOverride: new Map() };
      for (const ov of children(num, 'lvlOverride')) {
        const ilvl = attrNum(ov, 'ilvl') ?? 0;
        const start = attrNum(child(ov, 'startOverride'), 'val');
        if (start !== undefined) def.startOverride.set(ilvl, start);
        const lvl = child(ov, 'lvl');
        if (lvl) def.lvlOverride.set(ilvl, readLvl(lvl));
      }
      nums.set(numId, def);
    }
    return new Numbering(abstract, nums);
  }

  private lvlDef(numId: number, ilvl: number): LvlDef | undefined {
    const num = this.nums.get(numId);
    if (!num) return undefined;
    return num.lvlOverride.get(ilvl) ?? this.abstract.get(num.abstractId)?.get(ilvl);
  }

  private startOf(numId: number, ilvl: number): number {
    return this.nums.get(numId)?.startOverride.get(ilvl) ?? this.lvlDef(numId, ilvl)?.start ?? 1;
  }

  /** Indent hint (px) from the numbering level, for lists lacking paragraph ind. */
  levelIndent(numId: number, ilvl: number): { leftPx?: number; hangingPx?: number } {
    const def = this.lvlDef(numId, ilvl);
    return {
      leftPx: def?.indentLeftTwip !== undefined ? def.indentLeftTwip / 15 : undefined,
      hangingPx: def?.hangingTwip !== undefined ? def.hangingTwip / 15 : undefined,
    };
  }

  /**
   * Advance counters for (numId, ilvl) and return the rendered marker. Call
   * once per list paragraph, in document order. numId 0 means "no list".
   */
  marker(numId: number, ilvl: number): ListMarker | undefined {
    const def = this.lvlDef(numId, ilvl);
    if (!def) return undefined;

    // Ordered (and bullet) levels both count, so a bullet between items of a
    // deeper ordered level still resets it like Word does.
    let counter = this.counters.get(numId);
    if (!counter) {
      counter = new Map();
      this.counters.set(numId, counter);
    }
    const cur = counter.has(ilvl) ? counter.get(ilvl)! + 1 : this.startOf(numId, ilvl);
    counter.set(ilvl, cur);
    for (const [lvl] of counter) if (lvl > ilvl) counter.delete(lvl);

    if (def.numFmt === 'none') return { text: '', rPr: def.rPr, suffix: def.suffix };
    if (def.numFmt === 'bullet') {
      const { text, symbolic } = bulletChar(def.lvlText, def.rPr.font);
      // A symbol-font bullet has been mapped to Unicode, so drawing it in the
      // symbol font (which the browser lacks or maps differently) would undo that.
      const rPr = symbolic ? { ...def.rPr, font: undefined } : def.rPr;
      return { text, rPr, suffix: def.suffix, ...(symbolic && def.rPr.font ? { metricsFont: def.rPr.font } : {}) };
    }

    // Substitute %1..%9 in lvlText with each level's formatted counter.
    const text = def.lvlText.replace(/%(\d)/g, (_m, d: string) => {
      const targetLvl = Number(d) - 1;
      const val = targetLvl === ilvl ? cur : (counter!.get(targetLvl) ?? this.startOf(numId, targetLvl));
      const fmt = def.legal ? 'decimal' : (this.lvlDef(numId, targetLvl)?.numFmt ?? 'decimal');
      return formatNumber(val, fmt);
    });
    return { text, rPr: def.rPr, suffix: def.suffix };
  }
}

function readLvl(lvl: XmlNode): LvlDef {
  const ind = child(child(lvl, 'pPr'), 'ind');
  const suff = attr(child(lvl, 'suff'), 'val');
  const isLgl = child(lvl, 'isLgl');
  return {
    numFmt: attr(child(lvl, 'numFmt'), 'val') ?? 'decimal',
    lvlText: attr(child(lvl, 'lvlText'), 'val') ?? '',
    start: attrNum(child(lvl, 'start'), 'val') ?? 1,
    indentLeftTwip: attrNum(ind, 'left') ?? attrNum(ind, 'start'),
    hangingTwip: attrNum(ind, 'hanging'),
    rPr: rPrFrom(child(lvl, 'rPr')),
    suffix: suff === 'space' || suff === 'nothing' ? suff : 'tab',
    legal: isLgl !== undefined && attr(isLgl, 'val') !== '0' && attr(isLgl, 'val') !== 'false',
  };
}

/**
 * Glyphs Word draws for the private-use code points of its symbol fonts
 * (F0xx = the font's own 0xXX slot). Symbol, Wingdings and their kin are not
 * web fonts, so the Unicode equivalent is the only way to show the same mark.
 */
const SYMBOL_FONT_GLYPHS: Record<string, Record<number, string>> = {
  symbol: { 0xb7: '•', 0x2d: '−', 0xa7: '♣', 0xa8: '♦', 0xa9: '♥', 0xaa: '♠', 0xae: '→', 0xde: '⇒', 0xd8: '¬', 0x6f: 'ο' },
  wingdings: {
    0x6c: '●', 0x6d: '❍', 0x6e: '■', 0x6f: '□', 0x70: '◻', 0x71: '❑', 0x72: '❒', 0x73: '⬧', 0x74: '⧫',
    0x75: '◆', 0x76: '❖', 0x77: '⬥', 0x9f: '•', 0xa1: '○', 0xa7: '▪', 0xa8: '◻', 0xd8: '➢', 0xe0: '→',
    0xe8: '➔', 0xf0: '⇨', 0xfb: '✗', 0xfc: '✓', 0xfe: '☑', 0x4a: '☺', 0x46: '☞',
  },
  'wingdings 2': { 0x97: '●', 0x98: '○', 0xa2: '■', 0xa3: '□', 0x50: '✓', 0x4f: '✗' },
  'wingdings 3': { 0x7d: '▶', 0x75: '►', 0xc6: '▶', 0x84: '▸' },
};

function bulletChar(lvlText: string, font?: string): { text: string; symbolic: boolean } {
  const code = lvlText.codePointAt(0);
  const glyphs = font ? SYMBOL_FONT_GLYPHS[font.toLowerCase()] : undefined;
  if (glyphs && code !== undefined) {
    // Symbol fonts are addressed either at F0xx or at the plain 0xXX slot.
    const mapped = glyphs[code & 0xff];
    if (mapped) return { text: mapped, symbolic: true };
    return { text: '•', symbolic: true };
  }
  if (code !== undefined && code >= 0xf000 && code <= 0xf0ff) return { text: '•', symbolic: true };
  return { text: lvlText || '•', symbolic: false };
}

function formatNumber(n: number, fmt: string): string {
  switch (fmt) {
    case 'lowerLetter':
      return toLetter(n).toLowerCase();
    case 'upperLetter':
      return toLetter(n).toUpperCase();
    case 'lowerRoman':
      return toRoman(n).toLowerCase();
    case 'upperRoman':
      return toRoman(n).toUpperCase();
    case 'decimalZero':
      return n < 10 ? `0${n}` : String(n);
    case 'ordinal':
      return `${n}${ordinalSuffix(n)}`;
    default:
      return String(n);
  }
}

function ordinalSuffix(n: number): string {
  const t = n % 100;
  if (t >= 11 && t <= 13) return 'th';
  return ['th', 'st', 'nd', 'rd'][n % 10] ?? 'th';
}

/** Word's letter numbering repeats the letter: a…z, aa…zz, aaa… */
function toLetter(n: number): string {
  if (n <= 0) return 'A';
  const letter = String.fromCharCode(65 + ((n - 1) % 26));
  return letter.repeat(Math.floor((n - 1) / 26) + 1);
}

function toRoman(n: number): string {
  const table: [number, string][] = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  let x = n;
  let out = '';
  for (const [v, sym] of table) {
    while (x >= v) {
      out += sym;
      x -= v;
    }
  }
  return out || 'I';
}

/** Format a page number for a section's <w:pgNumType w:fmt>. */
export function formatPageNumber(n: number, fmt: string | undefined): string {
  return formatNumber(n, fmt ?? 'decimal');
}
