/**
 * List numbering (numbering.xml).
 *
 * Resolves a paragraph's (numId, ilvl) to a rendered list marker ("•", "1.",
 * "a)", "II.", …). Ordered lists need running counters, so {@link Numbering} is
 * stateful: the body walker calls {@link Numbering.marker} in document order and
 * counters advance/reset exactly as Word tracks them.
 */
import { child, children, attr, attrNum, type XmlNode } from '../../oxml/xml.js';

interface LvlDef {
  numFmt: string;
  lvlText: string;
  start: number;
  /** Left indent (twips) from the level's pPr, if any. */
  indentLeftTwip?: number;
  hangingTwip?: number;
  bulletFont?: string;
}

export class Numbering {
  /** abstractNumId -> ilvl -> level definition */
  private readonly abstract = new Map<number, Map<number, LvlDef>>();
  /** numId -> abstractNumId */
  private readonly numToAbstract = new Map<number, number>();
  /** live counters: numId -> ilvl -> current value */
  private readonly counters = new Map<number, Map<number, number>>();

  private constructor(
    abstract: Map<number, Map<number, LvlDef>>,
    numToAbstract: Map<number, number>,
  ) {
    this.abstract = abstract;
    this.numToAbstract = numToAbstract;
  }

  static parse(numberingXml: XmlNode | undefined): Numbering {
    const abstract = new Map<number, Map<number, LvlDef>>();
    const numToAbstract = new Map<number, number>();
    if (!numberingXml) return new Numbering(abstract, numToAbstract);

    for (const an of children(numberingXml, 'abstractNum')) {
      const id = attrNum(an, 'abstractNumId');
      if (id === undefined) continue;
      const levels = new Map<number, LvlDef>();
      for (const lvl of children(an, 'lvl')) {
        const ilvl = attrNum(lvl, 'ilvl') ?? 0;
        const ind = child(child(lvl, 'pPr'), 'ind');
        levels.set(ilvl, {
          numFmt: attr(child(lvl, 'numFmt'), 'val') ?? 'decimal',
          lvlText: attr(child(lvl, 'lvlText'), 'val') ?? '',
          start: attrNum(child(lvl, 'start'), 'val') ?? 1,
          indentLeftTwip: attrNum(ind, 'left') ?? attrNum(ind, 'start'),
          hangingTwip: attrNum(ind, 'hanging'),
          bulletFont: attr(child(child(lvl, 'rPr'), 'rFonts'), 'ascii'),
        });
      }
      abstract.set(id, levels);
    }

    for (const num of children(numberingXml, 'num')) {
      const numId = attrNum(num, 'numId');
      const absId = attrNum(child(num, 'abstractNumId'), 'val');
      if (numId !== undefined && absId !== undefined) numToAbstract.set(numId, absId);
    }
    return new Numbering(abstract, numToAbstract);
  }

  private lvlDef(numId: number, ilvl: number): LvlDef | undefined {
    const absId = this.numToAbstract.get(numId);
    if (absId === undefined) return undefined;
    return this.abstract.get(absId)?.get(ilvl);
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
   * Advance counters for (numId, ilvl) and return the rendered marker text.
   * Call once per list paragraph, in document order.
   */
  marker(numId: number, ilvl: number): string {
    const def = this.lvlDef(numId, ilvl);
    if (!def) return '';

    if (def.numFmt === 'bullet') return bulletChar(def.lvlText, def.bulletFont);

    // Ordered: advance this level, reset deeper levels.
    let counter = this.counters.get(numId);
    if (!counter) {
      counter = new Map();
      this.counters.set(numId, counter);
    }
    const cur = counter.has(ilvl) ? counter.get(ilvl)! + 1 : def.start;
    counter.set(ilvl, cur);
    for (const [lvl] of counter) if (lvl > ilvl) counter.delete(lvl);

    // Substitute %1..%9 in lvlText with each level's formatted counter.
    return def.lvlText.replace(/%(\d)/g, (_m, d: string) => {
      const targetLvl = Number(d) - 1;
      const val = targetLvl === ilvl ? cur : counter!.get(targetLvl) ?? this.lvlDef(numId, targetLvl)?.start ?? 1;
      const fmt = this.lvlDef(numId, targetLvl)?.numFmt ?? 'decimal';
      return formatNumber(val, fmt);
    });
  }
}

function bulletChar(lvlText: string, font?: string): string {
  // Symbol/Wingdings bullets map to common glyphs; otherwise pass through.
  const code = lvlText.codePointAt(0);
  if (font === 'Symbol' && code === 0xf0b7) return '•';
  if (font === 'Wingdings' && code === 0xf0a7) return '▪';
  if (font === 'Wingdings' && code === 0xf06e) return '■';
  if (font === 'Courier New' && lvlText === 'o') return '◦';
  return lvlText || '•';
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
    default:
      return String(n);
  }
}

function toLetter(n: number): string {
  let s = '';
  let x = n;
  while (x > 0) {
    const rem = (x - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s || 'A';
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
