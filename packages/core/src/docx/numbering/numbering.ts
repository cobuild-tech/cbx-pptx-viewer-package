/**
 * List numbering resolution for WordprocessingML.
 *
 * word/numbering.xml defines abstract numbering definitions (abstractNum) that
 * describe per-level formats, and concrete numbering instances (num) that point
 * to an abstract definition with optional per-level overrides.
 *
 * Paragraphs reference a numId + ilvl to get their bullet/number formatting.
 */
import { child, children, attr, attrNum, type XmlNode } from '../../oxml/xml.js';
import { twipsToPx, halfPtToPt } from '../units.js';
import type { Bullet } from '../model.js';

export interface ListLevel {
  numFmt: string;           // 'bullet', 'decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman', 'none', ...
  lvlText: string;          // bullet char (e.g. '•') or number format (e.g. '%1.')
  startAt: number;
  indentLeftPx: number;
  indentFirstLinePx: number;
  fontAscii?: string;
}

export class NumberingMap {
  /** abstractNumId -> per-level definitions */
  private abstracts = new Map<number, ListLevel[]>();
  /** numId -> abstractNumId + level overrides */
  private nums = new Map<number, { abstractId: number; overrides: Map<number, Partial<ListLevel>> }>();

  static parse(xml: XmlNode | undefined): NumberingMap {
    const map = new NumberingMap();
    if (!xml) return map;

    for (const absEl of children(xml, 'abstractNum')) {
      const id = attrNum(absEl, 'w:abstractNumId') ?? attrNum(absEl, 'abstractNumId');
      if (id === undefined) continue;
      const levels: ListLevel[] = [];
      for (const lvlEl of children(absEl, 'lvl')) {
        levels.push(parseLvl(lvlEl));
      }
      map.abstracts.set(id, levels);
    }

    for (const numEl of children(xml, 'num')) {
      const numId = attrNum(numEl, 'w:numId') ?? attrNum(numEl, 'numId');
      if (numId === undefined) continue;
      const abstractId =
        attrNum(child(numEl, 'abstractNumId'), 'w:val') ??
        attrNum(child(numEl, 'abstractNumId'), 'val') ??
        0;
      const overrides = new Map<number, Partial<ListLevel>>();
      for (const lvlOvr of children(numEl, 'lvlOverride')) {
        const ilvl = attrNum(lvlOvr, 'w:ilvl') ?? attrNum(lvlOvr, 'ilvl') ?? 0;
        const ovr: Partial<ListLevel> = {};
        const startOvr = child(lvlOvr, 'startOverride');
        if (startOvr) {
          ovr.startAt = attrNum(startOvr, 'w:val') ?? attrNum(startOvr, 'val') ?? 1;
        }
        overrides.set(ilvl, ovr);
      }
      map.nums.set(numId, { abstractId, overrides });
    }
    return map;
  }

  resolve(numId: number, ilvl: number): ListLevel | undefined {
    const num = this.nums.get(numId);
    if (!num) return undefined;
    const levels = this.abstracts.get(num.abstractId);
    const base = levels?.[ilvl];
    if (!base) return undefined;
    const ovr = num.overrides.get(ilvl);
    return ovr ? { ...base, ...ovr } : base;
  }

  /** Convert a resolved ListLevel to a Bullet IR node. */
  toBullet(level: ListLevel, counters: Map<string, number>, numId: number, ilvl: number): Bullet {
    if (level.numFmt === 'none') return { type: 'none' };

    if (level.numFmt === 'bullet') {
      return {
        type: 'char',
        char: level.lvlText || '•',
        font: level.fontAscii,
      };
    }

    // Numbered: increment counter for this numId+ilvl.
    const key = `${numId}:${ilvl}`;
    const current = counters.get(key) ?? (level.startAt - 1);
    const next = current + 1;
    counters.set(key, next);

    // Reset deeper levels.
    for (let i = ilvl + 1; i < 9; i++) {
      counters.delete(`${numId}:${i}`);
    }

    // Format the lvlText pattern ('%1.' etc.) with computed number values.
    const displayText = level.lvlText.replace(/%(\d+)/g, (_, n: string) => {
      const targetIlvl = parseInt(n, 10) - 1;
      if (targetIlvl === ilvl) return formatNumber(next, level.numFmt);
      const pKey = `${numId}:${targetIlvl}`;
      return formatNumber(counters.get(pKey) ?? 1, 'decimal');
    });

    return { type: 'char', char: displayText };
  }
}

function formatNumber(n: number, numFmt: string): string {
  switch (numFmt) {
    case 'decimal': return String(n);
    case 'lowerLetter': return nToLetter(n).toLowerCase();
    case 'upperLetter': return nToLetter(n).toUpperCase();
    case 'lowerRoman': return toRoman(n).toLowerCase();
    case 'upperRoman': return toRoman(n);
    default: return String(n);
  }
}

function nToLetter(n: number): string {
  let result = '';
  let num = n;
  while (num > 0) {
    result = String.fromCharCode(64 + ((num - 1) % 26 + 1)) + result;
    num = Math.floor((num - 1) / 26);
  }
  return result || 'a';
}

function toRoman(n: number): string {
  if (n <= 0) return String(n);
  const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
  const syms = ['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
  let result = '';
  for (let i = 0; i < vals.length; i++) {
    while (n >= vals[i]!) { result += syms[i]; n -= vals[i]!; }
  }
  return result;
}

function parseLvl(lvlEl: XmlNode): ListLevel {
  const numFmt =
    attr(child(lvlEl, 'numFmt'), 'w:val') ?? attr(child(lvlEl, 'numFmt'), 'val') ?? 'bullet';
  const lvlTextEl = child(lvlEl, 'lvlText');
  const lvlText =
    attr(lvlTextEl, 'w:val') ?? attr(lvlTextEl, 'val') ?? '•';
  const startAt =
    attrNum(child(lvlEl, 'start'), 'w:val') ?? attrNum(child(lvlEl, 'start'), 'val') ?? 1;

  const pPr = child(lvlEl, 'pPr');
  const indEl = child(pPr, 'ind');
  const left = attrNum(indEl, 'w:left') ?? attrNum(indEl, 'left') ?? 720;
  const hanging = attrNum(indEl, 'w:hanging') ?? attrNum(indEl, 'hanging');
  const firstLine = attrNum(indEl, 'w:firstLine') ?? attrNum(indEl, 'firstLine');

  const rPr = child(lvlEl, 'rPr');
  const fontsEl = child(rPr, 'rFonts');
  const fontAscii = attr(fontsEl, 'w:ascii') ?? attr(fontsEl, 'ascii') ?? attr(fontsEl, 'w:hAnsi') ?? attr(fontsEl, 'hAnsi');

  return {
    numFmt,
    lvlText,
    startAt,
    indentLeftPx: twipsToPx(left),
    indentFirstLinePx: hanging !== undefined ? -twipsToPx(hanging) : firstLine !== undefined ? twipsToPx(firstLine) : 0,
    fontAscii: fontAscii ?? undefined,
  };
}
