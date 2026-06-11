/**
 * Text-style inheritance.
 *
 * A run's final formatting is the merge, most-specific first, of:
 *   run rPr  ->  paragraph defRPr  ->  list-style level (lvlNpPr/defRPr)
 *   ->  inherited placeholder list style  ->  master txStyles  ->  theme defaults.
 *
 * The resolver assembles the ordered chain of `<a:lstStyle>`-like nodes for a
 * shape (see resolve/shape.ts) and hands it here. This class answers
 * "what are the effective run/paragraph properties at level N?".
 */
import { child, attr, attrNum, attrBool, type XmlNode } from '../xml.js';
import { emuToPx } from '../units.js';
import type { Bullet, Color, TextAlign } from '../model.js';
import { resolveContainerColor, type ColorContext } from './color.js';
import type { Theme } from './color.js';

export interface ResolvedRunProps {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  sizePt?: number;
  color?: Color;
  font?: string;
  baseline?: number;
}

export interface ResolvedParaProps {
  align?: TextAlign;
  marginLeftPx?: number;
  indentPx?: number;
  bullet?: Bullet;
  lineSpacingPct?: number;
  lineSpacingPt?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
}

function lvlKey(level: number): string {
  return `lvl${Math.min(9, level + 1)}pPr`;
}

function resolveFont(typeface: string | undefined, theme: Theme): string | undefined {
  if (!typeface) return undefined;
  if (typeface === '+mj-lt' || typeface === '+mj-ea' || typeface === '+mj-cs') {
    return theme.majorFont;
  }
  if (typeface === '+mn-lt' || typeface === '+mn-ea' || typeface === '+mn-cs') {
    return theme.minorFont;
  }
  return typeface;
}

/** Read run properties from a single rPr/defRPr/endParaRPr node. */
function runPropsFrom(
  rPr: XmlNode | undefined,
  ctx: ColorContext,
): ResolvedRunProps {
  if (!rPr) return {};
  const props: ResolvedRunProps = {};
  const sz = attrNum(rPr, 'sz');
  if (sz !== undefined) props.sizePt = sz / 100;
  if (attr(rPr, 'b') !== undefined) props.bold = attrBool(rPr, 'b');
  if (attr(rPr, 'i') !== undefined) props.italic = attrBool(rPr, 'i');
  const u = attr(rPr, 'u');
  if (u !== undefined) props.underline = u !== 'none';
  const strike = attr(rPr, 'strike');
  if (strike !== undefined) props.strike = strike !== 'noStrike';
  const baseline = attrNum(rPr, 'baseline');
  if (baseline !== undefined) props.baseline = baseline / 1000;

  const solidFill = child(rPr, 'solidFill');
  if (solidFill) {
    const color = resolveContainerColor(solidFill, ctx);
    if (color) props.color = color;
  }
  const typeface = attr(child(rPr, 'latin'), 'typeface');
  const font = resolveFont(typeface, ctx.theme);
  if (font) props.font = font;
  return props;
}

/** Read paragraph properties from a single pPr/lvlNpPr node. */
function paraPropsFrom(pPr: XmlNode | undefined, ctx: ColorContext): ResolvedParaProps {
  if (!pPr) return {};
  const props: ResolvedParaProps = {};
  const algn = attr(pPr, 'algn');
  if (algn === 'l' || algn === 'ctr' || algn === 'r' || algn === 'just') {
    props.align = algn;
  }
  const marL = attrNum(pPr, 'marL');
  if (marL !== undefined) props.marginLeftPx = emuToPx(marL);
  const indent = attrNum(pPr, 'indent');
  if (indent !== undefined) props.indentPx = emuToPx(indent);

  const bullet = parseBullet(pPr, ctx);
  if (bullet) props.bullet = bullet;

  const lnSpc = child(pPr, 'lnSpc');
  if (lnSpc) {
    const pct = attrNum(child(lnSpc, 'spcPct'), 'val');
    const pts = attrNum(child(lnSpc, 'spcPts'), 'val');
    if (pct !== undefined) props.lineSpacingPct = pct / 100000;
    else if (pts !== undefined) props.lineSpacingPt = pts / 100;
  }
  const spcBef = attrNum(child(child(pPr, 'spcBef'), 'spcPts'), 'val');
  if (spcBef !== undefined) props.spaceBeforePt = spcBef / 100;
  const spcAft = attrNum(child(child(pPr, 'spcAft'), 'spcPts'), 'val');
  if (spcAft !== undefined) props.spaceAfterPt = spcAft / 100;

  return props;
}

function parseBullet(pPr: XmlNode, ctx: ColorContext): Bullet | undefined {
  if (child(pPr, 'buNone')) return { type: 'none' };
  const buClr = resolveContainerColor(child(pPr, 'buClr'), ctx);
  const buFont = attr(child(pPr, 'buFont'), 'typeface');
  const buSzPct = attrNum(child(pPr, 'buSzPct'), 'val');

  const buChar = child(pPr, 'buChar');
  if (buChar) {
    const bullet: Bullet = { type: 'char', char: attr(buChar, 'char') ?? '•' };
    if (buFont) bullet.font = buFont;
    if (buClr) bullet.color = buClr;
    if (buSzPct !== undefined) bullet.sizePct = buSzPct / 100000;
    return bullet;
  }
  const buAutoNum = child(pPr, 'buAutoNum');
  if (buAutoNum) {
    const bullet: Bullet = {
      type: 'number',
      scheme: attr(buAutoNum, 'type') ?? 'arabicPeriod',
    };
    const startAt = attrNum(buAutoNum, 'startAt');
    if (startAt !== undefined) bullet.startAt = startAt;
    if (buClr) bullet.color = buClr;
    if (buSzPct !== undefined) bullet.sizePct = buSzPct / 100000;
    return bullet;
  }
  return undefined;
}

function merge<T extends object>(base: T, over: Partial<T>): T {
  const out = { ...base };
  for (const [k, v] of Object.entries(over)) {
    if (v !== undefined) (out as Record<string, unknown>)[k] = v;
  }
  return out;
}

export class TextStyleChain {
  /** Ordered most-specific-first list of lstStyle-like container nodes. */
  private readonly chain: XmlNode[];
  private readonly ctx: ColorContext;

  constructor(chain: XmlNode[], ctx: ColorContext) {
    this.chain = chain.filter(Boolean);
    this.ctx = ctx;
  }

  /** Effective run props at a level, with an optional explicit rPr on top. */
  runProps(level: number, rPr?: XmlNode): ResolvedRunProps {
    // Build least-specific first so later merges win.
    let props: ResolvedRunProps = {
      sizePt: 18,
      color: { hex: this.defaultTextHex() },
      font: this.ctx.theme.minorFont,
    };
    for (let i = this.chain.length - 1; i >= 0; i--) {
      const lvl = child(this.chain[i], lvlKey(level));
      props = merge(props, runPropsFrom(child(lvl, 'defRPr'), this.ctx));
    }
    props = merge(props, runPropsFrom(rPr, this.ctx));
    return props;
  }

  /** Effective paragraph props at a level, with an optional explicit pPr on top. */
  paraProps(level: number, pPr?: XmlNode): ResolvedParaProps {
    let props: ResolvedParaProps = {};
    for (let i = this.chain.length - 1; i >= 0; i--) {
      const lvl = child(this.chain[i], lvlKey(level));
      props = merge(props, paraPropsFrom(lvl, this.ctx));
    }
    props = merge(props, paraPropsFrom(pPr, this.ctx));
    return props;
  }

  private defaultTextHex(): string {
    const key = this.ctx.clrMap['tx1'] ?? 'dk1';
    return this.ctx.theme.colors[key] ?? '000000';
  }
}
