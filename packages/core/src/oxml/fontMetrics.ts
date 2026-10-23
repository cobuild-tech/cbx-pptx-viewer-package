/**
 * Font metrics for faithful line layout without the real fonts.
 *
 * Word's line breaks — and so its page breaks — depend on two numbers per font:
 * how wide the glyphs run, and how tall a "single" line is (the font's ascent +
 * descent + line gap, *not* a fixed 1.15 or 1.2). Office documents name fonts
 * (Calibri, Cambria, Aptos…) that a browser outside Windows rarely has, and the
 * stock substitute (Helvetica/Arial) is ~10% wider than Calibri with a shorter
 * line, so a 10-page document lays out as 11.
 *
 * Two remedies, layered:
 *  1. oxml/webfonts.ts fetches a metric-compatible clone (Carlito for Calibri…)
 *     where one exists, registered under the document's own font name.
 *  2. Until it arrives — or where none exists, or fetching is off — the CSS
 *     stack falls through to a *local* fallback face: a system font (Arial,
 *     Times New Roman, Courier New) scaled with `size-adjust` to the original's
 *     average glyph width and given its vertical metrics via the
 *     `*-override` descriptors. The glyphs are wrong but the lines break and
 *     stack where Word's do — the technique font-loading libraries use to stop
 *     layout shift.
 *
 * The numbers below were measured in Chrome from the Microsoft font files (the
 * fonts themselves are not redistributed): `w` is the mean advance per
 * character in em over representative English prose, `lh` the "normal" line
 * height in em (what Word calls single spacing), `asc`/`desc` the ascent and
 * descent split of it.
 */
import { installStyleSheet } from './stylesheet.js';

type Generic = 'sans-serif' | 'serif' | 'monospace';

interface Metrics {
  /** Mean advance per character, em. Absent = no fallback face (installed font). */
  w?: number;
  /** Mean advance of the bold face, em, when measured separately. */
  wBold?: number;
  /** "Normal" line height, em — Word's single line spacing for this font. */
  lh: number;
  asc?: number;
  desc?: number;
  generic: Generic;
}

/** Fallback base per generic: the system face scaled to imitate the original. */
const BASE: Record<Generic, { regular: string[]; bold: string[]; italic: string[]; w: number; wBold: number }> = {
  'sans-serif': {
    regular: ['Arial', 'ArialMT', 'Helvetica', 'Liberation Sans', 'Arimo'],
    bold: ['Arial Bold', 'Arial-BoldMT', 'Helvetica Bold', 'Liberation Sans Bold', 'Arimo Bold'],
    italic: ['Arial Italic', 'Arial-ItalicMT', 'Helvetica Oblique', 'Liberation Sans Italic'],
    w: 0.4618,
    wBold: 0.492,
  },
  serif: {
    regular: ['Times New Roman', 'TimesNewRomanPSMT', 'Times', 'Liberation Serif', 'Tinos'],
    bold: ['Times New Roman Bold', 'TimesNewRomanPS-BoldMT', 'Times Bold', 'Liberation Serif Bold'],
    italic: ['Times New Roman Italic', 'TimesNewRomanPS-ItalicMT', 'Times Italic', 'Liberation Serif Italic'],
    w: 0.4237,
    wBold: 0.4499,
  },
  monospace: {
    regular: ['Courier New', 'CourierNewPSMT', 'Courier', 'Liberation Mono', 'Cousine'],
    bold: ['Courier New Bold', 'CourierNewPS-BoldMT', 'Courier Bold', 'Liberation Mono Bold'],
    italic: ['Courier New Italic', 'CourierNewPS-ItalicMT', 'Liberation Mono Italic'],
    w: 0.6001,
    wBold: 0.6001,
  },
};

/** Keyed by lowercased family name. */
const METRICS: Record<string, Metrics> = {
  // Microsoft Office families (measured from the Office font files).
  calibri: { w: 0.4164, wBold: 0.4261, lh: 1.2205, asc: 0.952, desc: 0.269, generic: 'sans-serif' },
  'calibri light': { w: 0.4111, lh: 1.2205, asc: 0.952, desc: 0.269, generic: 'sans-serif' },
  cambria: { w: 0.439, wBold: 0.4703, lh: 1.172, asc: 0.95, desc: 0.222, generic: 'serif' },
  aptos: { w: 0.4386, wBold: 0.4619, lh: 1.2205, asc: 0.939, desc: 0.282, generic: 'sans-serif' },
  'aptos display': { w: 0.4386, wBold: 0.4619, lh: 1.2205, asc: 0.939, desc: 0.282, generic: 'sans-serif' },
  corbel: { w: 0.4245, lh: 1.2075, asc: 0.744, desc: 0.256, generic: 'sans-serif' },
  constantia: { w: 0.4486, lh: 1.2205, asc: 0.751, desc: 0.249, generic: 'serif' },
  consolas: { w: 0.5498, lh: 1.1709, asc: 0.92, desc: 0.251, generic: 'monospace' },
  'century gothic': { w: 0.496, lh: 1.226, asc: 1.006, desc: 0.22, generic: 'sans-serif' },
  'franklin gothic book': { w: 0.432, lh: 1.134, asc: 0.917, desc: 0.217, generic: 'sans-serif' },
  'gill sans mt': { w: 0.4243, lh: 1.1595, asc: 0.929, desc: 0.23, generic: 'sans-serif' },
  'tw cen mt': { w: 0.4137, lh: 1.089, asc: 0.856, desc: 0.233, generic: 'sans-serif' },
  'lucida sans': { w: 0.501, lh: 1.178, asc: 0.967, desc: 0.211, generic: 'sans-serif' },
  garamond: { w: 0.4093, lh: 1.125, asc: 0.862, desc: 0.263, generic: 'serif' },
  'book antiqua': { w: 0.4624, lh: 1.2055, asc: 0.923, desc: 0.282, generic: 'serif' },
  'palatino linotype': { w: 0.4623, lh: 1.3495, asc: 1.05, desc: 0.299, generic: 'serif' },
  'century schoolbook': { w: 0.4856, lh: 1.2025, asc: 0.986, desc: 0.216, generic: 'serif' },
  'bookman old style': { w: 0.5089, lh: 1.174, asc: 0.942, desc: 0.232, generic: 'serif' },
  rockwell: { w: 0.4716, lh: 1.1745, asc: 0.946, desc: 0.229, generic: 'serif' },
  // Commonly installed families: only their line height is needed.
  arial: { lh: 1.15, generic: 'sans-serif' },
  helvetica: { lh: 1.15, generic: 'sans-serif' },
  'times new roman': { lh: 1.15, generic: 'serif' },
  'courier new': { lh: 1.133, generic: 'monospace' },
  georgia: { lh: 1.136, generic: 'serif' },
  verdana: { lh: 1.2155, generic: 'sans-serif' },
  tahoma: { lh: 1.207, generic: 'sans-serif' },
  'trebuchet ms': { lh: 1.161, generic: 'sans-serif' },
  'segoe ui': { lh: 1.33, generic: 'sans-serif' },
  'source sans pro': { lh: 1.257, generic: 'sans-serif' },
  // Symbol fonts: their glyphs are mapped to Unicode, but a bullet drawn in
  // one still sets its line's height by the symbol font's metrics.
  // (Word's effective Symbol line is 1.205 — measured from Word's own layout,
  // a hair under the font's 1.225 hhea height.)
  symbol: { lh: 1.205, generic: 'sans-serif' },
  wingdings: { lh: 1.11, generic: 'sans-serif' },
};

/**
 * Legacy and abbreviated names Word still writes into documents (e.g. "Tms Rmn"
 * from Word 2.0-era templates), mapped to the face Word actually draws.
 */
const ALIASES: Record<string, string> = {
  'tms rmn': 'Times New Roman',
  'times new roman (body cs)': 'Times New Roman',
  helv: 'Arial',
  'ms sans serif': 'Arial',
  'ms serif': 'Times New Roman',
  'courier new (body cs)': 'Courier New',
  'calibri (body)': 'Calibri',
  'calibri light (headings)': 'Calibri Light',
};

/** Single line height used when the font is unknown, in em. */
export const DEFAULT_LINE_HEIGHT = 1.17;

/**
 * Word's single line height for East Asian faces, in em — measured from
 * Word's own layout (11pt, single spacing), which runs well above what the
 * fonts' tables give a browser, so these cannot be derived. Keyed lowercase.
 */
const EAST_ASIAN_LINE: Record<string, number> = {
  simsun: 1.318,
  nsimsun: 1.318,
  pmingliu: 1.318,
  mingliu: 1.318,
  dengxian: 1.368,
  'microsoft yahei': 1.727,
  'microsoft jhenghei': 1.727,
  'ms gothic': 1.727,
  'ms pgothic': 1.727,
  'ms mincho': 1.72,
  'ms pmincho': 1.72,
  'yu gothic': 1.72,
  'yu mincho': 1.72,
  'malgun gothic': 1.72,
  meiryo: 1.927,
};

/** East Asian faces Word does not recognise fall back to one about this tall. */
const DEFAULT_EAST_ASIAN_LINE = 1.727;

/** Text that Word sets in a run's East Asian font. */
export const EAST_ASIAN_TEXT = /[\u1100-\u11ff\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/;

/** Word's single line spacing for East Asian characters in a face, as a multiple of size. */
export function eastAsianLineHeightRatio(name: string | undefined): number {
  return (name && EAST_ASIAN_LINE[name.trim().toLowerCase()]) || DEFAULT_EAST_ASIAN_LINE;
}

const SERIF_NAME = /times|georgia|cambria|garamond|minion|antiqua|palatino|serif|roman|constantia|bookman|schoolbook|rockwell|baskerville|century(?! gothic)/i;
const MONO_NAME = /courier|consolas|mono|lucida console|typewriter/i;

/** The family Word really draws for a name written in a document. */
export function canonicalFont(name: string): string {
  const trimmed = name.trim();
  return ALIASES[trimmed.toLowerCase()] ?? trimmed;
}

function genericOf(name: string): Generic {
  const m = METRICS[name.toLowerCase()];
  if (m) return m.generic;
  if (MONO_NAME.test(name)) return 'monospace';
  return SERIF_NAME.test(name) ? 'serif' : 'sans-serif';
}

function fallbackFamily(name: string): string {
  return `cbx-fb-${name.toLowerCase().replace(/[^a-z0-9]+/g, '-')}`;
}

/**
 * A CSS `font-family` stack for a document font: the font itself, then its
 * metric-matched local fallback face (when we know its metrics), then the right
 * generic — so an uninstalled serif never falls back to sans or vice versa.
 */
export function fontStack(name: string, ...alternates: (string | undefined)[]): string {
  const family = canonicalFont(name);
  const parts = [quote(family)];
  const m = METRICS[family.toLowerCase()];
  if (m?.w !== undefined) parts.push(quote(fallbackFamily(family)));
  // Faces for other scripts (an East Asian font for CJK characters): the
  // browser falls through to them per character, as Word switches fonts.
  for (const alt of alternates) if (alt) parts.push(quote(canonicalFont(alt)));
  parts.push(genericOf(family));
  return parts.join(', ');
}

/** Word's single line spacing for a font, as a multiple of its size. */
export function lineHeightRatio(name: string | undefined): number {
  if (!name) return DEFAULT_LINE_HEIGHT;
  return METRICS[canonicalFont(name).toLowerCase()]?.lh ?? DEFAULT_LINE_HEIGHT;
}

/** @font-face rules for every metric-matched fallback face. */
export function fallbackFontFaceCss(): string {
  const rules: string[] = [];
  const pct = (n: number) => `${(n * 100).toFixed(2)}%`;
  for (const [key, m] of Object.entries(METRICS)) {
    if (m.w === undefined) continue;
    const base = BASE[m.generic];
    const family = fallbackFamily(key);
    const faces: [string, string, string[], number][] = [
      ['normal', 'normal', base.regular, m.w / base.w],
      ['bold', 'normal', base.bold, (m.wBold ?? (m.w * base.wBold) / base.w) / base.wBold],
      ['normal', 'italic', base.italic, m.w / base.w],
    ];
    for (const [weight, style, locals, adjust] of faces) {
      // The override descriptors are scaled by size-adjust, so divide it back
      // out to land on the original's absolute ascent/descent.
      const asc = m.asc ?? m.lh * 0.78;
      const desc = m.desc ?? m.lh * 0.22;
      const scale = m.lh / (asc + desc);
      rules.push(
        `@font-face{font-family:"${family}";font-weight:${weight};font-style:${style};` +
          `src:${locals.map((l) => `local("${l}")`).join(',')};` +
          `size-adjust:${pct(adjust)};ascent-override:${pct((asc * scale) / adjust)};` +
          `descent-override:${pct((desc * scale) / adjust)};line-gap-override:0%}`,
      );
    }
  }
  return rules.join('\n');
}

/**
 * Install the fallback faces into `doc` (ref-counted, shared by every viewer on
 * the page). Returns the disposer.
 */
export function installFallbackFonts(doc: Document): () => void {
  return installStyleSheet(doc, 'cbx-font-fallbacks', fallbackFontFaceCss());
}

function quote(name: string): string {
  return /^[a-z-]+$/i.test(name) ? name : `"${name.replace(/["\\]/g, '\\$&')}"`;
}
