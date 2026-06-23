/**
 * Color resolution: the part of fidelity most decks depend on.
 *
 * A `<a:schemeClr val="accent1">` doesn't carry a color — it references the
 * theme's color scheme, possibly remapped by the slide master's `<p:clrMap>`,
 * and then transformed by modifiers (lumMod/lumOff/shade/tint/alpha...). This
 * module turns any DrawingML color element into a concrete sRGB {@link Color}.
 */
import { child, attr, attrNum, localName, type XmlNode } from '../oxml/xml.js';
import type { Color } from './model.js';

export interface Theme {
  /** Theme color-scheme entries: dk1,lt1,dk2,lt2,accent1..6,hlink,folHlink -> hex. */
  colors: Record<string, string>;
  majorFont: string;
  minorFont: string;
  /**
   * `<a:effectStyleLst>` effect styles from the format scheme, in order. A
   * shape's `effectRef idx="N"` selects effectStyles[N-1] (idx 0 = none).
   */
  effectStyles: XmlNode[];
}

export interface ColorContext {
  theme: Theme;
  /** Slide-master color map: bg1,tx1,bg2,tx2,... -> theme color-scheme key. */
  clrMap: Record<string, string>;
  /** Current placeholder color, substituted for `<a:schemeClr val="phClr">`. */
  phClr?: Color;
}

const PRESET_COLORS: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  cyan: '00FFFF',
  magenta: 'FF00FF',
  gray: '808080',
  grey: '808080',
  darkGray: 'A9A9A9',
  lightGray: 'D3D3D3',
  orange: 'FFA500',
  purple: '800080',
};

const COLOR_TAGS = new Set([
  'srgbClr',
  'schemeClr',
  'sysClr',
  'prstClr',
  'scrgbClr',
  'hslClr',
]);

/** Find the first color element among a container node's children. */
export function findColorEl(node: XmlNode | undefined): XmlNode | undefined {
  if (!node) return undefined;
  return node.children.find((c) => COLOR_TAGS.has(localName(c.name)));
}

export function parseTheme(themeXml: XmlNode | undefined): Theme {
  const fallback: Theme = {
    colors: {
      dk1: '000000',
      lt1: 'FFFFFF',
      dk2: '44546A',
      lt2: 'E7E6E6',
      accent1: '4472C4',
      accent2: 'ED7D31',
      accent3: 'A5A5A5',
      accent4: 'FFC000',
      accent5: '5B9BD5',
      accent6: '70AD47',
      hlink: '0563C1',
      folHlink: '954F72',
    },
    majorFont: 'Calibri Light',
    minorFont: 'Calibri',
    effectStyles: [],
  };
  if (!themeXml) return fallback;

  const elements = child(themeXml, 'themeElements');
  const clrScheme = child(elements, 'clrScheme');
  if (clrScheme) {
    for (const entry of clrScheme.children) {
      const key = localName(entry.name);
      const colorEl = findColorEl(entry);
      const hex = colorEl ? baseHex(colorEl) : undefined;
      if (hex) fallback.colors[key] = hex;
    }
  }

  const fontScheme = child(elements, 'fontScheme');
  const major = attr(child(child(fontScheme, 'majorFont'), 'latin'), 'typeface');
  const minor = attr(child(child(fontScheme, 'minorFont'), 'latin'), 'typeface');
  if (major) fallback.majorFont = major;
  if (minor) fallback.minorFont = minor;

  const effectStyleLst = child(child(elements, 'fmtScheme'), 'effectStyleLst');
  if (effectStyleLst) {
    fallback.effectStyles = effectStyleLst.children.filter(
      (c) => localName(c.name) === 'effectStyle',
    );
  }

  return fallback;
}

/** Base (pre-modifier) hex for a color element, or undefined. */
function baseHex(el: XmlNode): string | undefined {
  switch (localName(el.name)) {
    case 'srgbClr':
      return normalizeHex(attr(el, 'val'));
    case 'sysClr':
      return normalizeHex(attr(el, 'lastClr')) ?? '000000';
    case 'prstClr': {
      const name = attr(el, 'val');
      return name ? PRESET_COLORS[name] ?? '000000' : undefined;
    }
    case 'scrgbClr': {
      const r = pctChannel(attrNum(el, 'r'));
      const g = pctChannel(attrNum(el, 'g'));
      const b = pctChannel(attrNum(el, 'b'));
      return toHex(r, g, b);
    }
    default:
      return undefined;
  }
}

function pctChannel(v: number | undefined): number {
  if (v === undefined) return 0;
  return Math.round((v / 100000) * 255);
}

function normalizeHex(v: string | undefined): string | undefined {
  if (!v) return undefined;
  const h = v.replace(/^#/, '').toUpperCase();
  return /^[0-9A-F]{6}$/.test(h) ? h : undefined;
}

/**
 * Resolve a color element (e.g. `<a:schemeClr val="accent1"><a:lumMod.../></a:schemeClr>`)
 * to a concrete {@link Color}, applying any modifier children.
 */
export function resolveColorEl(
  el: XmlNode | undefined,
  ctx: ColorContext,
): Color | undefined {
  if (!el) return undefined;

  let hex: string | undefined;
  let alpha: number | undefined;

  if (localName(el.name) === 'schemeClr') {
    let key = attr(el, 'val') ?? '';
    if (key === 'phClr' && ctx.phClr) {
      hex = ctx.phClr.hex;
      alpha = ctx.phClr.alpha;
    } else {
      // bg1/tx1/bg2/tx2 are mapped through the master's clrMap.
      if (key in ctx.clrMap) key = ctx.clrMap[key]!;
      // dk1/lt1/dk2/lt2 may still need remapping to themed names.
      hex = ctx.theme.colors[key] ?? ctx.theme.colors[ctx.clrMap[key] ?? ''] ?? '000000';
    }
  } else {
    hex = baseHex(el);
  }

  if (hex === undefined) return undefined;

  let rgb = hexToRgb(hex);
  // Apply modifiers in document order.
  for (const mod of el.children) {
    const name = localName(mod.name);
    const val = attrNum(mod, 'val');
    switch (name) {
      case 'alpha':
        if (val !== undefined) alpha = val / 100000;
        break;
      case 'lumMod':
        if (val !== undefined) rgb = adjustLum(rgb, val / 100000, 'mod');
        break;
      case 'lumOff':
        if (val !== undefined) rgb = adjustLum(rgb, val / 100000, 'off');
        break;
      case 'shade':
        if (val !== undefined) rgb = scaleRgb(rgb, val / 100000);
        break;
      case 'tint':
        if (val !== undefined) rgb = tintRgb(rgb, val / 100000);
        break;
      // satMod / hueMod / gray etc. are approximated as no-ops for now.
    }
  }

  const result: Color = { hex: toHex(rgb.r, rgb.g, rgb.b) };
  if (alpha !== undefined && alpha < 1) result.alpha = alpha;
  return result;
}

/** Resolve the color inside a fill/line container (finds the color child first). */
export function resolveContainerColor(
  container: XmlNode | undefined,
  ctx: ColorContext,
): Color | undefined {
  return resolveColorEl(findColorEl(container), ctx);
}

interface Rgb {
  r: number;
  g: number;
  b: number;
}

function hexToRgb(hex: string): Rgb {
  return {
    r: parseInt(hex.slice(0, 2), 16),
    g: parseInt(hex.slice(2, 4), 16),
    b: parseInt(hex.slice(4, 6), 16),
  };
}

function toHex(r: number, g: number, b: number): string {
  const c = (n: number) =>
    Math.max(0, Math.min(255, Math.round(n)))
      .toString(16)
      .padStart(2, '0')
      .toUpperCase();
  return `${c(r)}${c(g)}${c(b)}`;
}

/** shade: scale RGB toward black by `factor` (0..1). */
function scaleRgb(rgb: Rgb, factor: number): Rgb {
  return { r: rgb.r * factor, g: rgb.g * factor, b: rgb.b * factor };
}

/** tint: blend RGB toward white; `factor` is the weight kept of the original. */
function tintRgb(rgb: Rgb, factor: number): Rgb {
  const mix = (c: number) => c * factor + 255 * (1 - factor);
  return { r: mix(rgb.r), g: mix(rgb.g), b: mix(rgb.b) };
}

/** lumMod/lumOff operate on HSL luminance. */
function adjustLum(rgb: Rgb, val: number, kind: 'mod' | 'off'): Rgb {
  const { h, s, l } = rgbToHsl(rgb);
  const nl = kind === 'mod' ? l * val : Math.min(1, l + val);
  return hslToRgb(h, s, nl);
}

function rgbToHsl(rgb: Rgb): { h: number; s: number; l: number } {
  const r = rgb.r / 255;
  const g = rgb.g / 255;
  const b = rgb.b / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  let h = 0;
  let s = 0;
  const d = max - min;
  if (d !== 0) {
    s = d / (1 - Math.abs(2 * l - 1));
    switch (max) {
      case r:
        h = ((g - b) / d) % 6;
        break;
      case g:
        h = (b - r) / d + 2;
        break;
      default:
        h = (r - g) / d + 4;
    }
    h *= 60;
    if (h < 0) h += 360;
  }
  return { h, s, l };
}

function hslToRgb(h: number, s: number, l: number): Rgb {
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - c / 2;
  let r = 0;
  let g = 0;
  let b = 0;
  if (h < 60) [r, g, b] = [c, x, 0];
  else if (h < 120) [r, g, b] = [x, c, 0];
  else if (h < 180) [r, g, b] = [0, c, x];
  else if (h < 240) [r, g, b] = [0, x, c];
  else if (h < 300) [r, g, b] = [x, 0, c];
  else [r, g, b] = [c, 0, x];
  return { r: (r + m) * 255, g: (g + m) * 255, b: (b + m) * 255 };
}

/** CSS color string for a resolved {@link Color}. */
export function colorToCss(color: Color): string {
  if (color.alpha !== undefined && color.alpha < 1) {
    const rgb = hexToRgb(color.hex);
    return `rgba(${Math.round(rgb.r)},${Math.round(rgb.g)},${Math.round(rgb.b)},${color.alpha.toFixed(3)})`;
  }
  return `#${color.hex}`;
}
