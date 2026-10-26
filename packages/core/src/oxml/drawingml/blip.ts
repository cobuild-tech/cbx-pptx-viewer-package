/**
 * Blip (picture) effects shared by PowerPoint and Word: the recolouring and
 * transparency a `<a:blip>` applies to its image.
 *
 * - `<a:alphaModFix amt>` — the picture's transparency slider.
 * - `<a:clrChange>` — "Set Transparent Color": one colour (usually a logo's
 *   white background) replaced by another, typically fully transparent.
 * - `<a:duotone>` — the Recolor presets: luminance mapped from a dark colour
 *   to a light one.
 *
 * Opacity is plain CSS; the colour effects become one SVG filter, applied in
 * Office's order (colour change, then duotone).
 */
import { child, attrNum, type XmlNode } from '../xml.js';
import { findColorEl, resolveColorEl, type ColorContext } from './color.js';

export interface BlipEffects {
  /** Opacity 0..1 from `<a:alphaModFix>`. */
  alpha?: number;
  /** `<a:clrChange>`: pixels of `from` become `to` at `toAlpha` (0 = transparent). */
  clrChange?: { from: string; to: string; toAlpha: number };
  /** `<a:duotone>`: [dark, light] hex colours. */
  duotone?: [string, string];
}

/** Read the effects a `<a:blip>` applies, or undefined when it applies none. */
export function parseBlipEffects(blip: XmlNode | undefined, ctx: ColorContext): BlipEffects | undefined {
  if (!blip) return undefined;
  const out: BlipEffects = {};

  const alphaModFix = child(blip, 'alphaModFix');
  if (alphaModFix) {
    const amt = attrNum(alphaModFix, 'amt') ?? 100000;
    if (amt < 100000) out.alpha = Math.max(0, amt / 100000);
  }

  const clrChange = child(blip, 'clrChange');
  if (clrChange) {
    const from = resolveColorEl(findColorEl(child(clrChange, 'clrFrom')), ctx);
    const to = resolveColorEl(findColorEl(child(clrChange, 'clrTo')), ctx);
    if (from && to) out.clrChange = { from: from.hex, to: to.hex, toAlpha: to.alpha ?? 1 };
  }

  const duotone = child(blip, 'duotone');
  if (duotone) {
    // Its two colour children, in order: darks, then lights.
    const [dark, light] = duotone.children
      .map((c) => resolveColorEl(c, ctx)?.hex)
      .filter((hex): hex is string => hex !== undefined);
    if (dark && light) out.duotone = [dark, light];
  }

  return out.alpha !== undefined || out.clrChange || out.duotone ? out : undefined;
}

const SVG_NS = 'http://www.w3.org/2000/svg';

/**
 * Register (once per document and effect set) the SVG filter for a blip's
 * colour effects and return its id, or undefined when there are none. CSS
 * `filter: url()` only reaches filters in the same document, so the
 * definition lives in a hidden <svg> in the body.
 */
export function blipFilter(doc: Document, effects: BlipEffects): string | undefined {
  const { clrChange, duotone } = effects;
  if (!clrChange && !duotone) return undefined;
  const key = [
    clrChange ? `c${clrChange.from}-${clrChange.to}-${Math.round(clrChange.toAlpha * 100)}` : '',
    duotone ? `d${duotone[0]}-${duotone[1]}` : '',
  ].join('');
  const id = `cbx-blip-${key}`.toLowerCase();
  if (doc.getElementById(id)) return id;

  const make = (name: string, attrs: Record<string, string>) => {
    const node = doc.createElementNS(SVG_NS, name);
    for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, v);
    return node;
  };
  const ch = (hex: string, i: number) => parseInt(hex.slice(i * 2, i * 2 + 2), 16) / 255;
  const funcs = ['feFuncR', 'feFuncG', 'feFuncB'];

  const svg = make('svg', { width: '0', height: '0', 'aria-hidden': 'true' });
  svg.style.position = 'absolute';
  const filter = make('filter', { id, 'color-interpolation-filters': 'sRGB' });
  let input = 'SourceGraphic';

  if (clrChange) {
    // Mark each channel 1 where it differs from the key colour (a 256-step
    // table, so only the key's own step maps to 0), sum the channels into
    // alpha, and keep the source only where some channel differed.
    const diff = make('feComponentTransfer', { in: input, result: 'cbxDiff' });
    funcs.forEach((f, i) => {
      const key = Math.round(ch(clrChange.from, i) * 255);
      const table = Array.from({ length: 256 }, (_, v) => (v === key ? '0' : '1')).join(' ');
      diff.appendChild(make(f, { type: 'discrete', tableValues: table }));
    });
    filter.appendChild(diff);
    filter.appendChild(
      make('feColorMatrix', {
        in: 'cbxDiff',
        type: 'matrix',
        values: '0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 1 1 0 0',
        result: 'cbxKeep',
      }),
    );
    filter.appendChild(make('feComposite', { in: input, in2: 'cbxKeep', operator: 'in', result: 'cbxKept' }));
    input = 'cbxKept';
    if (clrChange.toAlpha > 0) {
      // An opaque replacement colour fills where the key colour was.
      filter.appendChild(
        make('feFlood', { 'flood-color': `#${clrChange.to}`, 'flood-opacity': `${clrChange.toAlpha}`, result: 'cbxTo' }),
      );
      filter.appendChild(make('feComposite', { in: 'cbxTo', in2: 'cbxKeep', operator: 'out', result: 'cbxToOut' }));
      filter.appendChild(make('feComposite', { in: 'cbxToOut', in2: 'SourceAlpha', operator: 'in', result: 'cbxFill' }));
      const merge = make('feMerge', { result: 'cbxChanged' });
      merge.appendChild(make('feMergeNode', { in: 'cbxFill' }));
      merge.appendChild(make('feMergeNode', { in: 'cbxKept' }));
      filter.appendChild(merge);
      input = 'cbxChanged';
    }
  }

  if (duotone) {
    const [dark, light] = duotone;
    filter.appendChild(
      make('feColorMatrix', {
        in: input,
        type: 'matrix',
        values: '0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0 0 0 1 0',
      }),
    );
    const transfer = make('feComponentTransfer', {});
    funcs.forEach((f, i) =>
      transfer.appendChild(make(f, { type: 'table', tableValues: `${ch(dark, i).toFixed(3)} ${ch(light, i).toFixed(3)}` })),
    );
    filter.appendChild(transfer);
  }

  svg.appendChild(filter);
  (doc.body ?? doc.documentElement).appendChild(svg);
  return id;
}
