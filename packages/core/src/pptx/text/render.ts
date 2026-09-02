/**
 * Text-body renderer: {@link TextBody} -> HTML.
 *
 * In the default (shape) mode the box is absolutely positioned and uses flex for
 * vertical anchoring. In `flow` mode (table cells) it's a normal-flow block with
 * the insets as padding, so the cell — and its table row — grows to fit the
 * content, just like PowerPoint.
 */
import type { TextBody, Paragraph, TextRun, Bullet } from '../model.js';
import { ptToPx } from '../../oxml/units.js';
import { colorToCss } from '../color.js';
import type { RenderDeps } from '../render/primitives.js';
import { EDIT_ATTR } from '../../oxml/edit/attrs.js';

export { EDIT_ATTR };

/**
 * PowerPoint's "single" (100%) line spacing is ~1.2x the font size — it derives
 * line height from the font's ascent+descent but ignores the font's full
 * line-gap. CSS `line-height: normal` does NOT: it's font-dependent and looser
 * for fonts like Open Sans (~1.36), which makes text overflow its box. So we
 * never leave line-height at `normal`; percentage spacing (spcPct) is taken
 * relative to this baseline, and the implicit default is single spacing.
 */
const SINGLE_LINE_HEIGHT = 1.2;

export function renderTextBody(body: TextBody, deps: RenderDeps, flow = false): HTMLDivElement {
  const box = document.createElement('div');
  box.style.boxSizing = 'border-box';
  if (flow) {
    box.style.padding = `${body.insets.t}px ${body.insets.r}px ${body.insets.b}px ${body.insets.l}px`;
  } else {
    box.style.position = 'absolute';
    box.style.left = `${body.insets.l}px`;
    box.style.top = `${body.insets.t}px`;
    box.style.right = `${body.insets.r}px`;
    box.style.bottom = `${body.insets.b}px`;
    box.style.display = 'flex';
    box.style.flexDirection = 'column';
    box.style.justifyContent =
      body.anchor === 'ctr' ? 'center' : body.anchor === 'bottom' ? 'flex-end' : 'flex-start';
    // PowerPoint shows text that overflows its box (the "do not autofit" default)
    // rather than clipping it; only the slide edge clips.
    box.style.overflow = 'visible';
    // Inside a non-uniformly scaled group, the ancestor CSS scale (sx,sy) would
    // squish glyphs. PowerPoint keeps text aspect-correct and reflows in the
    // scaled box, with the font following the vertical scale. Counter the
    // horizontal squish with scaleX(sy/sx) and widen the layout box by the
    // inverse so wrapping still happens at the geometric (scaled) width.
    const gs = deps.groupScale;
    if (gs && gs.sx > 0 && gs.sy > 0 && Math.abs(gs.sx - gs.sy) > 1e-4) {
      const k = gs.sy / gs.sx;
      box.style.right = 'auto';
      box.style.width = `calc((100% - ${body.insets.l + body.insets.r}px) / ${k})`;
      box.style.transform = `scaleX(${k})`;
      box.style.transformOrigin = 'left center';
    }
  }

  if (deps.edit?.editable(body)) {
    box.setAttribute(EDIT_ATTR.body, deps.edit.key(body));
    // Editable and *being edited* are different states: the box is addressable
    // as soon as the slide is editable, but typing into it only starts once the
    // user has entered it, so a plain click can select the shape instead.
    // Set as an attribute, not the property: the viewer finds the box it just
    // opened with an attribute selector, and the two do not reflect in every
    // DOM implementation.
    if (deps.edit.textEditing?.(body) ?? true) box.setAttribute('contenteditable', 'true');
    // Without this the browser inserts <b>/<font> markup on its own shortcuts,
    // which reconciliation would have to unpick; the toolbar drives formatting.
    box.spellcheck = false;
    // Focus/hover outlines come from the injected editor stylesheet — an inline
    // `outline: none` here would beat it.
  }

  // Track auto-number counters per level.
  const counters: number[] = [];
  for (const para of body.paragraphs) {
    box.appendChild(renderParagraph(para, body, counters, deps));
  }
  return box;
}

/** CSS vertical-align value for a text body's anchor (used by table cells). */
export function anchorToValign(anchor: TextBody['anchor']): string {
  return anchor === 'ctr' ? 'middle' : anchor === 'bottom' ? 'bottom' : 'top';
}

function renderParagraph(
  para: Paragraph,
  body: TextBody,
  counters: number[],
  deps: RenderDeps,
): HTMLDivElement {
  const p = document.createElement('div');
  if (deps.edit?.editable(body)) p.setAttribute(EDIT_ATTR.para, deps.edit.key(para));
  p.style.whiteSpace = body.wrap ? 'pre-wrap' : 'pre';
  p.style.margin = '0';
  // Keep natural height when text exceeds the box; PowerPoint clips/overflows
  // rather than letting flexbox compress paragraphs on top of each other.
  p.style.flexShrink = '0';
  p.style.boxSizing = 'border-box';
  if (para.align) {
    p.style.textAlign =
      para.align === 'ctr' ? 'center' : para.align === 'r' ? 'right' : para.align === 'just' ? 'justify' : 'left';
  }
  if (para.spaceBeforePt !== undefined) p.style.marginTop = `${ptToPx(para.spaceBeforePt)}px`;
  if (para.spaceAfterPt !== undefined) p.style.marginBottom = `${ptToPx(para.spaceAfterPt)}px`;
  // Anchor the line-box "strut" to the paragraph's own text size. Without this
  // the block inherits the ambient (browser default ~16px) font size, so a
  // unitless line-height inflates every line beyond the actual text height and
  // the paragraph overflows its box. Empty paragraphs also need this for height.
  if (para.defaultSizePt !== undefined) {
    const sz = body.fontScale ? para.defaultSizePt * body.fontScale : para.defaultSizePt;
    p.style.fontSize = `${ptToPx(sz)}px`;
  }
  if (para.lineSpacingPt !== undefined) {
    // Exact point line spacing: an absolute height, unaffected by autofit reduction.
    p.style.lineHeight = `${ptToPx(para.lineSpacingPt)}px`;
  } else {
    // Percentage spacing (default 100%), less any autofit line-spacing reduction,
    // measured against PowerPoint's single-spacing baseline (~1.2x font size).
    const pct = para.lineSpacingPct ?? 1;
    const reduced = Math.max(0, pct - (body.lnSpcReductionPct ?? 0));
    p.style.lineHeight = `${reduced * SINGLE_LINE_HEIGHT}`;
  }

  const marL = para.marginLeftPx ?? 0;
  const indent = para.indentPx ?? 0;
  const bulletStr = bulletText(para.bullet, para.level, counters);

  if (bulletStr && indent < 0) {
    // Hanging-indent bullet: the bullet occupies a fixed-width gutter (-indent)
    // and the text — including wrapped lines — aligns at marL, so the bullet
    // glyph's own width never shifts the text. Matches PowerPoint's layout,
    // where a tab after the bullet snaps the text to marL.
    p.style.marginLeft = `${marL}px`;
    const b = bulletSpan(bulletStr, para, body.fontScale);
    b.style.display = 'inline-block';
    b.style.width = `${-indent}px`;
    b.style.marginLeft = `${indent}px`;
    p.appendChild(b);
  } else {
    if (para.marginLeftPx !== undefined) p.style.marginLeft = `${marL}px`;
    // Clamp so the first line never starts left of the paragraph's own edge.
    if (para.indentPx !== undefined) p.style.textIndent = `${Math.max(indent, -marL)}px`;
    if (bulletStr) p.appendChild(bulletSpan(bulletStr + ' ', para, body.fontScale));
  }

  if (para.runs.length === 0) {
    // Preserve an empty line's height with a zero-width space.
    p.appendChild(document.createTextNode('​'));
  }
  for (const run of para.runs) {
    p.appendChild(renderRun(run, body.fontScale, deps));
  }
  return p;
}

/**
 * The bullet or number drawn in front of a paragraph.
 *
 * DrawingML's defaults for a marker are `buClrTx`, `buSzTx` and `buFontTx` —
 * *take it from the text* — so the glyph follows the paragraph's first run
 * rather than the box's ambient style: a bullet in front of 36pt red text is
 * 36pt and red. An explicit `<a:buClr>`, `<a:buSzPct>` or `<a:buFont>` on the
 * paragraph overrides that, and `buSzPct` scales the run's size rather than
 * replacing it.
 */
function bulletSpan(
  text: string,
  para: Paragraph,
  fontScale: number | undefined,
): HTMLSpanElement {
  const bullet = para.bullet;
  const lead = para.runs[0];
  const b = document.createElement('span');
  b.textContent = text;
  // A bullet is generated decoration, not content — it must not be typed into
  // or deleted, and reconciliation skips it. Set as an attribute so the DOM
  // agrees with the property in every implementation.
  b.setAttribute('contenteditable', 'false');

  const color = (bullet && 'color' in bullet && bullet.color) || lead?.color;
  if (color) b.style.color = colorToCss(color);

  const sizePt = lead?.sizePt ?? para.defaultSizePt;
  if (sizePt !== undefined) {
    const pct = bullet && 'sizePct' in bullet && bullet.sizePct ? bullet.sizePct : 1;
    b.style.fontSize = `${ptToPx((fontScale ? sizePt * fontScale : sizePt) * pct)}px`;
  }

  // A char bullet's own typeface matters (the glyph has to exist in it); a
  // number is drawn in the text's face.
  const font = (bullet?.type === 'char' ? bullet.font : undefined) ?? lead?.font;
  if (font) {
    b.style.fontFamily = `"${splitFontWeight(font).family}", Arial, Helvetica, sans-serif`;
  }
  return b;
}

function bulletText(bullet: Bullet | undefined, level: number, counters: number[]): string {
  if (!bullet || bullet.type === 'none') {
    // Nothing to draw, and the counters are left as they are: a paragraph with
    // no bullet in the middle of a numbered run does not restart the numbering,
    // which is what PowerPoint does too.
    return '';
  }
  if (bullet.type === 'char') return bullet.char;
  // Auto-numbered: maintain a per-level counter.
  counters[level] = (counters[level] ?? (bullet.startAt ?? 1) - 1) + 1;
  for (let i = level + 1; i < counters.length; i++) counters[i] = 0;
  const n = counters[level];
  if (bullet.scheme.includes('Paren')) return `${n})`;
  if (bullet.scheme.includes('alphaUc')) return `${String.fromCharCode(64 + n)}.`;
  if (bullet.scheme.includes('alphaLc')) return `${String.fromCharCode(96 + n)}.`;
  return `${n}.`;
}

function renderRun(run: TextRun, fontScale: number | undefined, deps: RenderDeps): HTMLElement {
  const span = document.createElement('span');
  span.textContent = run.text;
  if (deps.edit) {
    span.setAttribute(EDIT_ATTR.run, deps.edit.key(run));
    // A field's text is regenerated by PowerPoint, so it is shown but locked.
    if (deps.edit.isField(run)) span.setAttribute('contenteditable', 'false');
  }
  if (run.italic) span.style.fontStyle = 'italic';
  const decorations: string[] = [];
  if (run.underline) decorations.push('underline');
  if (run.strike) decorations.push('line-through');
  if (decorations.length) span.style.textDecoration = decorations.join(' ');
  if (run.sizePt !== undefined) {
    const pt = fontScale ? run.sizePt * fontScale : run.sizePt;
    span.style.fontSize = `${ptToPx(pt)}px`;
  }
  if (run.color) span.style.color = colorToCss(run.color);
  // A weight-suffixed family ("Open Sans Light") is the base family at a CSS
  // weight, not a distinct family name — split it so the base webfont applies.
  let weight = run.bold ? 700 : undefined;
  if (run.font) {
    const { family, weight: w } = splitFontWeight(run.font);
    span.style.fontFamily = `"${family}", Arial, Helvetica, sans-serif`;
    if (w !== undefined && !run.bold) weight = w;
  }
  if (weight !== undefined) span.style.fontWeight = String(weight);
  if (run.highlight) span.style.backgroundColor = colorToCss(run.highlight);
  if (run.letterSpacingPt) span.style.letterSpacing = `${ptToPx(run.letterSpacingPt)}px`;
  if (run.caps === 'all') span.style.textTransform = 'uppercase';
  else if (run.caps === 'small') span.style.fontVariant = 'small-caps';
  if (run.baseline) {
    span.style.verticalAlign = run.baseline > 0 ? 'super' : 'sub';
    span.style.fontSize = span.style.fontSize || 'smaller';
  }

  if (run.hyperlink) {
    const a = document.createElement('a');
    a.href = run.hyperlink;
    a.target = '_blank';
    a.rel = 'noopener noreferrer';
    a.style.color = 'inherit';
    a.appendChild(span);
    return a;
  }
  return span;
}

/** Trailing weight words PowerPoint bakes into a family name (longest first). */
const WEIGHT_SUFFIXES: Array<[RegExp, number]> = [
  [/\s+(extra\s?light|ultra\s?light)$/i, 200],
  [/\s+(semi\s?bold|demi\s?bold)$/i, 600],
  [/\s+(extra\s?bold|ultra\s?bold)$/i, 800],
  [/\s+(thin|hairline)$/i, 100],
  [/\s+light$/i, 300],
  [/\s+medium$/i, 500],
  [/\s+(black|heavy)$/i, 900],
  [/\s+(bold)$/i, 700],
  [/\s+(regular|normal|book)$/i, 400],
];

/**
 * Split a typeface name like "Open Sans Light" into the base family and a CSS
 * weight. PowerPoint treats weight variants as distinct families, but the web
 * exposes them as one family at different `font-weight`s — so normalizing lets a
 * single loaded webfont (e.g. "Open Sans") cover Light/Semibold/etc.
 */
export function splitFontWeight(font: string): { family: string; weight?: number } {
  const name = font.trim();
  for (const [re, weight] of WEIGHT_SUFFIXES) {
    if (re.test(name)) return { family: name.replace(re, '').trim(), weight };
  }
  return { family: name };
}
