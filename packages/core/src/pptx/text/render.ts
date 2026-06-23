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

export function renderTextBody(body: TextBody, _deps: RenderDeps, flow = false): HTMLDivElement {
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
  }

  // Track auto-number counters per level.
  const counters: number[] = [];
  for (const para of body.paragraphs) {
    box.appendChild(renderParagraph(para, body, counters));
  }
  return box;
}

/** CSS vertical-align value for a text body's anchor (used by table cells). */
export function anchorToValign(anchor: TextBody['anchor']): string {
  return anchor === 'ctr' ? 'middle' : anchor === 'bottom' ? 'bottom' : 'top';
}

function renderParagraph(para: Paragraph, body: TextBody, counters: number[]): HTMLDivElement {
  const p = document.createElement('div');
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
  if (para.marginLeftPx !== undefined) p.style.marginLeft = `${para.marginLeftPx}px`;
  if (para.indentPx !== undefined) {
    // A hanging indent (negative `indent`) makes room left of `marL` for a
    // bullet. PowerPoint never renders the first line left of the paragraph's
    // own left edge, so clamp it — otherwise a marL=0/negative-indent paragraph
    // (e.g. a tab-stop "hang" with no bullet) spills out and clips its start.
    const ml = para.marginLeftPx ?? 0;
    p.style.textIndent = `${Math.max(para.indentPx, -ml)}px`;
  }
  if (para.spaceBeforePt !== undefined) p.style.marginTop = `${ptToPx(para.spaceBeforePt)}px`;
  if (para.spaceAfterPt !== undefined) p.style.marginBottom = `${ptToPx(para.spaceAfterPt)}px`;
  if (para.lineSpacingPct !== undefined) p.style.lineHeight = `${para.lineSpacingPct}`;
  else if (para.lineSpacingPt !== undefined) p.style.lineHeight = `${ptToPx(para.lineSpacingPt)}px`;

  const bulletStr = bulletText(para.bullet, para.level, counters);
  if (bulletStr) {
    const b = document.createElement('span');
    b.textContent = bulletStr + ' ';
    if (para.bullet && 'color' in para.bullet && para.bullet.color) {
      b.style.color = colorToCss(para.bullet.color);
    }
    if (para.bullet && para.bullet.type === 'char' && para.bullet.font) {
      b.style.fontFamily = para.bullet.font;
    }
    p.appendChild(b);
  }

  if (para.runs.length === 0) {
    // Preserve an empty line's height with a zero-width space.
    p.appendChild(document.createTextNode('​'));
  }
  for (const run of para.runs) {
    p.appendChild(renderRun(run, body.fontScale));
  }
  return p;
}

function bulletText(bullet: Bullet | undefined, level: number, counters: number[]): string {
  if (!bullet || bullet.type === 'none') {
    // No explicit bullet: reset deeper counters but render nothing.
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

function renderRun(run: TextRun, fontScale: number | undefined): HTMLElement {
  const span = document.createElement('span');
  span.textContent = run.text;
  if (run.bold) span.style.fontWeight = 'bold';
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
  if (run.font) span.style.fontFamily = `"${run.font}", Arial, Helvetica, sans-serif`;
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
