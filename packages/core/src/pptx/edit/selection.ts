/**
 * Selection formatting.
 *
 * `document.execCommand` is deprecated and each browser invents its own markup
 * for it (`<b>`, `<font>`, inline styles), which reconciliation would then have
 * to reverse-engineer. Instead the toolbar wraps the selection in one span that
 * carries the intended {@link RunFormat} as data, and paints the equivalent CSS
 * so the change is visible immediately. Reconciliation reads the data back —
 * the CSS is only ever presentation.
 */
import { ptToPx } from '../../oxml/units.js';
import type { TextRun } from '../model.js';
import { EDIT_ATTR } from '../text/render.js';
import { mergeFormat, type RunFormat } from './format.js';
import type { Resolver } from './reconcile.js';

/** Paint a format onto an element so the user sees it before the commit. */
function paint(el: HTMLElement, f: RunFormat): void {
  if (f.bold !== undefined) el.style.fontWeight = f.bold ? '700' : '400';
  if (f.italic !== undefined) el.style.fontStyle = f.italic ? 'italic' : 'normal';
  if (f.sizePt !== undefined) el.style.fontSize = `${ptToPx(f.sizePt)}px`;
  if (f.colorHex !== undefined) el.style.color = `#${f.colorHex.replace(/^#/, '')}`;
  if (f.font !== undefined) el.style.fontFamily = `"${f.font}", Arial, Helvetica, sans-serif`;
  const decorations: string[] = [];
  if (f.underline) decorations.push('underline');
  if (f.strike) decorations.push('line-through');
  el.style.textDecoration = decorations.length ? decorations.join(' ') : '';
}

/** The editable text-body element a node sits inside, if any. */
export function bodyElementOf(node: Node | null, root: Element): HTMLElement | null {
  let el: Node | null = node;
  while (el && el !== root) {
    if (el.nodeType === 1 && (el as Element).hasAttribute(EDIT_ATTR.body)) return el as HTMLElement;
    el = el.parentNode;
  }
  return null;
}

/**
 * Wrap the current selection in a format marker. Returns the text-body element
 * that changed so the caller knows what to commit, or null if the selection is
 * empty or outside an editable body.
 */
export function applyFormatToSelection(root: Element, format: RunFormat): HTMLElement | null {
  const sel = root.ownerDocument?.defaultView?.getSelection();
  if (!sel || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (range.collapsed) return null;

  const body = bodyElementOf(range.commonAncestorContainer, root);
  if (!body) return null;

  const doc = root.ownerDocument!;
  const wrapper = doc.createElement('span');
  wrapper.setAttribute(EDIT_ATTR.fmt, JSON.stringify(format));
  paint(wrapper, format);

  // extractContents may split spans mid-run; each half keeps its run key, so
  // both still resolve to the same source run and inherit the wrapper's format.
  wrapper.appendChild(range.extractContents());
  range.insertNode(wrapper);

  // Keep the same text selected so repeated toolbar clicks compound.
  sel.removeAllRanges();
  const after = doc.createRange();
  after.selectNodeContents(wrapper);
  sel.addRange(after);

  return body;
}

/**
 * The formatting in effect at the caret, for the toolbar's active state. Read
 * from the resolved model run (not its `<a:rPr>`) so inherited properties —
 * from the layout, master and theme — are reflected too.
 */
export function formatAtSelection(root: Element, resolve: Resolver): RunFormat {
  const sel = root.ownerDocument?.defaultView?.getSelection();
  if (!sel || sel.rangeCount === 0) return {};

  let node: Node | null = sel.getRangeAt(0).startContainer;
  let pending: RunFormat = {};
  const overrides: RunFormat[] = [];

  while (node && node !== root) {
    if (node.nodeType === 1) {
      const el = node as Element;
      const raw = el.getAttribute(EDIT_ATTR.fmt);
      if (raw) {
        try {
          overrides.unshift(JSON.parse(raw) as RunFormat);
        } catch {
          /* a malformed marker just contributes nothing */
        }
      }
      const runKey = el.getAttribute(EDIT_ATTR.run);
      if (runKey) {
        const run = resolve(runKey) as TextRun | undefined;
        if (run) {
          pending = {
            ...(run.bold !== undefined ? { bold: run.bold } : {}),
            ...(run.italic !== undefined ? { italic: run.italic } : {}),
            ...(run.underline !== undefined ? { underline: run.underline } : {}),
            ...(run.strike !== undefined ? { strike: run.strike } : {}),
            ...(run.sizePt !== undefined ? { sizePt: run.sizePt } : {}),
            ...(run.color ? { colorHex: run.color.hex } : {}),
            ...(run.font ? { font: run.font } : {}),
          };
        }
        break;
      }
    }
    node = node.parentNode;
  }

  // Nearer markers win over the run's own resolved properties.
  return overrides.reduce<RunFormat>((acc, o) => mergeFormat(acc, o), pending);
}
