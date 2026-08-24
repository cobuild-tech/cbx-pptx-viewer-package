/**
 * Editor affordances — showing the user where the editable regions are.
 *
 * Format-agnostic: keyed on the data attributes in ./attrs.ts.
 *
 * Delivered as one injected stylesheet rather than per-element inline styles so
 * `:hover` and `:focus-within` do the work natively; the renderer only has to
 * stamp `data-cbx-body`, which it already does. One `<style>` per document,
 * reference-counted so several viewers on a page don't fight over it.
 */

export type TextBoxOutline = 'hover' | 'always' | 'none';

const STYLE_ID = 'cbx-edit-styles';

/**
 * The outline sits on the text *body* box, which is inset from the shape by the
 * text insets — so it frames where text actually flows rather than the shape
 * bounds. Widths are in px and the slide is CSS-scaled, so they thin out as the
 * slide shrinks; that is the same behaviour as any other slide geometry.
 */
function css(mode: TextBoxOutline): string {
  const rules = [
    // A text cursor is the cheapest hint that something is editable.
    `[data-cbx-body]{cursor:text;}`,
    // Give an empty box a clickable target instead of a zero-height sliver.
    `[data-cbx-body]{min-height:1em;}`,
    // Focus always wins, in every mode.
    `[data-cbx-body]:focus,[data-cbx-body]:focus-within{outline:2px solid #0d6efd;outline-offset:2px;}`,
  ];
  if (mode === 'hover') {
    rules.push(
      `[data-cbx-body]:hover:not(:focus):not(:focus-within){outline:1px dashed rgba(13,110,253,.65);outline-offset:2px;}`,
    );
  }
  if (mode === 'always') {
    rules.push(
      `[data-cbx-body]:not(:focus):not(:focus-within){outline:1px dashed rgba(13,110,253,.45);outline-offset:2px;}`,
      `[data-cbx-body]:hover:not(:focus):not(:focus-within){outline-color:rgba(13,110,253,.9);}`,
    );
  }
  return rules.join('\n');
}

interface StyleHandle {
  el: HTMLStyleElement;
  refs: number;
}

const installed = new WeakMap<Document, StyleHandle>();

/**
 * Ensure the editor stylesheet is present in `doc` and set to `mode`. Returns a
 * disposer that removes it once the last viewer using it goes away.
 */
export function installEditStyles(doc: Document, mode: TextBoxOutline): () => void {
  let handle = installed.get(doc);
  if (!handle) {
    const el = doc.createElement('style');
    el.id = STYLE_ID;
    doc.head.appendChild(el);
    handle = { el, refs: 0 };
    installed.set(doc, handle);
  }
  handle.refs++;
  handle.el.textContent = css(mode);

  let released = false;
  return () => {
    if (released) return;
    released = true;
    const h = installed.get(doc);
    if (!h) return;
    h.refs--;
    if (h.refs <= 0) {
      h.el.remove();
      installed.delete(doc);
    }
  };
}
