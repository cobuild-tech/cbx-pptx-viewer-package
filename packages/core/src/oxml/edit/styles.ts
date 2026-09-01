/**
 * Editor affordances — showing the user where the editable regions are.
 *
 * Format-agnostic: keyed on the data attributes in ./attrs.ts.
 *
 * Delivered as one injected stylesheet rather than per-element inline styles so
 * `:hover` and `:focus-within` do the work natively; the renderer only has to
 * stamp `data-cbx-body`, which it already does.
 */
import { installStyleSheet } from '../stylesheet.js';

export type TextBoxOutline = 'hover' | 'always' | 'none';

const STYLE_ID = 'cbx-edit-styles';

/**
 * The outline sits on the text *body* box, which is inset from the shape by the
 * text insets — so it frames where text actually flows rather than the shape
 * bounds. Widths are in px and the slide is CSS-scaled, so they thin out as the
 * slide shrinks; that is the same behaviour as any other slide geometry.
 */
function css(mode: TextBoxOutline): string {
  // Every rule is scoped to a box that is *currently* contentEditable, not
  // merely addressable. A format where selection comes first (PPTX shapes) only
  // turns that on for the box the user has entered, and a text cursor over a
  // box you cannot yet type in would be a lie.
  const box = '[data-cbx-body][contenteditable="true"]';
  const rules = [
    // A text cursor is the cheapest hint that something is editable.
    `${box}{cursor:text;}`,
    // Give an empty box a clickable target instead of a zero-height sliver.
    `${box}{min-height:1em;}`,
    // Focus always wins, in every mode.
    `${box}:focus,${box}:focus-within{outline:2px solid #0d6efd;outline-offset:2px;}`,
  ];
  if (mode === 'hover') {
    rules.push(
      `${box}:hover:not(:focus):not(:focus-within){outline:1px dashed rgba(13,110,253,.65);outline-offset:2px;}`,
    );
  }
  if (mode === 'always') {
    rules.push(
      `${box}:not(:focus):not(:focus-within){outline:1px dashed rgba(13,110,253,.45);outline-offset:2px;}`,
      `${box}:hover:not(:focus):not(:focus-within){outline-color:rgba(13,110,253,.9);}`,
    );
  }
  return rules.join('\n');
}

/**
 * Ensure the editor stylesheet is present in `doc` and set to `mode`. Returns a
 * disposer that removes it once the last viewer using it goes away.
 */
export function installEditStyles(doc: Document, mode: TextBoxOutline): () => void {
  return installStyleSheet(doc, STYLE_ID, css(mode));
}
