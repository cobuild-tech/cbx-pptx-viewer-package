/**
 * DOM attribute names the edit layer writes and reads back.
 *
 * Shared by every format's renderer and reconciler, so they live in oxml/ —
 * pptx/ and docx/ must never import from each other.
 */
export const EDIT_ATTR = {
  body: 'data-cbx-body',
  para: 'data-cbx-para',
  run: 'data-cbx-run',
  /** Marks a toolbar-applied format that reconciliation must turn into a run. */
  fmt: 'data-cbx-fmt',
} as const;
