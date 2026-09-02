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
  /**
   * Marks a paragraph-level format (list, level, alignment, spacing) that
   * reconciliation must write into the paragraph's own properties. Separate
   * from `fmt` because a run format splits runs and a paragraph format must not.
   */
  paraFmt: 'data-cbx-parafmt',
  /** Marks a shape the user may select and move (PPTX). */
  shape: 'data-cbx-shape',
  /**
   * Marks the editor's own chrome — a toolbar and everything in it. Focus
   * moving in here is not the user leaving the text: the command that follows
   * still needs the caret, so the viewer keeps the editing session alive.
   */
  chrome: 'data-cbx-toolbar',
} as const;
