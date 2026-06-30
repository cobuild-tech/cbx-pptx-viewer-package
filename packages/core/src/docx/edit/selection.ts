/**
 * Browser-only bridge from a DOM Selection to per-run character segments, so
 * the WYSIWYG toolbar can apply formatting to exactly what the user selected.
 *
 * Each run span carries data-docx-id (stamped by the renderer). For every run
 * the selection touches we compute the [start, end) character offsets within
 * that run; the editor then splits the run at those offsets (see
 * DocxDocument.formatRunRanges / splitRunOps).
 */

export interface RunSegment {
  runId: string;
  /** Char offset within the run where the selection starts. */
  start: number;
  /** Char offset within the run where the selection ends (exclusive). */
  end: number;
  /** Total length of the run's text. */
  length: number;
}

/** Reconstruct a run span's text the way the model stores it (<br> → \n, tab → \t). */
export function domRunText(span: HTMLElement): string {
  let out = '';
  span.childNodes.forEach((n) => {
    if (n.nodeType === 3 /* TEXT_NODE */) {
      out += n.nodeValue ?? '';
    } else if (n instanceof HTMLElement) {
      if (n.tagName === 'BR') out += '\n';
      else if (n.classList.contains('docx-tab')) out += '\t';
      else out += n.textContent ?? '';
    }
  });
  return out;
}

/** Map the current selection to per-run segments within `root` (empty if collapsed). */
export function selectionToRunSegments(root: HTMLElement): RunSegment[] {
  if (typeof window === 'undefined' || typeof document === 'undefined') return [];
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return [];
  const range = sel.getRangeAt(0);
  if (!root.contains(range.commonAncestorContainer)) return [];

  const segments: RunSegment[] = [];
  root.querySelectorAll<HTMLElement>('.docx-run[data-docx-id]').forEach((el) => {
    const runId = el.dataset.docxId;
    if (!runId || !range.intersectsNode(el)) return;
    const length = domRunText(el).length;
    const start = el.contains(range.startContainer)
      ? charsBefore(el, range.startContainer, range.startOffset)
      : 0;
    const end = el.contains(range.endContainer)
      ? charsBefore(el, range.endContainer, range.endOffset)
      : length;
    if (end > start) segments.push({ runId, start, end, length });
  });
  return segments;
}

/** Number of characters inside `el` before the (container, offset) point. */
function charsBefore(el: HTMLElement, container: Node, offset: number): number {
  const r = document.createRange();
  r.selectNodeContents(el);
  try {
    r.setEnd(container, offset);
  } catch {
    return 0;
  }
  return r.toString().length;
}
