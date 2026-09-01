/**
 * One injected `<style>` per document, per id, reference-counted.
 *
 * Viewer chrome (editable-region outlines, the slide filmstrip) is delivered as
 * a stylesheet rather than inline styles so `:hover`/`:focus` do the work
 * natively. Several viewers can be mounted on one page, so the sheet is shared
 * and only removed once the last user of it goes away.
 */

interface SheetHandle {
  el: HTMLStyleElement;
  refs: number;
}

const installed = new WeakMap<Document, Map<string, SheetHandle>>();

/**
 * Ensure a stylesheet with `id` exists in `doc` and holds `css`. Returns a
 * disposer that drops one reference, removing the sheet at zero.
 *
 * Calling this again with the same id *replaces* the css — which is how a
 * viewer restyles itself in place (e.g. changing the outline mode).
 */
export function installStyleSheet(doc: Document, id: string, css: string): () => void {
  let sheets = installed.get(doc);
  if (!sheets) {
    sheets = new Map();
    installed.set(doc, sheets);
  }
  let handle = sheets.get(id);
  if (handle && !handle.el.isConnected) {
    // Something outside us detached the sheet — a host swapping the document,
    // a test harness resetting <head>. Re-attach the element we already own
    // rather than silently styling nothing and losing the reference count.
    doc.head.appendChild(handle.el);
  }
  if (!handle) {
    const el = doc.createElement('style');
    el.id = id;
    doc.head.appendChild(el);
    handle = { el, refs: 0 };
    sheets.set(id, handle);
  }
  handle.refs++;
  handle.el.textContent = css;

  let released = false;
  return () => {
    if (released) return;
    released = true;
    const map = installed.get(doc);
    const h = map?.get(id);
    if (!map || !h) return;
    h.refs--;
    if (h.refs <= 0) {
      h.el.remove();
      map.delete(id);
    }
  };
}
