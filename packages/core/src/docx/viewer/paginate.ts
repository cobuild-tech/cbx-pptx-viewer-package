/**
 * Measurement-based pagination.
 *
 * Word lays a section's content into fixed-size pages: content flows in the box
 * between the top and bottom margins, and overflow moves to the next page. We
 * reproduce that by rendering blocks into an off-screen measurer at the content
 * width and greedily filling pages of the section's content height, splitting
 * paragraphs by line and tables by row when a block straddles a page boundary.
 *
 * Requires a DOM (browser); it runs at viewer mount, not at parse time.
 */
import { renderBlock, SHEET_FONT, type RenderDeps } from '../render/dom.js';
import type { DocxSection, DocxPage, DocxBlock, DocxParagraph, DocxTable, DocxRun } from '../model.js';

/** Flow every section into fixed-size pages. */
export function paginate(sections: DocxSection[], deps: RenderDeps): DocxPage[] {
  const measure = createMeasurer();
  try {
    const pages: DocxPage[] = [];
    for (const section of sections) {
      for (const page of paginateSection(section, deps, measure)) {
        pages.push({ ...page, index: pages.length });
      }
    }
    return pages;
  } finally {
    measure.host.remove();
  }
}

interface Measurer {
  host: HTMLElement;
  /** Measure a block's height (including vertical margins) at the content width. */
  height(block: DocxBlock, deps: RenderDeps, widthPx: number): number;
  /** Raw element height for an already-rendered clone. */
  el: HTMLDivElement;
}

function createMeasurer(): Measurer {
  const host = document.createElement('div');
  host.style.cssText =
    'position:absolute;left:-99999px;top:0;visibility:hidden;pointer-events:none;';
  const el = document.createElement('div');
  el.style.boxSizing = 'border-box';
  el.style.fontFamily = SHEET_FONT.fontFamily;
  el.style.fontSize = SHEET_FONT.fontSize;
  el.style.lineHeight = SHEET_FONT.lineHeight;
  (el.style as CSSStyleDeclaration & { overflowWrap: string }).overflowWrap = 'break-word';
  host.appendChild(el);
  (document.body ?? document.documentElement).appendChild(host);

  return {
    host,
    el,
    height(block, deps, widthPx) {
      el.style.width = `${widthPx}px`;
      el.replaceChildren(renderBlock(block, deps));
      const child = el.firstElementChild as HTMLElement;
      const cs = getComputedStyle(child);
      return child.offsetHeight + parseFloat(cs.marginTop || '0') + parseFloat(cs.marginBottom || '0');
    },
  };
}

function paginateSection(section: DocxSection, deps: RenderDeps, m: Measurer): DocxPage[] {
  const contentW = section.size.wPx - section.margins.leftPx - section.margins.rightPx;
  const contentH = section.size.hPx - section.margins.topPx - section.margins.bottomPx;

  const pages: DocxPage[] = [];
  let current: DocxBlock[] = [];
  let usedH = 0;

  const makePage = (): DocxPage => ({
    index: pages.length,
    size: section.size,
    margins: section.margins,
    elements: current,
    ...(section.header ? { header: section.header } : {}),
    ...(section.footer ? { footer: section.footer } : {}),
  });
  const flush = () => {
    pages.push(makePage());
    current = [];
    usedH = 0;
  };

  const queue: DocxBlock[] = [...section.blocks];
  while (queue.length) {
    const block = queue.shift()!;
    const avail = contentH - usedH;
    const h = m.height(block, deps, contentW);

    if (h <= avail) {
      current.push(block);
      usedH += h;
      continue;
    }

    const split = trySplit(block, avail, contentW, deps, m);
    if (split) {
      current.push(split.head);
      queue.unshift(split.tail);
      flush();
      continue;
    }

    // Block doesn't fit and can't be split into the remaining space.
    if (current.length > 0) {
      flush();
      queue.unshift(block); // retry on a fresh page
    } else {
      // Fresh page and still doesn't fit: place it whole (it overflows).
      current.push(block);
      flush();
    }
  }

  if (current.length > 0 || pages.length === 0) flush();
  return pages;
}

interface Split {
  head: DocxBlock;
  tail: DocxBlock;
}

function trySplit(
  block: DocxBlock,
  availH: number,
  contentW: number,
  deps: RenderDeps,
  m: Measurer,
): Split | null {
  if (block.kind === 'paragraph') return splitParagraph(block, availH, contentW, deps, m);
  if (block.kind === 'table') return splitTable(block, availH, contentW, deps, m);
  return null;
}

// ─── Paragraph splitting (by line) ───────────────────────────────────────────

function splitParagraph(
  p: DocxParagraph,
  availH: number,
  contentW: number,
  deps: RenderDeps,
  m: Measurer,
): Split | null {
  const total = p.runs.reduce((n, r) => n + r.text.length, 0);
  if (total === 0) return null;

  // Largest character prefix whose rendered height fits availH.
  let lo = 0;
  let hi = total;
  let best = 0;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const h = m.height(slice(p, 0, mid), deps, contentW);
    if (h <= availH) {
      best = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  if (best <= 0) return null; // not even one line fits here

  // Back off to a word boundary so we don't split mid-word.
  const cut = wordBoundary(p, best);
  if (cut <= 0) return null;

  return { head: slice(p, 0, cut, 'head'), tail: slice(p, cut, total, 'tail') };
}

/** Build a paragraph containing only the runs' text in global range [start,end). */
function slice(p: DocxParagraph, start: number, end: number, part?: 'head' | 'tail'): DocxParagraph {
  const runs: DocxRun[] = [];
  let off = 0;
  for (const run of p.runs) {
    const len = run.text.length;
    const s = Math.max(start, off);
    const e = Math.min(end, off + len);
    if (e > s) {
      const text = run.text.slice(s - off, e - off);
      // Keep a break/tab only when this run begins within the slice.
      const atStart = off >= start && off < end;
      runs.push({
        ...run,
        text,
        ...(atStart ? {} : { breakBefore: undefined, tabBefore: undefined }),
      });
    }
    off += len;
  }

  const out: DocxParagraph = { ...p, runs };
  if (part === 'head') {
    out.spaceAfterPt = undefined;
  } else if (part === 'tail') {
    out.listMarker = undefined;
    out.spaceBeforePt = undefined;
    out.pageBreakBefore = undefined;
  }
  return out;
}

/** Nearest whitespace at or before `idx` in the paragraph's concatenated text. */
function wordBoundary(p: DocxParagraph, idx: number): number {
  const text = p.runs.map((r) => r.text).join('');
  for (let i = idx; i > 0; i--) {
    if (/\s/.test(text[i - 1]!)) return i;
  }
  return idx; // a single long word — hard-split it
}

// ─── Table splitting (by row) ────────────────────────────────────────────────

function splitTable(
  table: DocxTable,
  availH: number,
  contentW: number,
  deps: RenderDeps,
  m: Measurer,
): Split | null {
  if (table.rows.length <= 1) return null;

  let best = 0;
  for (let k = 1; k < table.rows.length; k++) {
    const h = m.height({ ...table, rows: table.rows.slice(0, k) }, deps, contentW);
    if (h <= availH) best = k;
    else break;
  }
  if (best <= 0) return null; // not even one row fits here

  return {
    head: { ...table, rows: table.rows.slice(0, best) },
    tail: { ...table, rows: table.rows.slice(best) },
  };
}
