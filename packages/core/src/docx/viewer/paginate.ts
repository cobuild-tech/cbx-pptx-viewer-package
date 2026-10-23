/**
 * Measurement-based pagination.
 *
 * Word lays a section's content into fixed-size pages: content flows in the box
 * between the top and bottom margins, and overflow moves to the next page. We
 * reproduce that by rendering blocks into an off-screen measurer at the content
 * width and greedily filling pages of the section's content height, splitting
 * paragraphs by line and tables by row when a block straddles a page boundary.
 *
 * Word's flow rules on top of that, each mirrored here:
 *  - a page break (<w:br w:type="page"/>, <w:pageBreakBefore>) ends the page;
 *  - <w:keepNext> keeps a paragraph on the page of the one after it — chains of
 *    them move as a unit, unless the chain is longer than a page;
 *  - <w:keepLines> never splits a paragraph; <w:widowControl> never leaves a
 *    single line of one at the bottom or top of a page;
 *  - adjacent paragraph spacing collapses to the larger of the two (as CSS
 *    sibling margins do on the page), so heights are summed the same way;
 *  - space-before is dropped at the top of a page reached by natural overflow;
 *  - a table's header rows repeat on every page it continues onto;
 *  - a continuous section flows on in the same page;
 *  - a footnote goes on the page its reference lands on, and the text above
 *    stops short of it (and of the separator rule over the first one).
 *
 * Requires a DOM (browser); it runs at viewer mount, not at parse time.
 */
import { renderBlock, layoutTabs, stackBlocks, applyWordTextDefaults, anchorFloat, SHEET_FONT, isPageAnchored, type RenderDeps } from '../render/dom.js';
import { formatPageNumber } from '../numbering/numbering.js';
import type {
  DocxSection,
  DocxPage,
  DocxBlock,
  DocxParagraph,
  DocxTable,
  DocxRun,
  DocxAnchor,
  DocxPageSize,
  DocxPageMargins,
} from '../model.js';

/** Flow every section into fixed-size pages. */
export function paginate(sections: DocxSection[], deps: RenderDeps): DocxPage[] {
  const measure = createMeasurer();
  try {
    const pages: DocxPage[] = [];
    const sectionOf: DocxSection[] = [];
    const flow = new PageFlow(deps, measure, pages, sectionOf);
    sections.forEach((section, i) => flow.section(section, i === 0));
    flow.finish();
    numberPages(pages, sectionOf);
    return pages;
  } finally {
    measure.host.remove();
  }
}

/** Display page numbers, honouring each section's restart and number format. */
function numberPages(pages: DocxPage[], sectionOf: DocxSection[]): void {
  let n = 0;
  const perSection = new Map<DocxSection, number>();
  for (const s of sectionOf) perSection.set(s, (perSection.get(s) ?? 0) + 1);
  pages.forEach((page, i) => {
    const section = sectionOf[i]!;
    const startsSection = i === 0 || sectionOf[i - 1] !== section;
    n = startsSection && section.pageNumberStart !== undefined ? section.pageNumberStart : n + 1;
    page.index = i;
    page.pageNumber = formatPageNumber(n, section.pageNumberFormat);
    page.pageCount = pages.length;
    page.sectionPageCount = perSection.get(section) ?? 1;
  });
}

/** A float from an earlier block still overhanging this point of the page. */
interface FloatStub {
  side: 'left' | 'right';
  /** Width it takes from the line, margins included. */
  width: number;
  height: number;
}

interface Measurer {
  host: HTMLElement;
  /**
   * Measure a block's height (including vertical margins) at the content
   * width — beside `stubs`, floats overhanging from earlier blocks, which
   * narrow its lines as they do on the page.
   */
  height(block: DocxBlock, deps: RenderDeps, widthPx: number, stubs?: FloatStub[]): number;
  /** Top/bottom margins of the block as last measured (zero if never measured). */
  margins(block: DocxBlock): { top: number; bottom: number };
  el: HTMLDivElement;
}

function createMeasurer(): Measurer {
  const host = document.createElement('div');
  host.style.cssText = 'position:absolute;left:-99999px;top:0;visibility:hidden;pointer-events:none;';
  const el = document.createElement('div');
  el.style.boxSizing = 'border-box';
  el.style.fontFamily = SHEET_FONT.fontFamily;
  el.style.fontSize = SHEET_FONT.fontSize;
  el.style.lineHeight = SHEET_FONT.lineHeight;
  applyWordTextDefaults(el);
  // Same block layout as the page sheet, so a measured height is the height
  // the block takes on the page.
  stackBlocks(el);
  host.appendChild(el);
  (document.body ?? document.documentElement).appendChild(host);

  const margins = new WeakMap<DocxBlock, { top: number; bottom: number }>();
  return {
    host,
    el,
    margins: (block) => margins.get(block) ?? { top: 0, bottom: 0 },
    height(block, deps, widthPx, stubs) {
      el.style.width = `${widthPx}px`;
      const stubEls = (stubs ?? []).map((f) => {
        const d = document.createElement('div');
        d.style.cssText = `float:${f.side};width:${f.width}px;height:${f.height}px;`;
        return d;
      });
      const rendered = renderBlock(block, deps);
      el.replaceChildren(...stubEls, rendered);
      layoutTabs(el);
      const child = rendered;
      const cs = getComputedStyle(child);
      // Fractional height: offsetHeight rounds to whole pixels, and over the
      // dozens of blocks on a page the lost fractions add up to a line Word
      // would have pushed to the next page.
      const h = child.getBoundingClientRect().height;
      const m = { top: parseFloat(cs.marginTop || '0'), bottom: parseFloat(cs.marginBottom || '0') };
      margins.set(block, m);
      return h + m.top + m.bottom;
    },
  };
}

/** Header/footer reserve and geometry of one kind of page in a section. */
interface PageFrame {
  top: number;
  bottom: number;
  header?: DocxBlock[];
  footer?: DocxBlock[];
}

/** One column of a band: where it sits and what has been placed in it. */
interface Column {
  xPx: number;
  widthPx: number;
  blocks: DocxBlock[];
  /** Height used, from the band's top. */
  used: number;
}

/**
 * A horizontal band of the page holding one section's run of text, in that
 * section's columns. A page is a stack of bands: a continuous section break
 * closes one (balancing its columns) and opens the next below it.
 */
interface Band {
  section: DocxSection;
  top: number;
  cols: Column[];
  separator: boolean;
  /** Blocks fed into this band, before any split here — what balancing re-flows. */
  inputs: DocxBlock[];
}

/**
 * The page currently being filled, carried across sections so a continuous
 * section can flow on in it. Text fills the current band's columns one after
 * another, then the next page.
 */
class PageFlow {
  private bands: Band[] = [];
  private band!: Band;
  private col = 0;
  private usedH = 0;
  /** Bottom margin of the last block placed, which the next block's top margin collapses with. */
  private lastMarginBottom = 0;
  /** The last block placed was a keepNext paragraph — the current block continues its chain. */
  private prevKeepNext = false;
  /** Footnote content referenced from the current page, in reference order. */
  private pageNotes: DocxBlock[] = [];
  /** Notes already placed — a split paragraph's tail must not place them again. */
  private readonly placedNotes = new WeakSet<DocxBlock>();
  /** Wrapped drawings in this column, with how far down they reach (band y). */
  private floats: (FloatStub & { bottom: number })[] = [];
  /** Tails of blocks split in the current band (not fresh input to balance). */
  private bandTails = new WeakSet<DocxBlock>();
  /** Render dependencies with the current column as the text column. */
  private colDeps: RenderDeps;
  /** Section and frame the current page was started with. */
  private pageSection: DocxSection | undefined;
  /** Section whose blocks are flowing — differs from pageSection on a page a continuous section joined. */
  private flowSection: DocxSection | undefined;
  private frame: PageFrame | undefined;
  /** Pages already emitted for the section being flowed (titlePg uses page 0). */
  private sectionPages = 0;
  /** Displayed number of the page being filled — its parity picks odd/even headers. */
  private pageNumber = 0;
  private readonly lastSeenStyleText: Record<string, string> = {};
  private readonly heights = new WeakMap<DocxBlock, Map<number, number>>();
  /** Width of the column being filled. */
  private contentW = 0;

  constructor(
    private readonly deps: RenderDeps,
    private readonly m: Measurer,
    private readonly pages: DocxPage[],
    private readonly sectionOf: DocxSection[],
  ) {
    this.colDeps = deps;
  }

  /** Blocks of the column being filled. */
  private get current(): DocxBlock[] {
    return this.band.cols[this.col]!.blocks;
  }

  private pageEmpty(): boolean {
    return this.bands.every((b) => b.cols.every((c) => c.blocks.length === 0));
  }

  /** Flow one section's blocks. */
  section(section: DocxSection, first: boolean): void {
    const prev = this.pageSection;
    const continuous =
      !first &&
      section.start === 'continuous' &&
      prev !== undefined &&
      prev.size.wPx === section.size.wPx &&
      prev.size.hPx === section.size.hPx;
    this.flowSection = section;
    if (continuous) {
      // The section joins the current page below the previous section's text,
      // whose columns are balanced first; its own first page has passed.
      this.sectionPages = 1;
      this.closeBand(true);
      this.openBand(section, this.bandBottom());
    } else {
      if (!first && !this.pageEmpty()) this.flush(true);
      this.sectionPages = 0;
      // An odd/even-page section leaves a blank page to land on the right side.
      const nextNumber = this.pages.length + 1;
      if (!first && ((section.start === 'oddPage' && nextNumber % 2 === 0) || (section.start === 'evenPage' && nextNumber % 2 === 1))) {
        this.startPage(section);
        this.flush(true);
      }
      this.startPage(section);
    }

    const queue = expandPageBreaks(section.blocks);
    let topOfNaturalPage = false;
    while (queue.length) {
      let block = queue.shift()!;
      if (!this.bandTails.has(block)) this.band.inputs.push(block);

      if (block.kind === 'paragraph' && block.pageBreakBefore && (this.current.length > 0 || !this.pageEmpty())) {
        if (block.columnBreakBefore) this.nextColumn();
        else this.flush();
        topOfNaturalPage = false;
      }
      // Word drops space-before at the top of a page the text overflowed onto.
      if (topOfNaturalPage && block.kind === 'paragraph' && block.spaceBeforePt) {
        block = { ...block, spaceBeforePt: undefined };
      }
      topOfNaturalPage = false;

      const h = this.height(block);
      // The part of this block's top margin hidden in the previous block's
      // bottom margin (the larger of the two is the gap) gives extra room.
      const overlap = this.current.length ? Math.min(this.lastMarginBottom, this.m.margins(block).top) : 0;
      // Footnotes this block references must fit on the same page.
      const notes = footnotesIn(block).filter((b) => !this.placedNotes.has(b));
      const notesH = this.notesHeight(notes);
      const avail = this.contentH() - this.usedH + overlap - notesH;

      // keepNext: if this block and the start of what it is chained to will
      // not fit here but would on a fresh page, start that page now. Word
      // decides this once, where the chain begins; a chain already under way
      // (it was too long for one page) just flows.
      if (block.kind === 'paragraph' && block.keepNext && !this.prevKeepNext && this.current.length > 0) {
        const chain = this.chainHeight(block, queue);
        if (chain > avail && chain <= this.contentH()) {
          this.nextColumn();
          queue.unshift(block);
          topOfNaturalPage = true;
          continue;
        }
      }

      // Space-after at the foot of a page takes no room: it only has to fit
      // when something follows on the same page, and then it collapses anyway.
      if (h - this.m.margins(block).bottom <= avail) {
        this.place(block, h);
        this.addNotes(notes, notesH);
        continue;
      }

      const split = this.trySplit(block, avail);
      if (split) {
        this.place(split.head, this.m.height(split.head, this.colDeps, this.contentW));
        this.addNotes(notes, notesH);
        this.prevKeepNext = block.kind === 'paragraph' && !!block.keepNext;
        this.bandTails.add(split.tail);
        queue.unshift(split.tail);
        this.nextColumn();
        topOfNaturalPage = true;
        continue;
      }

      // Doesn't fit and can't be split into the remaining space.
      if (this.current.length > 0) {
        this.bandTails.add(block); // already counted as input
        this.nextColumn();
        queue.unshift(block); // retry in the next column / page
        topOfNaturalPage = true;
      } else {
        // An empty column and still doesn't fit: place it whole (it overflows).
        this.place(block, h);
        this.addNotes(notes, notesH);
        this.nextColumn();
        topOfNaturalPage = true;
      }
    }
  }

  finish(): void {
    if (!this.pageEmpty() || this.pages.length === 0) this.flush(true);
  }

  private place(block: DocxBlock, h: number): void {
    const m = this.m.margins(block);
    const overlap = this.current.length ? Math.min(this.lastMarginBottom, m.top) : 0;
    this.addFloats(block, this.usedH - overlap + m.top);
    this.current.push(block);
    this.usedH += h - overlap;
    this.lastMarginBottom = m.bottom;
    this.prevKeepNext = block.kind === 'paragraph' && !!block.keepNext;
  }

  /** Height the given notes add to the page, the separator rule included if they are its first. */
  private notesHeight(notes: DocxBlock[]): number {
    if (!notes.length) return 0;
    let h = 0;
    if (!this.pageNotes.length) {
      for (const b of this.pageSection?.footnoteSeparator ?? []) h += this.height(b);
    }
    for (const b of notes) h += this.height(b);
    return h;
  }

  private addNotes(notes: DocxBlock[], h: number): void {
    if (!notes.length) return;
    for (const b of notes) this.placedNotes.add(b);
    this.pageNotes.push(...notes);
    this.usedH += h;
  }

  /** Height a column of the current band may fill. */
  private contentH(): number {
    return this.pageSection!.size.hPx - this.frame!.top - this.frame!.bottom - this.band.top;
  }

  private height(block: DocxBlock): number {
    // Beside an overhanging float the block's lines are shorter: measure it
    // there, and don't cache a height that depends on where it sits.
    const stubs = this.floats
      .filter((f) => f.bottom > this.usedH)
      .map((f) => ({ side: f.side, width: f.width, height: f.bottom - this.usedH }));
    if (stubs.length) return this.m.height(block, this.colDeps, this.contentW, stubs);
    // Cached per width: the same block measures differently in another column.
    let byWidth = this.heights.get(block);
    if (!byWidth) this.heights.set(block, (byWidth = new Map()));
    const cached = byWidth.get(this.contentW);
    if (cached !== undefined) return cached;
    const h = this.m.height(block, this.colDeps, this.contentW);
    byWidth.set(this.contentW, h);
    return h;
  }

  /** Record the wrapped drawings a just-placed paragraph floats, from its top at `top`. */
  private addFloats(block: DocxBlock, top: number): void {
    if (block.kind !== 'paragraph' || !block.anchors) return;
    for (const a of block.anchors) {
      const f = anchorFloat(a, this.colDeps.column, block.indentLeftPx ?? 0, block.indentRightPx ?? 0);
      if (!f) continue;
      const width = f.fullWidth ? this.contentW : f.inset + f.w + f.gap;
      this.floats.push({ side: f.side, width, height: f.h, bottom: top + f.top + f.h + f.bottom });
    }
  }

  /** Height of a keepNext chain: every chained block whole, plus the first line of the one it ends on. */
  private chainHeight(first: DocxParagraph, queue: DocxBlock[]): number {
    let total = this.height(first);
    for (const next of queue) {
      if (next.kind === 'paragraph' && next.keepNext) {
        total += this.height(next);
        continue;
      }
      total += this.firstLineHeight(next);
      break;
    }
    return total;
  }

  private firstLineHeight(block: DocxBlock): number {
    if (block.kind === 'paragraph') {
      const total = textLength(block);
      return total > 0 ? this.m.height(slice(block, 0, 1, 'head'), this.colDeps, this.contentW) : this.height(block);
    }
    if (block.kind === 'table' && block.rows.length > 1) {
      return this.m.height({ ...block, rows: block.rows.slice(0, 1) }, this.colDeps, this.contentW);
    }
    return this.height(block);
  }

  // ─── Bands and columns ────────────────────────────────────────────────────

  /** Open a band for `section` at `top` (content-box y), in its first column. */
  private openBand(section: DocxSection, top: number): void {
    const cols = section.columns?.cols ?? [
      { xPx: 0, widthPx: section.size.wPx - section.margins.leftPx - section.margins.rightPx },
    ];
    this.band = {
      section,
      top,
      cols: cols.map((c) => ({ xPx: c.xPx, widthPx: c.widthPx, blocks: [], used: 0 })),
      separator: section.columns?.separator ?? false,
      inputs: [],
    };
    this.bands.push(this.band);
    this.bandTails = new WeakSet();
    this.enterColumn(0);
  }

  private enterColumn(i: number): void {
    this.col = i;
    const c = this.band.cols[i]!;
    const section = this.band.section;
    this.contentW = c.widthPx;
    this.colDeps = { ...this.deps, column: { leftPx: section.margins.leftPx + c.xPx, widthPx: c.widthPx } };
    this.usedH = c.used;
    this.lastMarginBottom = 0;
    this.floats = [];
  }

  /** The current column is full: move to the next one, or the next page after the last. */
  private nextColumn(): void {
    this.band.cols[this.col]!.used = this.usedH;
    if (this.col < this.band.cols.length - 1) this.enterColumn(this.col + 1);
    else this.flush();
  }

  /** Bottom of the current band (after {@link closeBand}): its tallest column. */
  private bandBottom(): number {
    return this.band.top + Math.max(0, ...this.band.cols.map((c) => c.used));
  }

  /**
   * Close the current band. At a continuous section break Word balances a
   * multi-column band: its text is shared out so the columns end level.
   */
  private closeBand(balance: boolean): void {
    this.band.cols[this.col]!.used = this.usedH;
    if (!balance || this.band.cols.length < 2 || !this.band.inputs.length) return;
    const total = Math.max(...this.band.cols.map((c) => c.used));
    const fill = (h: number) => fillColumns(this.band.inputs, this.band.cols, h, this.deps, this.band.section, this.m);
    let lo = 0;
    let hi = total;
    let best = fill(hi);
    if (!best) return;
    for (let i = 0; i < 14 && hi - lo > 0.5; i++) {
      const mid = (lo + hi) / 2;
      const r = fill(mid);
      if (r) {
        best = r;
        hi = mid;
      } else lo = mid;
    }
    this.band.cols.forEach((c, i) => {
      c.blocks = best![i]!.blocks;
      c.used = best![i]!.used;
    });
  }

  /** Begin a page in `section`, using its first-page frame if it is the section's first. */
  private startPage(section: DocxSection): void {
    this.pageSection = section;
    // The displayed number restarts where the section says so.
    this.pageNumber = this.sectionPages === 0 && section.pageNumberStart !== undefined ? section.pageNumberStart : this.pageNumber + 1;
    const first = this.sectionPages === 0 && section.titlePg;
    const even = !first && section.evenAndOddHeaders && this.pageNumber % 2 === 0;
    const header = first ? section.firstHeader : even ? section.evenHeader : section.header;
    const footer = first ? section.firstFooter : even ? section.evenFooter : section.footer;
    const fullW = section.size.wPx - section.margins.leftPx - section.margins.rightPx;
    this.frame = this.frameFor(section, header, footer, fullW);
    this.bands = [];
    this.openBand(section, 0);
  }

  /**
   * A tall header/footer (banner) can extend past its margin band into the
   * body. Word reserves space for it: the body starts below max(topMargin,
   * header distance + header height) and ends above the mirror for the footer.
   */
  private frameFor(section: DocxSection, header: DocxBlock[] | undefined, footer: DocxBlock[] | undefined, fullW: number): PageFrame {
    const { margins, size } = section;
    const deps = { ...this.deps, column: { leftPx: margins.leftPx, widthPx: fullW } };
    const top = Math.max(
      margins.topPx,
      header ? margins.headerPx + bandHeight(header, deps, fullW, this.m) : 0,
      pageEdgeReserve(header, size, margins, 'top'),
    );
    const bottom = Math.max(
      margins.bottomPx,
      footer ? margins.footerPx + bandHeight(footer, deps, fullW, this.m) : 0,
      pageEdgeReserve(footer, size, margins, 'bottom'),
    );
    return { top, bottom, ...(header ? { header } : {}), ...(footer ? { footer } : {}) };
  }

  /** Emit the current page and start the next one in the same section. */
  private flush(final = false): void {
    const section = this.pageSection!;
    const frame = this.frame!;
    this.band.cols[this.col]!.used = this.usedH;
    const elements = this.bands.flatMap((b) => b.cols.flatMap((c) => c.blocks));

    const pageResolved = { ...this.lastSeenStyleText };
    const firstOccurrences: Record<string, string> = {};
    const checkBlock = (b: DocxBlock) => {
      if (b.kind === 'paragraph') {
        const style = b.styleName?.toLowerCase();
        const text = b.runs.map((r) => (r.hidden ? '' : r.text)).join('').trim();
        if (style && text) {
          if (firstOccurrences[style] === undefined) firstOccurrences[style] = text;
          this.lastSeenStyleText[style] = text;
        }
      } else if (b.kind === 'table') {
        for (const row of b.rows) for (const cell of row) if (cell) cell.content.forEach(checkBlock);
      }
    };
    elements.forEach(checkBlock);
    Object.assign(pageResolved, firstOccurrences);

    const notes = this.pageNotes.length ? [...(section.footnoteSeparator ?? []), ...this.pageNotes] : undefined;
    this.pageNotes = [];
    const multi = this.bands.some((b) => b.cols.length > 1);
    const page: DocxPage = {
      ...(notes ? { footnotes: notes } : {}),
      index: this.pages.length,
      size: section.size,
      margins: section.margins,
      elements,
      contentTopPx: frame.top,
      contentBottomPx: frame.bottom,
      resolvedStyles: pageResolved,
      ...(frame.header ? { header: frame.header } : {}),
      ...(frame.footer ? { footer: frame.footer } : {}),
    };
    if (multi) {
      page.bands = this.bands
        .filter((b) => b.cols.some((c) => c.blocks.length))
        .map((b) => ({
          columns: b.cols.map((c) => ({ xPx: c.xPx, widthPx: c.widthPx })),
          blocks: b.cols.map((c) => c.blocks),
          heightPx: Math.max(0, ...b.cols.map((c) => c.used)),
          ...(b.separator ? { separator: true } : {}),
        }));
    }
    this.pages.push(page);
    this.sectionOf.push(section);
    this.sectionPages++;
    if (!final) this.startPage(this.flowSection ?? section);
  }

  private trySplit(block: DocxBlock, availH: number): Split | null {
    if (block.kind === 'paragraph') {
      if (block.keepLines) return null;
      return splitParagraph(block, availH, this.contentW, this.colDeps, this.m);
    }
    // On an otherwise empty page a table that keeps together still has to
    // break somewhere.
    if (block.kind === 'table') return splitTable(block, availH, this.contentW, this.colDeps, this.m, this.pageEmpty());
    return null;
  }
}

/**
 * Flow `inputs` into columns of height `h` — whole blocks while they fit,
 * splitting the one that straddles a column's foot. Returns each column's
 * blocks and height, or undefined if the text does not fit. Used to balance
 * columns at a continuous section break.
 */
function fillColumns(
  inputs: DocxBlock[],
  cols: Column[],
  h: number,
  deps: RenderDeps,
  section: DocxSection,
  m: Measurer,
): { blocks: DocxBlock[]; used: number }[] | undefined {
  const out = cols.map(() => ({ blocks: [] as DocxBlock[], used: 0 }));
  const queue = [...inputs];
  let i = 0;
  let lastBottom = 0;
  while (queue.length) {
    if (i >= cols.length) return undefined;
    const col = cols[i]!;
    const colDeps = { ...deps, column: { leftPx: section.margins.leftPx + col.xPx, widthPx: col.widthPx } };
    const block = queue.shift()!;
    const bh = m.height(block, colDeps, col.widthPx);
    const mg = m.margins(block);
    const o = out[i]!;
    const overlap = o.blocks.length ? Math.min(lastBottom, mg.top) : 0;
    const avail = h - o.used + overlap;
    if (bh - mg.bottom <= avail) {
      o.blocks.push(block);
      o.used += bh - overlap;
      lastBottom = mg.bottom;
      continue;
    }
    const split =
      block.kind === 'paragraph' && !block.keepLines
        ? splitParagraph(block, avail, col.widthPx, colDeps, m)
        : block.kind === 'table'
          ? splitTable(block, avail, col.widthPx, colDeps, m, o.blocks.length === 0)
          : null;
    if (split) {
      o.blocks.push(split.head);
      o.used += m.height(split.head, colDeps, col.widthPx) - overlap;
      queue.unshift(split.tail);
    } else if (!o.blocks.length) {
      // Too tall for any column of this height.
      return undefined;
    } else {
      queue.unshift(block);
    }
    i++;
    lastBottom = 0;
  }
  return out;
}

/** The content of every footnote a block references, in reference order. */
function footnotesIn(block: DocxBlock): DocxBlock[] {
  const out: DocxBlock[] = [];
  const visit = (b: DocxBlock) => {
    if (b.kind === 'paragraph') {
      for (const r of b.runs) if (r.note?.kind === 'footnote') out.push(...r.note.blocks);
    } else if (b.kind === 'table') {
      for (const row of b.rows) for (const cell of row) if (cell) cell.content.forEach(visit);
    }
  };
  visit(block);
  return out;
}

/**
 * Split paragraphs at their page breaks, so the paginator sees each page's part
 * as its own block: the part after a break starts a new page.
 */
function expandPageBreaks(blocks: DocxBlock[]): DocxBlock[] {
  const out: DocxBlock[] = [];
  for (const b of blocks) {
    if (b.kind !== 'paragraph' || !b.runs.some((r) => r.pageBreakBefore)) {
      out.push(b);
      continue;
    }
    let start = 0;
    let first = true;
    const pushPart = (end: number) => {
      const runs = b.runs.slice(start, end);
      const part: DocxParagraph = { ...b, runs };
      if (!first) {
        // The break itself is consumed; the part begins the next page (or column).
        if (runs[0]!.columnBreak) part.columnBreakBefore = true;
        runs[0] = { ...runs[0]!, breakBefore: undefined, pageBreakBefore: undefined, columnBreak: undefined };
        part.pageBreakBefore = true;
        part.listMarker = undefined;
        part.spaceBeforePt = undefined;
        part.anchors = undefined;
      }
      if (end < b.runs.length) part.spaceAfterPt = undefined;
      out.push(part);
      first = false;
    };
    b.runs.forEach((r, i) => {
      if (r.pageBreakBefore && i > 0) {
        pushPart(i);
        start = i;
      } else if (r.pageBreakBefore && i === 0) {
        // A paragraph opening with a page break: its text starts the next page.
        out.push({ ...b, runs: [], listMarker: undefined, spaceAfterPt: undefined, anchors: undefined });
        first = false;
      }
    });
    pushPart(b.runs.length);
  }
  return out;
}

/**
 * Visual height of a header/footer band: the taller of its flowed blocks and
 * any floating anchors it carries (a banner image often overhangs the table it
 * sits beside). Used to grow the body's top/bottom reserve so text clears it.
 */
function bandHeight(blocks: DocxBlock[], deps: RenderDeps, contentW: number, m: Measurer): number {
  let flowed = 0;
  let anchored = 0;
  for (const b of blocks) {
    flowed += m.height(b, deps, contentW);
    anchored = Math.max(anchored, blockAnchorBottom(b));
  }
  return Math.max(flowed, anchored);
}

/**
 * Largest bottom edge (vOffset + height) of FLOW-framed anchors in a block —
 * these sit inside the header/footer band. Page-framed anchors are reserved
 * separately in page coordinates.
 */
function blockAnchorBottom(block: DocxBlock): number {
  let max = 0;
  if (block.kind === 'paragraph') {
    for (const a of block.anchors ?? []) {
      if (!isPageAnchored(a)) max = Math.max(max, (a.vOffsetPx ?? 0) + a.hPx);
    }
  } else if (block.kind === 'table') {
    for (const row of block.rows) {
      for (const cell of row) {
        if (!cell) continue;
        for (const sub of cell.content) max = Math.max(max, blockAnchorBottom(sub));
      }
    }
  }
  return max;
}

/**
 * Body reserve at the top/bottom page edge for page-framed banners: a banner
 * near the top edge pushes content down past its bottom; one near the bottom
 * pushes it up past its top.
 */
function pageEdgeReserve(
  blocks: DocxBlock[] | undefined,
  size: DocxPageSize,
  margins: DocxPageMargins,
  edge: 'top' | 'bottom',
): number {
  if (!blocks) return 0;
  let r = 0;
  for (const a of collectAnchors(blocks)) {
    if (!isPageAnchored(a)) continue;
    // Behind-text art (a watermark, a full-page background) does not push text.
    if (a.behindDoc || a.wrap === 'none') continue;
    const top = anchorPageTop(a, size, margins);
    const inTopHalf = top < size.hPx / 2;
    if (edge === 'top' && inTopHalf) r = Math.max(r, top + a.hPx);
    else if (edge === 'bottom' && !inTopHalf) r = Math.max(r, size.hPx - top);
  }
  return r;
}

/** All anchors under a block tree (paragraph anchors + table cells). */
function collectAnchors(blocks: DocxBlock[], out: DocxAnchor[] = []): DocxAnchor[] {
  for (const b of blocks) {
    if (b.kind === 'paragraph') out.push(...(b.anchors ?? []));
    else if (b.kind === 'table') {
      for (const row of b.rows) for (const cell of row) if (cell) collectAnchors(cell.content, out);
    }
  }
  return out;
}

/** Page-space top of a page-framed anchor (mirrors renderSheetAnchor). */
function anchorPageTop(a: DocxAnchor, size: DocxPageSize, margins: DocxPageMargins): number {
  if (a.vAlign === 'center') return (size.hPx - a.hPx) / 2;
  if (a.vAlign === 'bottom') return a.relV === 'page' || a.relV === 'bottomMargin' ? size.hPx - a.hPx : size.hPx - margins.bottomPx - a.hPx;
  if (a.vAlign === 'top') return a.relV === 'page' || a.relV === 'topMargin' ? 0 : margins.topPx;
  const off = a.vOffsetPx ?? 0;
  if (a.relV === 'page' || a.relV === 'topMargin') return off;
  if (a.relV === 'bottomMargin' || a.relV === 'outsideMargin') return size.hPx - margins.bottomPx + off;
  return margins.topPx + off;
}

interface Split {
  head: DocxBlock;
  tail: DocxBlock;
}

// ─── Paragraph splitting (by line) ───────────────────────────────────────────

function textLength(p: DocxParagraph): number {
  return p.runs.reduce((n, r) => n + r.text.length, 0);
}

function splitParagraph(
  p: DocxParagraph,
  availH: number,
  contentW: number,
  deps: RenderDeps,
  m: Measurer,
): Split | null {
  const total = textLength(p);
  if (total === 0) return null;

  const cut = fittingCut(p, availH, contentW, deps, m);
  if (cut <= 0) return null;
  if (!p.widowControl) return { head: slice(p, 0, cut, 'head'), tail: slice(p, cut, total, 'tail') };

  // Widow/orphan control: at least two lines on each side, or move it all.
  // Count lines from the text's own height: the paragraph's spacing must come
  // off both the part and the one-line yardstick, or a short head miscounts.
  const textH = (part: DocxParagraph) => {
    const h = m.height(part, deps, contentW);
    const mg = m.margins(part);
    return h - mg.top - mg.bottom;
  };
  const lineH = textH(slice(p, 0, 1, 'body'));
  const lines = (part: DocxParagraph) => Math.round(textH(part) / lineH);
  const head = slice(p, 0, cut, 'body');
  const tail = slice(p, cut, total, 'body');
  const headLines = lines(head);
  if (headLines < 2) return null;
  if (lines(tail) >= 2) return { head: slice(p, 0, cut, 'head'), tail: slice(p, cut, total, 'tail') };
  if (headLines < 3) return null; // giving the tail a line would orphan the head
  // Hand the tail one more line: fit the head into one line less of space.
  const cut2 = fittingCut(p, m.height(slice(p, 0, cut, 'head'), deps, contentW) - lineH + 0.5, contentW, deps, m);
  if (cut2 <= 0) return null;
  return { head: slice(p, 0, cut2, 'head'), tail: slice(p, cut2, total, 'tail') };
}

/** Largest word-boundary character offset whose head part fits in `availH`. */
function fittingCut(p: DocxParagraph, availH: number, contentW: number, deps: RenderDeps, m: Measurer): number {
  const total = textLength(p);
  let lo = 0;
  let hi = total;
  let best = 0;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const h = m.height(slice(p, 0, mid, 'head'), deps, contentW);
    if (h <= availH) {
      best = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  if (best <= 0) return 0; // not even one line fits here
  if (best >= total) return total;
  // Back off to a word boundary so we don't split mid-word.
  return wordBoundary(p, best);
}

/**
 * Build a paragraph containing only the runs' text in global range [start,end).
 * A run without text (an inline picture, a lone break) belongs to the part its
 * position falls in. `head` drops the space-after, `tail` the marker and
 * space-before; `body` keeps both (for line counting).
 */
function slice(p: DocxParagraph, start: number, end: number, part: 'head' | 'tail' | 'body'): DocxParagraph {
  const total = textLength(p);
  const runs: DocxRun[] = [];
  let off = 0;
  for (const run of p.runs) {
    const len = run.text.length;
    const s = Math.max(start, off);
    const e = Math.min(end, off + len);
    const inRange = len === 0 ? off >= start && (off < end || (end === total && off === total)) : e > s;
    if (inRange) {
      const text = len === 0 ? '' : run.text.slice(s - off, e - off);
      // Keep a break/tab only when this run begins within the slice.
      const atStart = off >= start && off < end;
      runs.push({
        ...run,
        text,
        ...(atStart || len === 0 ? {} : { breakBefore: undefined, tabBefore: undefined, ptab: undefined }),
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
    // Floating drawings belong to the page their anchor paragraph starts on.
    out.anchors = undefined;
    // The tail continues the paragraph: no first-line indent of its own.
    if (out.indentFirstLinePx !== undefined && out.indentFirstLinePx > 0) out.indentFirstLinePx = undefined;
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
  freshPage = false,
): Split | null {
  const header = table.headerRows ?? 0;
  if (table.rows.length <= Math.max(1, header) && table.rows.length !== 1) return null;

  let best = 0;
  for (let k = 1; k < table.rows.length; k++) {
    const h = m.height({ ...table, rows: table.rows.slice(0, k), rowHeights: table.rowHeights?.slice(0, k) }, deps, contentW);
    if (h <= availH) best = k;
    else break;
  }

  // A row whose paragraphs are <w:keepNext> stays on the page of the row after
  // it — set on every row, that is how a table is kept in one piece.
  if (!freshPage) {
    while (best > 0 && rowKeepsWithNext(table, best - 1)) best--;
    if (best === 0 && table.rows.length > 0 && rowKeepsWithNext(table, 0)) return null;
  }

  // The first row that does not fit may itself break across the page, as Word
  // does unless the row is <w:cantSplit> (or fixed-height).
  const rowSplit = splitRowAcross(table, best, availH, contentW, deps, m);
  if (rowSplit) return rowSplit;

  // Header rows alone on a page are not a start: Word moves the table on.
  if (best <= header) return null;
  return cutTable(table, best);
}

/** A row is kept with the next when its paragraphs say <w:keepNext>. */
function rowKeepsWithNext(table: DocxTable, r: number): boolean {
  const row = table.rows[r];
  if (!row || r === table.rows.length - 1) return false;
  return row.some((cell) => cell?.content.some((b) => b.kind === 'paragraph' && b.keepNext));
}

/** Split a table between rows: rows [0, at) stay, the rest continue (with repeated header rows). */
function cutTable(table: DocxTable, at: number): Split {
  const header = table.headerRows ?? 0;
  const headRows = table.rows.slice(0, at).map((r) => r.slice());
  const tailRows = table.rows.slice(at).map((r) => r.slice());
  repairVerticalMerges(table, at, headRows, tailRows);

  const head: DocxTable = { ...table, rows: headRows };
  const tail: DocxTable = { ...table, rows: tailRows };
  if (table.rowHeights) {
    head.rowHeights = table.rowHeights.slice(0, at);
    tail.rowHeights = table.rowHeights.slice(at);
  }
  if (table.cantSplit) {
    head.cantSplit = table.cantSplit.slice(0, at);
    tail.cantSplit = table.cantSplit.slice(at);
  }
  delete head.headerRows;
  if (header > 0) {
    // The header rows repeat at the top of the continuation.
    tail.rows = [...table.rows.slice(0, header).map((r) => r.slice()), ...tailRows];
    tail.rowHeights = [...(table.rowHeights?.slice(0, header) ?? Array(header).fill(undefined)), ...(tail.rowHeights ?? [])];
    tail.headerRows = header;
  }
  return { head, tail };
}

/**
 * Break row `k` across the page: each cell keeps the content that fits in the
 * room left and the rest continues in a copy of the row on the next page.
 * Done by giving the table that row twice (head part, tail part) and cutting
 * between the two, so vertical merges through the row carry over like any
 * other cut.
 */
function splitRowAcross(
  table: DocxTable,
  k: number,
  availH: number,
  contentW: number,
  deps: RenderDeps,
  m: Measurer,
): Split | null {
  const row = table.rows[k];
  if (!row || k < (table.headerRows ?? 0)) return null;
  if (table.cantSplit?.[k] || table.rowHeights?.[k]?.rule === 'exact') return null;
  // A cell merged down from this row would need splitting across rows too.
  if (row.some((c) => c && c.rowSpan > 1)) return null;

  // Origins above whose merge runs through row k now span one row more.
  const rows = table.rows.map((r, ri) =>
    r.map((c) => (c && ri < k && ri + c.rowSpan > k ? { ...c, rowSpan: c.rowSpan + 1 } : c)),
  );
  const headRow = row.map((c) => c && { ...c, content: [] as DocxBlock[] });
  const tailRow = row.map((c) => c && { ...c, content: [] as DocxBlock[] });
  const expanded: DocxTable = {
    ...table,
    rows: [...rows.slice(0, k), headRow, tailRow, ...rows.slice(k + 1)],
  };
  if (table.rowHeights) {
    expanded.rowHeights = [...table.rowHeights.slice(0, k), undefined, undefined, ...table.rowHeights.slice(k + 1)];
  }
  if (table.cantSplit) {
    expanded.cantSplit = [...table.cantSplit.slice(0, k), false, false, ...table.cantSplit.slice(k + 1)];
  }

  const above = k > 0 ? m.height({ ...table, rows: table.rows.slice(0, k), rowHeights: table.rowHeights?.slice(0, k) }, deps, contentW) : 0;
  let room = availH - above;
  for (let attempt = 0; attempt < 4 && room > 0; attempt++) {
    let anyHead = false;
    let anyTail = false;
    let col = 0;
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (!cell) continue;
      col = c;
      const pad = cell.cellPaddingPx ?? { top: 0, right: 0, bottom: 0, left: 0 };
      const width = table.colWidths.slice(col, col + cell.colSpan).reduce((a, b) => a + b, 0) - pad.left - pad.right;
      const parts = fillBlocks(cell.content, room - pad.top - pad.bottom, width, deps, m);
      headRow[c]!.content = parts.head;
      tailRow[c]!.content = parts.tail;
      if (parts.head.length) anyHead = true;
      if (parts.tail.length) anyTail = true;
    }
    if (!anyHead || !anyTail) return null;
    const h = m.height({ ...expanded, rows: expanded.rows.slice(0, k + 1), rowHeights: expanded.rowHeights?.slice(0, k + 1) }, deps, contentW);
    if (h <= availH) return cutTable(expanded, k + 1);
    room -= h - availH + 1;
  }
  return null;
}

/**
 * Fill a cell's blocks into `availH`: whole blocks while they fit, then the
 * first that doesn't is split by line if it is a paragraph.
 */
function fillBlocks(
  blocks: DocxBlock[],
  availH: number,
  widthPx: number,
  deps: RenderDeps,
  m: Measurer,
): { head: DocxBlock[]; tail: DocxBlock[] } {
  const head: DocxBlock[] = [];
  let used = 0;
  let lastBottom = 0;
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i]!;
    const h = m.height(b, deps, widthPx);
    const margins = m.margins(b);
    const overlap = head.length ? Math.min(lastBottom, margins.top) : 0;
    if (used + h - overlap - margins.bottom <= availH) {
      head.push(b);
      used += h - overlap;
      lastBottom = margins.bottom;
      continue;
    }
    if (b.kind === 'paragraph') {
      const split = splitParagraph(b, availH - used + overlap, widthPx, deps, m);
      if (split) return { head: [...head, split.head], tail: [split.tail, ...blocks.slice(i + 1)] };
    }
    return { head, tail: blocks.slice(i) };
  }
  return { head, tail: [] };
}

/**
 * A vertically-merged (rowSpan) cell whose origin lands in the head but whose
 * span reaches into the tail leaves the tail's covered positions as `null` —
 * the renderer skips those, so the following cells shift into the wrong column.
 * Clamp the head origin's rowSpan to the head, and re-materialise a (blank)
 * continuation origin at the top of the tail so column alignment is preserved.
 */
function repairVerticalMerges(
  table: DocxTable,
  best: number,
  headRows: DocxTable['rows'][number][number][][],
  tailRows: DocxTable['rows'][number][number][][],
): void {
  const colCount = Math.max(0, ...table.rows.map((r) => r.length));

  // owner[r][c] → the origin cell (and its start row/col) covering (r, c).
  const owner: ({ cell: NonNullable<DocxTable['rows'][number][number]>; r: number; c: number } | null)[][] =
    table.rows.map(() => Array(colCount).fill(null));
  for (let r = 0; r < table.rows.length; r++) {
    for (let c = 0; c < table.rows[r]!.length; c++) {
      const cell = table.rows[r]![c];
      if (!cell) continue;
      for (let dr = 0; dr < cell.rowSpan; dr++) {
        for (let dc = 0; dc < cell.colSpan; dc++) {
          const or = owner[r + dr];
          if (or) or[c + dc] = { cell, r, c };
        }
      }
    }
  }

  const boundary = owner[best];
  if (!boundary) return;
  for (let c = 0; c < colCount; c++) {
    const o = boundary[c];
    if (!o || o.r >= best || o.c !== c) continue; // only the origin's start column, crossing the split
    const headSpan = best - o.r;
    const tailSpan = o.cell.rowSpan - headSpan;
    if (tailSpan <= 0) continue;
    // Clamp the head origin so it doesn't overhang past the page.
    headRows[o.r]![c] = { ...o.cell, rowSpan: headSpan };
    // Blank continuation origin at the top of the tail (keeps the column).
    tailRows[0]![c] = { ...o.cell, content: [], rowSpan: tailSpan };
  }
}
