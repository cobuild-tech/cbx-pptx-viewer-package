/**
 * The slide rail — PowerPoint's left-hand thumbnail strip.
 *
 * Thumbnails go through the *same* {@link renderSlide} as the main stage, just
 * CSS-scaled down, so a thumbnail can never disagree with the slide it stands
 * for. The cost of that is a full render per slide, which is why they are drawn
 * lazily: each thumbnail starts as an empty frame of the right size (so the
 * scrollbar is honest from the first paint) and only renders when it scrolls
 * into view. A 200-slide deck therefore costs about ten renders on open.
 *
 * Thumbnails are deliberately rendered *without* the edit context. They are a
 * navigation control, not a second place to type: giving them contentEditable
 * regions would put two DOM nodes in the running for the same model object and
 * let a stray click start an edit the user cannot see.
 */
import type { Slide, SlideSize } from '../model.js';
import { renderSlide } from '../render/dom.js';
import { installStyleSheet } from '../../oxml/stylesheet.js';

export interface FilmstripOptions {
  /** Live read of the deck's slides — re-read on every rebuild. */
  slides: () => Slide[];
  size: SlideSize;
  imageUrl: (part: string) => string | undefined;
  /** Called when the user picks a thumbnail. */
  onSelect: (index: number) => void;
  /** Rail width in CSS px, including the number column. Default 200. */
  width?: number;
}

const STYLE_ID = 'cbx-filmstrip-styles';
const DEFAULT_WIDTH = 200;
/** Width of the slide-number gutter, matching PowerPoint's. */
const NUMBER_COLUMN = 26;
const PADDING = 8;

/**
 * Colours are exposed as custom properties so a host can theme the rail without
 * forking it; the defaults sit next to the neutral grey stage a slide is
 * normally shown on.
 */
function css(width: number): string {
  return [
    `.cbx-strip{--cbx-strip-bg:#1f1f1f;--cbx-strip-fg:#9a9a9a;--cbx-strip-accent:#c43b1c;`,
    `flex:0 0 ${width}px;width:${width}px;overflow-y:auto;overflow-x:hidden;`,
    `background:var(--cbx-strip-bg);box-sizing:border-box;padding:${PADDING}px 0;`,
    `font:12px system-ui,sans-serif;-webkit-user-select:none;user-select:none;}`,

    `.cbx-thumb{display:flex;align-items:flex-start;gap:4px;width:100%;`,
    `padding:4px ${PADDING}px;border:0;background:none;cursor:pointer;`,
    `color:var(--cbx-strip-fg);font:inherit;text-align:left;}`,

    `.cbx-thumb-num{flex:0 0 ${NUMBER_COLUMN - 4}px;padding-top:2px;text-align:right;`,
    `font-variant-numeric:tabular-nums;line-height:1;}`,

    // The frame clips the scaled slide and carries the selection ring. A slide
    // is white-ish by default, so the idle border only has to be a hairline.
    `.cbx-thumb-frame{position:relative;overflow:hidden;background:#fff;`,
    `outline:1px solid rgba(255,255,255,.18);outline-offset:-1px;}`,
    `.cbx-thumb:hover .cbx-thumb-frame{outline:1px solid rgba(255,255,255,.5);}`,
    `.cbx-thumb[aria-selected="true"]{color:#fff;}`,
    `.cbx-thumb[aria-selected="true"] .cbx-thumb-frame{outline:2px solid var(--cbx-strip-accent);outline-offset:0;}`,
    `.cbx-thumb:focus-visible .cbx-thumb-frame{outline:2px solid #0d6efd;outline-offset:0;}`,

    // The rendered slide sits at full size and is scaled from its top-left.
    `.cbx-thumb-slide{position:absolute;top:0;left:0;transform-origin:top left;pointer-events:none;}`,
  ].join('');
}

interface Thumb {
  button: HTMLButtonElement;
  frame: HTMLDivElement;
  /** Null until the thumbnail has been lazily rendered. */
  rendered: HTMLElement | null;
}

export class Filmstrip {
  /** The rail element, for the viewer to place beside its stage. */
  readonly el: HTMLDivElement;
  private readonly doc: Document;
  private readonly options: Required<Pick<FilmstripOptions, 'width'>> & FilmstripOptions;
  private thumbs: Thumb[] = [];
  private active = -1;
  private observer: IntersectionObserver | null = null;
  private readonly disposeStyles: () => void;

  constructor(doc: Document, options: FilmstripOptions) {
    this.doc = doc;
    this.options = { width: DEFAULT_WIDTH, ...options };
    this.disposeStyles = installStyleSheet(doc, STYLE_ID, css(this.options.width));

    this.el = doc.createElement('div');
    this.el.className = 'cbx-strip';
    this.el.setAttribute('role', 'listbox');
    this.el.setAttribute('aria-label', 'Slides');

    this.rebuild();
  }

  get count(): number {
    return this.thumbs.length;
  }

  /** The scale a slide is drawn at inside a thumbnail frame. */
  private get scale(): number {
    const frameWidth = this.options.width - NUMBER_COLUMN - PADDING * 2 - 4;
    return frameWidth / this.options.size.wPx;
  }

  /**
   * Rebuild every thumbnail. For when the deck's slide *list* changed — one
   * deleted, reordered — where thumbnails no longer line up with their slides.
   * Scroll position is preserved so the rail doesn't jump under the user.
   */
  rebuild(): void {
    const scrollTop = this.el.scrollTop;
    this.teardownObserver();
    this.el.textContent = '';
    this.thumbs = [];

    const { wPx, hPx } = this.options.size;
    const scale = this.scale;
    const slides = this.options.slides();

    for (let i = 0; i < slides.length; i++) {
      const button = this.doc.createElement('button');
      button.type = 'button';
      button.className = 'cbx-thumb';
      button.setAttribute('role', 'option');
      button.setAttribute('aria-selected', 'false');
      button.setAttribute('aria-label', `Slide ${i + 1}`);
      const index = i;
      button.addEventListener('click', () => this.options.onSelect(index));

      const num = this.doc.createElement('span');
      num.className = 'cbx-thumb-num';
      num.textContent = String(i + 1);

      const frame = this.doc.createElement('div');
      frame.className = 'cbx-thumb-frame';
      // Sized up front from the slide's aspect ratio, so the rail scrolls
      // correctly before a single thumbnail has been rendered.
      frame.style.width = `${Math.round(wPx * scale)}px`;
      frame.style.height = `${Math.round(hPx * scale)}px`;
      frame.dataset['cbxThumb'] = String(i);

      button.append(num, frame);
      this.el.appendChild(button);
      this.thumbs.push({ button, frame, rendered: null });
    }

    this.el.scrollTop = scrollTop;
    this.setupObserver();
    if (this.active >= 0) this.setActive(Math.min(this.active, this.thumbs.length - 1));
  }

  /**
   * Re-render one thumbnail, for when that slide's *content* changed. A
   * thumbnail that was never rendered stays unrendered — it will pick up the
   * new content when it first scrolls into view.
   */
  refresh(index: number): void {
    const thumb = this.thumbs[index];
    if (!thumb || !thumb.rendered) return;
    thumb.rendered = null;
    this.render(index);
  }

  /** Mark a thumbnail current and scroll it into view. */
  setActive(index: number): void {
    for (let i = 0; i < this.thumbs.length; i++) {
      this.thumbs[i]!.button.setAttribute('aria-selected', i === index ? 'true' : 'false');
    }
    this.active = index;
    const thumb = this.thumbs[index];
    if (!thumb) return;
    // Only scroll when the thumbnail is actually out of view, so clicking a
    // visible thumbnail doesn't shunt the rail around.
    const railTop = this.el.scrollTop;
    const railBottom = railTop + this.el.clientHeight;
    const top = thumb.button.offsetTop;
    const bottom = top + thumb.button.offsetHeight;
    if (top < railTop) this.el.scrollTop = top;
    else if (bottom > railBottom) this.el.scrollTop = bottom - this.el.clientHeight;
  }

  /** Draw a thumbnail's slide, if it hasn't been drawn already. */
  private render(index: number): void {
    const thumb = this.thumbs[index];
    const slide = this.options.slides()[index];
    if (!thumb || thumb.rendered || !slide) return;

    const el = renderSlide(slide, this.options.size, { imageUrl: this.options.imageUrl });
    el.classList.add('cbx-thumb-slide');
    el.style.transform = `scale(${this.scale})`;
    thumb.frame.textContent = '';
    thumb.frame.appendChild(el);
    thumb.rendered = el;
  }

  private setupObserver(): void {
    // jsdom and other non-browser DOMs have no IntersectionObserver; without it
    // there is no scroll to react to either, so draw everything up front.
    if (typeof IntersectionObserver === 'undefined') {
      for (let i = 0; i < this.thumbs.length; i++) this.render(i);
      return;
    }
    this.observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          const index = Number((entry.target as HTMLElement).dataset['cbxThumb']);
          this.render(index);
          this.observer?.unobserve(entry.target);
        }
      },
      // A screen of lead-in either way, so scrolling lands on drawn thumbnails.
      { root: this.el, rootMargin: '200px 0px' },
    );
    for (const thumb of this.thumbs) this.observer.observe(thumb.frame);
  }

  private teardownObserver(): void {
    this.observer?.disconnect();
    this.observer = null;
  }

  destroy(): void {
    this.teardownObserver();
    this.el.remove();
    this.thumbs = [];
    this.disposeStyles();
  }
}
