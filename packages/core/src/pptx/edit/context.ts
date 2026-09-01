/**
 * EditContext — the bridge between model objects and the DOM.
 *
 * The renderer needs a string it can put in an attribute; the reconciler needs
 * to turn that string back into the model object. This holds both directions of
 * that mapping for one render pass, and answers the renderer's questions about
 * what may be edited.
 *
 * (This is the `address.ts` role from the design: model identity for the DOM.)
 */
import { localName } from '../../oxml/xml.js';
import type { Deck } from '../deck/deck.js';
import type { Shape, TextBody, TextRun } from '../model.js';
import type { EditRenderContext, ShapeFrame } from '../render/primitives.js';
import { SLIDE_FRAME } from '../render/primitives.js';

export class EditContext implements EditRenderContext {
  private readonly deck: Deck;
  private slideIndex: number;
  private byKey = new Map<string, object>();
  private keys = new WeakMap<object, string>();
  /** Where each selectable shape's coordinate space lands on the slide. */
  private frames = new WeakMap<object, ShapeFrame>();
  private next = 0;
  /** The text body currently open for typing, if the user has entered one. */
  private editing: TextBody | null = null;

  constructor(deck: Deck, slideIndex: number) {
    this.deck = deck;
    this.slideIndex = slideIndex;
  }

  /** Point the context at a different slide and drop the previous keys. */
  retarget(slideIndex: number): void {
    // Re-rendering the same slide (after a commit, a resize, a font landing)
    // must not close a text box the user is typing in; moving to a different
    // slide always does.
    if (slideIndex !== this.slideIndex) this.editing = null;
    this.slideIndex = slideIndex;
    this.reset();
  }

  /**
   * Open a text body for typing, or close whichever one is open. The renderer
   * reads this to decide where contentEditable goes; the caller re-renders.
   */
  setTextEditing(body: TextBody | null): void {
    this.editing = body;
  }

  /** The text body currently open for typing. */
  get editingBody(): TextBody | null {
    return this.editing;
  }

  /** Forget every key. Call before re-rendering a slide. */
  reset(): void {
    this.byKey = new Map();
    this.keys = new WeakMap();
    this.frames = new WeakMap();
    this.next = 0;
  }

  key(model: object): string {
    const existing = this.keys.get(model);
    if (existing) return existing;
    const k = `k${this.next++}`;
    this.keys.set(model, k);
    this.byKey.set(k, model);
    return k;
  }

  /** The model object a key refers to, or undefined if it is stale. */
  resolve(key: string | null | undefined): object | undefined {
    return key ? this.byKey.get(key) : undefined;
  }

  editable(body: TextBody): boolean {
    return this.deck.isEditable(this.slideIndex, body);
  }

  textEditing(body: TextBody): boolean {
    return this.editing === body;
  }

  shapeFrame(shape: Shape, frame: ShapeFrame): void {
    this.frames.set(shape, frame);
  }

  /**
   * The space a shape was last drawn in — the slide's own unless the shape sits
   * inside a group.
   */
  frameOf(shape: Shape): ShapeFrame {
    return this.frames.get(shape) ?? SLIDE_FRAME;
  }

  selectable(shape: Shape): boolean {
    return this.deck.isEditable(this.slideIndex, shape);
  }

  isField(run: TextRun): boolean {
    const src = this.deck.sourceOf(run);
    return !!src && localName(src.node.name) === 'fld';
  }
}
