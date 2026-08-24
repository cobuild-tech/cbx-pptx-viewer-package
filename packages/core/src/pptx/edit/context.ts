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
import type { TextBody, TextRun } from '../model.js';
import type { EditRenderContext } from '../render/primitives.js';

export class EditContext implements EditRenderContext {
  private readonly deck: Deck;
  private slideIndex: number;
  private byKey = new Map<string, object>();
  private keys = new WeakMap<object, string>();
  private next = 0;

  constructor(deck: Deck, slideIndex: number) {
    this.deck = deck;
    this.slideIndex = slideIndex;
  }

  /** Point the context at a different slide and drop the previous keys. */
  retarget(slideIndex: number): void {
    this.slideIndex = slideIndex;
    this.reset();
  }

  /** Forget every key. Call before re-rendering a slide. */
  reset(): void {
    this.byKey = new Map();
    this.keys = new WeakMap();
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

  isField(run: TextRun): boolean {
    const src = this.deck.sourceOf(run);
    return !!src && localName(src.node.name) === 'fld';
  }
}
