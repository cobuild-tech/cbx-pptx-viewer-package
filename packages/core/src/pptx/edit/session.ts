/**
 * EditSession — the seam between a UI and the XML.
 *
 * Everything funnels through {@link commitTextBody}: snapshot the part, mutate
 * the cached XML node, mark it dirty, re-parse the slide. The viewer and the
 * React bindings hold a session; neither of them touches XML directly.
 */
import type { Deck } from '../deck/deck.js';
import type { Slide, TextBody } from '../model.js';
import { History } from '../../oxml/edit/history.js';
import { writeTextBody, type ParaEdit } from './xmlWrite.js';

export interface EditSessionOptions {
  /** Called after any change to the deck (commit, undo or redo). */
  onChange?: (slideIndex: number) => void;
  /** Cap on retained undo snapshots. */
  historyLimit?: number;
}

export class EditSession {
  private readonly deck: Deck;
  private readonly history: History;
  private readonly onChange: EditSessionOptions['onChange'];

  constructor(deck: Deck, options: EditSessionOptions = {}) {
    this.deck = deck;
    this.history = new History(options.historyLimit);
    if (options.onChange) this.onChange = options.onChange;
  }

  get canUndo(): boolean {
    return this.history.canUndo;
  }

  get canRedo(): boolean {
    return this.history.canRedo;
  }

  get hasEdits(): boolean {
    return this.deck.hasEdits;
  }

  /**
   * True if this text body may be edited — it must belong to the slide's own
   * part, not to an inherited layout or master.
   */
  isEditable(slideIndex: number, body: TextBody): boolean {
    return this.deck.isEditable(slideIndex, body);
  }

  /**
   * Replace a text body's paragraphs and re-parse the slide.
   *
   * Returns the rebuilt slide, or undefined if the body is not editable or has
   * no recorded source (in which case nothing is mutated).
   */
  commitTextBody(slideIndex: number, body: TextBody, paras: ParaEdit[]): Slide | undefined {
    const src = this.deck.sourceOf(body);
    if (!src || !this.deck.isEditable(slideIndex, body)) return undefined;

    const before = this.deck.snapshotPart(src.part);
    if (before !== undefined) this.history.push({ part: src.part, xml: before });

    writeTextBody(src.node, paras, (model) => this.deck.sourceOf(model));
    this.deck.markSlideDirty(slideIndex);

    const slide = this.deck.rebuildSlide(slideIndex);
    this.onChange?.(slideIndex);
    return slide;
  }

  undo(): Slide | undefined {
    return this.restore(this.history.undo((p) => this.deck.snapshotPart(p)));
  }

  redo(): Slide | undefined {
    return this.restore(this.history.redo((p) => this.deck.snapshotPart(p)));
  }

  private restore(snapshot: { part: string; xml: string } | undefined): Slide | undefined {
    if (!snapshot) return undefined;
    this.deck.restorePart(snapshot.part, snapshot.xml);
    const index = this.deck.slides.findIndex((s) => s.part === snapshot.part);
    if (index === -1) return undefined;
    // setPart drops the cached parse, so this re-reads the restored XML.
    const slide = this.deck.rebuildSlide(index);
    this.onChange?.(index);
    return slide;
  }

  /** Re-zip the deck with all edits applied. */
  exportBlob(): Blob {
    return this.deck.exportBlob();
  }

  /** Re-zip the deck with all edits applied, as raw bytes. */
  exportBytes(): Uint8Array {
    return this.deck.toBytes();
  }
}
