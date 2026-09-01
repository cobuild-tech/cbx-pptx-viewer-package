/**
 * EditSession — the seam between a UI and the XML.
 *
 * Everything funnels through {@link commitTextBody}: snapshot the part, mutate
 * the cached XML node, mark it dirty, re-parse the slide. The viewer and the
 * React bindings hold a session; neither of them touches XML directly.
 */
import type { Deck } from '../deck/deck.js';
import type { Shape, Slide, TextBody, Transform } from '../model.js';
import { History, type Snapshot } from '../../oxml/edit/history.js';
import type { XmlNode } from '../../oxml/xml.js';
import { writeTextBody, type ParaEdit } from './xmlWrite.js';
import { removeShape, reorderShape, writeTransform, type ZOrderMove } from './shapeOps.js';

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

  /** True if the slide at `index` may be deleted (a deck must keep one slide). */
  canDeleteSlide(index: number): boolean {
    return this.deck.canDeleteSlide(index);
  }

  /**
   * Delete a slide and rebuild the deck. Unlike a text edit this spans several
   * parts — presentation.xml, its relationships, [Content_Types].xml and the
   * slide part itself — so the whole set is snapshotted as one undo entry.
   *
   * Returns false, having changed nothing, if the slide cannot be deleted.
   */
  deleteSlide(index: number): boolean {
    if (!this.deck.canDeleteSlide(index)) return false;
    const before = this.deck.slideDeletionSnapshots(index);
    if (before.length === 0) return false;
    if (!this.deck.deleteSlide(index)) return false;
    this.history.push(before);
    this.onChange?.(Math.min(index, this.deck.slides.length - 1));
    return true;
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

  // ─── Shape edits ───────────────────────────────────────────────────────────
  //
  // Moving, resizing, deleting and restacking all rewrite the slide's own part
  // and nothing else, so they share the text path's undo model: snapshot the
  // part, mutate the cached XML, re-parse the slide.

  /** True if this shape belongs to the slide itself and may be manipulated. */
  isShapeEditable(slideIndex: number, shape: Shape): boolean {
    return this.deck.isEditable(slideIndex, shape);
  }

  /** Write a new position/size/rotation for a shape. */
  commitShapeTransform(slideIndex: number, shape: Shape, transform: Transform): Slide | undefined {
    return this.commitShapeTransforms(slideIndex, [{ shape, transform }]);
  }

  /**
   * Write new geometry for several shapes as **one** undoable change — dragging
   * a multiple selection is a single user action, and undoing it a shape at a
   * time would be wrong.
   */
  commitShapeTransforms(
    slideIndex: number,
    edits: Array<{ shape: Shape; transform: Transform }>,
  ): Slide | undefined {
    return this.editShapes(
      slideIndex,
      edits.map((e) => ({ shape: e.shape, mutate: (node: XmlNode) => writeTransform(node, e.transform) })),
    );
  }

  /** Delete shapes from the slide, as one undoable change. */
  deleteShapes(slideIndex: number, shapes: Shape[]): Slide | undefined {
    return this.editShapes(
      slideIndex,
      shapes.map((shape) => ({ shape, mutate: (node: XmlNode, root: XmlNode) => removeShape(root, node) })),
    );
  }

  /** Move a shape through the z-order. */
  reorderShape(slideIndex: number, shape: Shape, move: ZOrderMove): Slide | undefined {
    return this.editShapes(slideIndex, [
      { shape, mutate: (node: XmlNode, root: XmlNode) => reorderShape(root, node, move) },
    ]);
  }

  /**
   * Run shape mutations against the slide's XML as one undoable change.
   *
   * The snapshot is taken before any mutation but only *kept* if at least one
   * reports that it changed something — otherwise a no-op (dropping a shape
   * back where it started, or sending one that is already at the back further
   * back) would leave an undo entry that undoes nothing.
   */
  private editShapes(
    slideIndex: number,
    edits: Array<{ shape: Shape; mutate: (node: XmlNode, root: XmlNode) => boolean }>,
  ): Slide | undefined {
    const root = this.deck.slideXml(slideIndex);
    if (!root || edits.length === 0) return undefined;

    const targets: Array<{ node: XmlNode; mutate: (node: XmlNode, root: XmlNode) => boolean }> = [];
    let part: string | undefined;
    for (const { shape, mutate } of edits) {
      const src = this.deck.sourceOf(shape);
      if (!src || !this.deck.isEditable(slideIndex, shape)) continue;
      targets.push({ node: src.node, mutate });
      part = src.part;
    }
    if (part === undefined) return undefined;

    const before = this.deck.snapshotPart(part);
    let changed = false;
    for (const t of targets) changed = t.mutate(t.node, root) || changed;
    if (!changed) return undefined;
    if (before !== undefined) this.history.push({ part, xml: before });

    this.deck.markSlideDirty(slideIndex);
    const slide = this.deck.rebuildSlide(slideIndex);
    this.onChange?.(slideIndex);
    return slide;
  }

  undo(): Slide | undefined {
    return this.restore(this.history.undo((p) => this.deck.snapshot(p)));
  }

  redo(): Slide | undefined {
    return this.restore(this.history.redo((p) => this.deck.snapshot(p)));
  }

  private restore(snapshots: Snapshot[] | undefined): Slide | undefined {
    if (!snapshots || snapshots.length === 0) return undefined;
    // A change set that touches presentation.xml moved the running order, so the
    // whole deck has to be re-read; anything else is one slide's own XML.
    const structural = snapshots.some((s) => s.part === this.deck.presentationPart);
    for (const snap of snapshots) this.deck.restore(snap);

    if (structural) {
      this.deck.rebuild();
      // Undoing a deletion puts a slide part back, so land the viewer on it.
      // Redoing one leaves no slide part in the set, and the caller clamps its
      // own position instead.
      const restored = new Set(snapshots.filter((s) => !s.absent).map((s) => s.part));
      const slide = this.deck.slides.find((s) => restored.has(s.part));
      this.onChange?.(slide ? slide.index : 0);
      return slide;
    }

    const index = this.deck.slides.findIndex((s) => s.part === snapshots[0]!.part);
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
