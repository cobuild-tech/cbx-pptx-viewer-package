/**
 * Undo/redo as whole-part snapshots.
 *
 * A snapshot is the serialized XML of one part. That is coarse compared with
 * inverse ops, but it is exact by construction — restoring cannot drift from
 * what the edit actually did — and the parts we edit are small. `OpcPackage`
 * already provides both halves (`serializePart` / `setPart`).
 *
 * One user-visible change may span several parts: formatting a spreadsheet cell
 * rewrites `xl/styles.xml` *and* the worksheet. So an entry is a *change set* —
 * a list of snapshots restored together — and single-part callers simply push
 * one snapshot.
 */

export interface Snapshot {
  part: string;
  xml: string;
}

/** Default cap on retained change sets, oldest dropped first. */
const DEFAULT_LIMIT = 100;

export class History {
  private undoStack: Snapshot[][] = [];
  private redoStack: Snapshot[][] = [];
  private readonly limit: number;

  constructor(limit = DEFAULT_LIMIT) {
    this.limit = limit;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /**
   * Record the pre-edit state of one part, or of every part an edit is about to
   * touch. Clears the redo branch.
   */
  push(snapshot: Snapshot | Snapshot[]): void {
    const set = Array.isArray(snapshot) ? snapshot : [snapshot];
    if (set.length === 0) return;
    this.undoStack.push(set);
    if (this.undoStack.length > this.limit) this.undoStack.shift();
    this.redoStack = [];
  }

  /**
   * Pop the change set to restore, given a reader for each part's current state
   * to keep for redo.
   */
  undo(current: (part: string) => string | undefined): Snapshot[] | undefined {
    return this.step(this.undoStack, this.redoStack, current);
  }

  redo(current: (part: string) => string | undefined): Snapshot[] | undefined {
    return this.step(this.redoStack, this.undoStack, current);
  }

  private step(
    from: Snapshot[][],
    to: Snapshot[][],
    current: (part: string) => string | undefined,
  ): Snapshot[] | undefined {
    const set = from.pop();
    if (!set) return undefined;
    const inverse: Snapshot[] = [];
    for (const snap of set) {
      const now = current(snap.part);
      if (now !== undefined) inverse.push({ part: snap.part, xml: now });
    }
    if (inverse.length) to.push(inverse);
    return set;
  }

  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
  }
}
