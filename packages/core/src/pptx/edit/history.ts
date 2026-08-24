/**
 * Undo/redo as whole-part snapshots.
 *
 * A snapshot is the serialized XML of one slide part. That is coarse compared
 * with inverse ops, but it is exact by construction — restoring cannot drift
 * from what the edit actually did — and slide parts are small. `OpcPackage`
 * already provides both halves (`serializePart` / `setPart`).
 */

export interface Snapshot {
  part: string;
  xml: string;
}

/** Default cap on retained snapshots, oldest dropped first. */
const DEFAULT_LIMIT = 100;

export class History {
  private undoStack: Snapshot[] = [];
  private redoStack: Snapshot[] = [];
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

  /** Record the pre-edit state of a part. Clears the redo branch. */
  push(snapshot: Snapshot): void {
    this.undoStack.push(snapshot);
    if (this.undoStack.length > this.limit) this.undoStack.shift();
    this.redoStack = [];
  }

  /**
   * Pop the state to restore, given the part's current state to keep for redo.
   */
  undo(current: (part: string) => string | undefined): Snapshot | undefined {
    const prev = this.undoStack.pop();
    if (!prev) return undefined;
    const now = current(prev.part);
    if (now !== undefined) this.redoStack.push({ part: prev.part, xml: now });
    return prev;
  }

  redo(current: (part: string) => string | undefined): Snapshot | undefined {
    const next = this.redoStack.pop();
    if (!next) return undefined;
    const now = current(next.part);
    if (now !== undefined) this.undoStack.push({ part: next.part, xml: now });
    return next;
  }

  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
  }
}
