/**
 * PDF version store — saves snapshots of edit state so the user can
 * restore the document to any previously saved version.
 *
 * A version is a snapshot of the entire edit map (blockId → newText) at a
 * given point in time.  Restore replaces the live edit map with that snapshot.
 */
import type { PdfEditOp } from '../model.js';

export interface PdfVersion {
  /** Unique version identifier. */
  id: string;
  /** Human-readable label shown in the UI. */
  label: string;
  /** Unix timestamp in ms when the version was saved. */
  savedAt: number;
  /** Full snapshot of all edits at save time (blockId → newText). */
  edits: Map<string, string>;
  /** SHA-1-style content hash for change detection. */
  contentHash: string;
}

export interface PdfVersionStore {
  save(version: PdfVersion): void;
  list(): PdfVersion[];
  get(id: string): PdfVersion | undefined;
}

export class InMemoryPdfVersionStore implements PdfVersionStore {
  private readonly versions: PdfVersion[] = [];

  save(version: PdfVersion): void {
    this.versions.push(version);
  }

  list(): PdfVersion[] {
    return [...this.versions];
  }

  get(id: string): PdfVersion | undefined {
    return this.versions.find(v => v.id === id);
  }
}

/** Deterministic hash for the current edit map. */
export function hashEdits(edits: Map<string, string>): string {
  const parts: string[] = [];
  for (const [k, v] of edits) {
    parts.push(`${k}\x00${v}`);
  }
  parts.sort();
  return parts.join('\x01');
}

/** Build a PdfVersion from the current edit map. */
export function makeVersion(
  edits: Map<string, string>,
  label: string,
): PdfVersion {
  return {
    id: `v-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    label,
    savedAt: Date.now(),
    edits: new Map(edits),
    contentHash: hashEdits(edits),
  };
}

/** Restore a saved version into a live edits map. */
export function restoreVersion(
  version: PdfVersion,
  edits: Map<string, string>,
): PdfEditOp[] {
  const ops: PdfEditOp[] = [];

  // Build the set of all block IDs touched by either the version or current edits.
  const allIds = new Set([...edits.keys(), ...version.edits.keys()]);
  for (const id of allIds) {
    const currentText = edits.get(id);
    const savedText   = version.edits.get(id);
    if (currentText !== savedText) {
      ops.push({
        kind: 'replaceText',
        blockId: id,
        pageIndex: 0, // will be resolved by the caller
        oldText: currentText ?? '',
        newText: savedText ?? '',
      });
    }
  }

  // Apply
  edits.clear();
  for (const [k, v] of version.edits) edits.set(k, v);

  return ops;
}
