/**
 * Version storage for edited DOCX documents.
 *
 * A version is a self-contained snapshot of every part that differs from the
 * originally-loaded document (in practice just `word/document.xml`, a few KB),
 * plus the op-log since the previous version (history/audit). Because each
 * version stores the *complete* XML of the edited parts, restore is trivial and
 * needs no replay: revert all edited parts to the original baseline, then apply
 * the snapshot.
 *
 * The store is an interface so persistence is pluggable — an in-memory default
 * ships here; a backend (e.g. writing each VersionPayload as JSON to a folder)
 * implements the same four methods.
 */
import type { EditOp } from './ops.js';

export interface VersionMeta {
  id: string;
  /** The version this one was saved on top of. */
  parentId?: string;
  label?: string;
  /** Epoch millis, supplied by the caller (core never reads the clock). */
  createdAt: number;
  /** Stable hash of the snapshot content, for dedup / change detection. */
  contentHash: string;
}

/** What the document hands the store to persist (the store assigns the id). */
export interface VersionInput {
  parentId?: string;
  label?: string;
  createdAt: number;
  contentHash: string;
  /** partPath → full serialized XML of every edited part. */
  changedParts: Record<string, string>;
  /** Ops applied since the parent version (history). */
  ops: EditOp[];
}

export interface VersionPayload {
  meta: VersionMeta;
  changedParts: Record<string, string>;
  ops: EditOp[];
}

export interface DocxVersionStore {
  list(docId: string): Promise<VersionMeta[]>;
  load(docId: string, versionId: string): Promise<VersionPayload | undefined>;
  /** Persist a new version; the store assigns and returns its id. */
  save(docId: string, input: VersionInput): Promise<VersionMeta>;
  remove?(docId: string, versionId: string): Promise<void>;
}

/** Default in-memory store (dev/tests). Versions live until the page unloads. */
export class InMemoryVersionStore implements DocxVersionStore {
  private readonly store = new Map<string, VersionPayload[]>();
  private seq = 0;

  async list(docId: string): Promise<VersionMeta[]> {
    return (this.store.get(docId) ?? []).map((p) => p.meta);
  }

  async load(docId: string, versionId: string): Promise<VersionPayload | undefined> {
    return (this.store.get(docId) ?? []).find((p) => p.meta.id === versionId);
  }

  async save(docId: string, input: VersionInput): Promise<VersionMeta> {
    const meta: VersionMeta = {
      id: `v${++this.seq}`,
      parentId: input.parentId,
      label: input.label,
      createdAt: input.createdAt,
      contentHash: input.contentHash,
    };
    const arr = this.store.get(docId) ?? [];
    arr.push({ meta, changedParts: input.changedParts, ops: input.ops });
    this.store.set(docId, arr);
    return meta;
  }

  async remove(docId: string, versionId: string): Promise<void> {
    const arr = this.store.get(docId);
    if (arr) this.store.set(docId, arr.filter((p) => p.meta.id !== versionId));
  }
}

/** Deterministic FNV-1a hash of a string (no crypto dependency). */
export function hashString(s: string): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(16).padStart(8, '0');
}
