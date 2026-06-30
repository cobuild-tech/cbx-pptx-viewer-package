/**
 * DocxDocument — the top-level entry point and edit controller for DOCX files.
 *
 * load() unzips the OPC package, parses styles/numbering, and builds the page
 * IR. The document is always editable: the parsed OOXML tree of word/document.xml
 * is the single source of truth, edits mutate it surgically via EditOps, and the
 * IR is re-derived after each change. export() re-zips a valid .docx (untouched
 * parts byte-for-byte, only edited parts re-serialized).
 *
 * Pagination paths:
 *  • pages — pre-computed with the heuristic height estimator (ready immediately).
 *  • repaginate(heightFn) — re-runs pagination with a caller-supplied height fn
 *    (the viewer passes a DOM measurer for pixel-accurate page breaks).
 */
import { OpcPackage } from '../../oxml/package.js';
import { DocxRelType } from '../relTypes.js';
import type { DocxPage, EmbeddedFont } from '../model.js';
import type { XmlNode } from '../../oxml/xml.js';
import { StyleMap } from '../styles/styles.js';
import { NumberingMap } from '../numbering/numbering.js';
import {
  collectDocxContent,
  paginateDocxContent,
  type DocxFlatContent,
  type BlockHeightFn,
} from './body.js';
import {
  applyOp,
  ops,
  splitRunOps,
  clonedParagraphXml,
  emptyParagraphXml,
  clonedBlankXml,
  type EditOp,
  type RunPropPatch,
  type ParaPropPatch,
} from '../edit/ops.js';
import { parentNodeId, decodeNodeId } from '../edit/nodeId.js';
import {
  hashString,
  type DocxVersionStore,
  type VersionMeta,
} from '../edit/versions.js';

export class DocxDocument {
  /** Pages computed with the heuristic estimator; replaced on every edit. */
  pages: DocxPage[];
  readonly embeddedFonts: EmbeddedFont[];
  private readonly pkg: OpcPackage;
  private readonly urlCache = new Map<string, string>();
  /** Flat parsed content; kept so the viewer can re-paginate with DOM measurements. */
  private flatContent: DocxFlatContent;
  // Retained for re-derivation after edits/restore.
  private readonly docPart: string;
  private readonly styles: StyleMap;
  private readonly numbering: NumberingMap;
  // Undo/redo: each entry is a *group* of inverse ops (applied in order) so one
  // user action (e.g. a multi-run paragraph commit, or a run split) is one step.
  private readonly undoStack: EditOp[][] = [];
  private readonly redoStack: EditOp[][] = [];
  private readonly listeners = new Set<() => void>();
  // Versioning (optional; configured via configureVersioning).
  private versionStore?: DocxVersionStore;
  private docId = 'document';
  private currentVersion?: VersionMeta;
  private opsSinceVersion: EditOp[] = [];

  private constructor(
    pkg: OpcPackage,
    docPart: string,
    styles: StyleMap,
    numbering: NumberingMap,
    flatContent: DocxFlatContent,
    pages: DocxPage[],
    embeddedFonts: EmbeddedFont[],
  ) {
    this.pkg = pkg;
    this.docPart = docPart;
    this.styles = styles;
    this.numbering = numbering;
    this.flatContent = flatContent;
    this.pages = pages;
    this.embeddedFonts = embeddedFonts;
  }

  static load(data: ArrayBuffer | Uint8Array): DocxDocument {
    const pkg = OpcPackage.load(data);

    const docPart = pkg.relByType('', DocxRelType.OfficeDocument)?.target;
    if (!docPart) throw new Error('Not a WordprocessingML package: no document part.');
    const docXml = pkg.getXml(docPart);
    if (!docXml) throw new Error('document.xml is missing or empty.');

    const bodyEl = findChild(docXml, 'body');
    if (!bodyEl) throw new Error('document.xml has no <w:body>.');

    const stylesPart = pkg.relByType(docPart, DocxRelType.Styles)?.target;
    const styles = StyleMap.parse(stylesPart ? pkg.getXml(stylesPart) : undefined);

    const numberingPart = pkg.relByType(docPart, DocxRelType.Numbering)?.target;
    const numbering = NumberingMap.parse(numberingPart ? pkg.getXml(numberingPart) : undefined);

    const bodyPath = [docXml.children.indexOf(bodyEl)];
    const flatContent = collectDocxContent(bodyEl, { pkg, docPart, styles, numbering }, bodyPath);
    const pages = paginateDocxContent(flatContent); // heuristic default

    return new DocxDocument(pkg, docPart, styles, numbering, flatContent, pages, []);
  }

  /**
   * Re-run pagination with a custom height function.
   * The viewer passes a DOM measurer for pixel-accurate page breaks.
   */
  repaginate(heightFn: BlockHeightFn): DocxPage[] {
    return paginateDocxContent(this.flatContent, heightFn);
  }

  // ─── Editing API ───────────────────────────────────────────────────────────

  get canUndo(): boolean { return this.undoStack.length > 0; }
  get canRedo(): boolean { return this.redoStack.length > 0; }
  /** True once any edit has been applied. */
  get isEdited(): boolean { return this.pkg.hasEdits; }

  /** Subscribe to post-edit changes (re-render trigger). Returns an unsubscribe fn. */
  onChange(cb: () => void): () => void {
    this.listeners.add(cb);
    return () => this.listeners.delete(cb);
  }

  /** Apply a single edit op as one undo step. */
  applyEdit(opToApply: EditOp): void {
    this.applyEdits([opToApply]);
  }

  /** Apply several ops as one atomic undo step (re-derives once at the end). */
  applyEdits(opsToApply: EditOp[]): void {
    if (opsToApply.length === 0) return;
    const inverses: EditOp[] = [];
    for (const o of opsToApply) {
      inverses.push(applyOp(this.pkg, o));
      this.opsSinceVersion.push(o);
    }
    inverses.reverse(); // undo applies inverses in reverse order
    this.undoStack.push(inverses);
    this.redoStack.length = 0;
    this.rebuild();
  }

  undo(): void {
    const group = this.undoStack.pop();
    if (!group) return;
    const redo: EditOp[] = [];
    for (const o of group) redo.push(applyOp(this.pkg, o));
    redo.reverse();
    this.redoStack.push(redo);
    this.rebuild();
  }

  redo(): void {
    const group = this.redoStack.pop();
    if (!group) return;
    const undo: EditOp[] = [];
    for (const o of group) undo.push(applyOp(this.pkg, o));
    undo.reverse();
    this.undoStack.push(undo);
    this.rebuild();
  }

  // Convenience editors (build the right EditOp and apply it).

  /** Replace the text of a run, addressed by its nodeId. */
  editRunText(runNodeId: string, text: string): void {
    this.applyEdit(ops.replaceText(runNodeId, text));
  }

  /** Toggle/clear run formatting (bold/italic/underline/strike/color/size/font). */
  setRunProps(runNodeId: string, props: RunPropPatch): void {
    this.applyEdit(ops.setRunProps(runNodeId, props));
  }

  /**
   * Apply formatting to a character range within a run, splitting it as needed
   * (one undo step). For a whole-run range this is equivalent to setRunProps.
   */
  formatRunRange(runNodeId: string, start: number, end: number, props: RunPropPatch): void {
    this.applyEdits(splitRunOps(this.pkg, runNodeId, start, end, props));
  }

  /**
   * Apply formatting to several run ranges (e.g. a cross-run selection) as one
   * undo step. Ranges are processed in descending node-path order so splitting
   * one run never shifts the path of an earlier run in the same parent.
   */
  formatRunRanges(
    ranges: { runId: string; start: number; end: number }[],
    props: RunPropPatch,
  ): void {
    if (ranges.length === 0) return;
    const sorted = ranges.slice().sort((a, b) => {
      const pa = decodeNodeId(a.runId).path;
      const pb = decodeNodeId(b.runId).path;
      for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
        const da = pa[i] ?? -1;
        const db = pb[i] ?? -1;
        if (da !== db) return db - da; // descending
      }
      return 0;
    });
    const all: EditOp[] = [];
    for (const r of sorted) all.push(...splitRunOps(this.pkg, r.runId, r.start, r.end, props));
    this.applyEdits(all);
  }

  /** Set paragraph alignment/style. */
  setParagraphProps(paraNodeId: string, props: ParaPropPatch): void {
    this.applyEdit(ops.setParaProps(paraNodeId, props));
  }

  /** Delete a node (paragraph, run/image, table, …) by nodeId. */
  deleteNode(nodeId: string): void {
    this.applyEdit(ops.remove(nodeId));
  }

  /** Reorder a node within its parent. */
  moveNode(nodeId: string, toIndex: number): void {
    this.applyEdit(ops.move(nodeId, toIndex));
  }

  /** Insert a new paragraph after the given one (cloning its style by default). */
  insertParagraphAfter(paraNodeId: string, cloneStyle = true): void {
    const xml = cloneStyle ? clonedParagraphXml(this.pkg, paraNodeId) : emptyParagraphXml();
    this.applyEdit(ops.insertAfter(paraNodeId, xml));
  }

  /** Insert a blank row after the row containing the given cell. */
  insertRowAfter(cellNodeId: string): void {
    const rowId = parentNodeId(cellNodeId);
    if (!rowId) return;
    this.applyEdit(ops.insertAfter(rowId, clonedBlankXml(this.pkg, rowId)));
  }

  /** Delete the row containing the given cell. */
  deleteRow(cellNodeId: string): void {
    const rowId = parentNodeId(cellNodeId);
    if (rowId) this.applyEdit(ops.remove(rowId));
  }

  // ─── Export ───────────────────────────────────────────────────────────────

  /** Serialize the (edited) document back to a valid .docx byte stream. */
  export(): Uint8Array {
    return this.pkg.toBytes();
  }

  /** Serialize the (edited) document to a .docx Blob (browser / Node 18+). */
  exportBlob(): Blob {
    return this.pkg.toBlob();
  }

  /** Replace a part's bytes directly (used by version restore), then re-derive. */
  setPart(partPath: string, data: string | Uint8Array): void {
    this.pkg.setPart(partPath, data);
    this.undoStack.length = 0;
    this.redoStack.length = 0;
    this.rebuild();
  }

  /** Current serialized XML of a part (for version snapshots). */
  serializePart(partPath: string): string | undefined {
    return this.pkg.serializePart(partPath);
  }

  /** The OPC part path of the main document (e.g. 'word/document.xml'). */
  get documentPart(): string {
    return this.docPart;
  }

  // ─── Versioning ──────────────────────────────────────────────────────────────

  /** Attach a pluggable version store and this document's id within it. */
  configureVersioning(store: DocxVersionStore, docId: string): void {
    this.versionStore = store;
    this.docId = docId;
  }

  get versioningEnabled(): boolean {
    return this.versionStore !== undefined;
  }

  /** List saved versions (as ordered by the store). */
  listVersions(): Promise<VersionMeta[]> {
    return this.versionStore ? this.versionStore.list(this.docId) : Promise.resolve([]);
  }

  /**
   * Snapshot the current edited state as a new version.
   * @param createdAt epoch millis, supplied by the caller (core never reads the clock).
   */
  async saveVersion(label?: string, createdAt = 0): Promise<VersionMeta> {
    if (!this.versionStore) throw new Error('saveVersion: no version store configured.');
    const changedParts: Record<string, string> = {};
    for (const part of this.pkg.editedParts()) {
      const xml = this.pkg.serializePart(part);
      if (xml !== undefined) changedParts[part] = xml;
    }
    const contentHash = hashString(
      Object.keys(changedParts).sort().map((k) => `${k} ${changedParts[k]}`).join(''),
    );
    // Nothing changed since the last saved version → return it, no duplicate.
    if (this.currentVersion && this.currentVersion.contentHash === contentHash) {
      return this.currentVersion;
    }
    const meta = await this.versionStore.save(this.docId, {
      parentId: this.currentVersion?.id,
      label,
      createdAt,
      contentHash,
      changedParts,
      ops: this.opsSinceVersion.slice(),
    });
    this.currentVersion = meta;
    this.opsSinceVersion = [];
    return meta;
  }

  /** True if there are edits not captured by the latest saved version. */
  get hasUnsavedChanges(): boolean {
    if (!this.pkg.hasEdits) return false;
    if (!this.currentVersion) return true;
    const parts: Record<string, string> = {};
    for (const part of this.pkg.editedParts()) {
      const xml = this.pkg.serializePart(part);
      if (xml !== undefined) parts[part] = xml;
    }
    const hash = hashString(Object.keys(parts).sort().map((k) => `${k} ${parts[k]}`).join(''));
    return hash !== this.currentVersion.contentHash;
  }

  /** Restore a previously-saved version, replacing the current document state. */
  async restore(versionId: string): Promise<void> {
    if (!this.versionStore) throw new Error('restore: no version store configured.');
    const payload = await this.versionStore.load(this.docId, versionId);
    if (!payload) throw new Error(`restore: version not found: ${versionId}`);
    // Revert every currently-edited part to the original baseline, then apply the
    // snapshot — so the result is exactly the saved state with nothing stale left.
    for (const part of this.pkg.editedParts()) this.pkg.resetPart(part);
    for (const [part, xml] of Object.entries(payload.changedParts)) this.pkg.setPart(part, xml);
    this.undoStack.length = 0;
    this.redoStack.length = 0;
    this.opsSinceVersion = [];
    this.currentVersion = payload.meta;
    this.rebuild();
  }

  // ─── Internal re-derivation ──────────────────────────────────────────────────

  /** Re-parse the (mutated/replaced) document part and rebuild pages, then notify. */
  private rebuild(): void {
    const docXml = this.pkg.getXml(this.docPart);
    if (!docXml) return;
    const bodyEl = findChild(docXml, 'body');
    if (!bodyEl) return;
    const bodyPath = [docXml.children.indexOf(bodyEl)];
    this.flatContent = collectDocxContent(
      bodyEl,
      { pkg: this.pkg, docPart: this.docPart, styles: this.styles, numbering: this.numbering },
      bodyPath,
    );
    this.pages = paginateDocxContent(this.flatContent);
    for (const listener of this.listeners) listener();
  }

  // ─── Media / lifecycle (unchanged) ───────────────────────────────────────────

  /** Object URL for an embedded media part (cached). Browser only. */
  imageUrl(part: string): string | undefined {
    const cached = this.urlCache.get(part);
    if (cached) return cached;
    const bytes = this.pkg.getBytes(part);
    if (typeof URL === 'undefined' || typeof Blob === 'undefined') return undefined;
    if (!bytes) return undefined;
    const type = this.pkg.contentType(part) ?? 'application/octet-stream';
    const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
    this.urlCache.set(part, url);
    return url;
  }

  /** Raw bytes of an embedded font part, for FontFace registration. */
  fontBytes(part: string): Uint8Array | undefined {
    return this.pkg.getBytes(part);
  }

  /** Release all object URLs created for media. */
  dispose(): void {
    if (typeof URL !== 'undefined') {
      for (const url of this.urlCache.values()) URL.revokeObjectURL(url);
    }
    this.urlCache.clear();
  }
}

function findChild(
  node: XmlNode,
  localName: string,
): XmlNode | undefined {
  return node.children.find((c) => {
    const n = c.name;
    return n === localName || n.endsWith(`:${localName}`);
  });
}
