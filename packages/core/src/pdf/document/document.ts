/**
 * PdfDocument — top-level entry point for PDF files.
 *
 * Uses pdf.js for rendering and text extraction.
 * Uses pdf-lib (via pdf/edit/export.ts) for serializing edits back to bytes.
 *
 * Usage (view):
 *   const doc = await PdfDocument.load(arrayBuffer);
 *   await doc.renderPage(0, 2.0);   // → HTMLCanvasElement
 *   doc.dispose();
 *
 * Usage (edit):
 *   const doc = await PdfDocument.load(arrayBuffer);
 *   await doc.loadAllTextBlocks();  // extract text layer (required before editing)
 *   doc.applyEdit({ kind:'replaceText', blockId:'block-0-1', ... });
 *   doc.applyEdit({ kind:'addAnnotation', annotation: { ... } });
 *   const bytes = await doc.exportBytes();
 *   doc.dispose();
 */
import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy } from 'pdfjs-dist';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import type { PdfPage, PdfTextBlock, PdfEditOp, PdfAnnotation, PdfBlockStyle } from '../model.js';
import { extractPageTextItems, groupItemsIntoBlocks } from '../edit/textLayer.js';
import { applyOp, applyOps, EditHistory } from '../edit/ops.js';
import {
  InMemoryPdfVersionStore,
  makeVersion,
  restoreVersion,
  hashEdits,
  type PdfVersionStore,
  type PdfVersion,
} from '../edit/versions.js';
import { exportPdfWithEdits } from '../edit/export.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

export class PdfDocument {
  /** Page metadata at scale 1.0 — available synchronously after load. */
  readonly pages: PdfPage[];

  /** Text blocks per page — populated after loadAllTextBlocks(). */
  readonly textBlocks: PdfTextBlock[][] = [];

  /** Live text edits map: blockId → current text. */
  readonly editsMap = new Map<string, string>();
  private get edits() { return this.editsMap; }

  /** Live annotations map: annotationId → PdfAnnotation. */
  readonly annotationsMap = new Map<string, PdfAnnotation>();

  /** Live block style overrides map: blockId → PdfBlockStyle. */
  readonly blockStylesMap = new Map<string, PdfBlockStyle>();

  private readonly history = new EditHistory();
  private versionStore: PdfVersionStore = new InMemoryPdfVersionStore();
  private lastSavedHash = '';

  private readonly proxy: PDFDocumentProxy;
  private readonly rawBytes: Uint8Array;

  /** Called whenever any edit is applied, undone, or redone. */
  onChange?: (edits: Map<string, string>) => void;

  private constructor(
    proxy: PDFDocumentProxy,
    pages: PdfPage[],
    rawBytes: Uint8Array,
  ) {
    this.proxy = proxy;
    this.pages = pages;
    this.rawBytes = rawBytes;
  }

  // ── Loading ────────────────────────────────────────────────────────────────

  static async load(data: ArrayBuffer | Uint8Array): Promise<PdfDocument> {
    // Keep an independent copy of the bytes for export.
    // pdf.js getDocument() may TRANSFER the underlying ArrayBuffer to its Web
    // Worker, which detaches it — making any existing Uint8Array view over that
    // buffer zero-length.  We must clone the bytes into a fresh ArrayBuffer
    // before handing anything to pdf.js.
    const sourceBytes = data instanceof Uint8Array ? data : new Uint8Array(data);
    const rawBytes = new Uint8Array(
      sourceBytes.buffer.slice(
        sourceBytes.byteOffset,
        sourceBytes.byteOffset + sourceBytes.byteLength,
      ),
    );

    // Give pdf.js its own copy so a transfer cannot affect rawBytes.
    const pdfJsBytes = new Uint8Array(rawBytes.byteLength);
    pdfJsBytes.set(rawBytes);

    const proxy = await pdfjsLib.getDocument({ data: pdfJsBytes }).promise;

    const pages: PdfPage[] = [];
    for (let i = 0; i < proxy.numPages; i++) {
      const pageProxy = await proxy.getPage(i + 1);
      const viewport = pageProxy.getViewport({ scale: 1.0 });
      pages.push({ index: i, widthPx: viewport.width, heightPx: viewport.height });
      pageProxy.cleanup();
    }

    return new PdfDocument(proxy, pages, rawBytes);
  }

  // ── Rendering ──────────────────────────────────────────────────────────────

  async renderPage(index: number, scale: number): Promise<HTMLCanvasElement> {
    const pageProxy = await this.proxy.getPage(index + 1);
    const viewport = pageProxy.getViewport({ scale });

    const canvas = document.createElement('canvas');
    canvas.width  = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    canvas.style.display = 'block';

    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('Failed to acquire 2D canvas context.');

    await pageProxy.render({ canvasContext: ctx, viewport }).promise;
    pageProxy.cleanup();
    return canvas;
  }

  // ── Text layer ─────────────────────────────────────────────────────────────

  /**
   * Extract and cache text blocks for all pages.
   * Must be called before editing or exporting with edits.
   */
  async loadAllTextBlocks(): Promise<void> {
    if (this.textBlocks.length > 0) return; // already loaded
    for (let i = 0; i < this.pages.length; i++) {
      const page = this.pages[i]!;
      const items = await extractPageTextItems(this.proxy, i, page.heightPx);
      this.textBlocks.push(groupItemsIntoBlocks(items));
    }
  }

  /** Return (possibly cached) text blocks for a single page. */
  async getPageTextBlocks(pageIndex: number): Promise<PdfTextBlock[]> {
    if (this.textBlocks[pageIndex]) return this.textBlocks[pageIndex]!;
    const page = this.pages[pageIndex]!;
    const items = await extractPageTextItems(this.proxy, pageIndex, page.heightPx);
    const blocks = groupItemsIntoBlocks(items);
    this.textBlocks[pageIndex] = blocks;
    return blocks;
  }

  // ── Edit API ───────────────────────────────────────────────────────────────

  /** Apply a single edit operation and push the inverse onto the undo stack. */
  applyEdit(op: PdfEditOp): void {
    const inverse = applyOp(op, this.edits, this.annotationsMap, this.blockStylesMap);
    this.history.push([inverse]);
    this.onChange?.(new Map(this.edits));
  }

  /** Apply multiple ops as a single undoable group. */
  applyEdits(ops: PdfEditOp[]): void {
    const inverses = applyOps(ops, this.edits, this.annotationsMap, this.blockStylesMap);
    this.history.push(inverses);
    this.onChange?.(new Map(this.edits));
  }

  undo(): boolean {
    const inverses = this.history.undo(this.edits, this.annotationsMap, this.blockStylesMap);
    if (!inverses) return false;
    this.onChange?.(new Map(this.edits));
    return true;
  }

  redo(): boolean {
    const inverses = this.history.redo(this.edits, this.annotationsMap, this.blockStylesMap);
    if (!inverses) return false;
    this.onChange?.(new Map(this.edits));
    return true;
  }

  get canUndo(): boolean { return this.history.canUndo; }
  get canRedo(): boolean { return this.history.canRedo; }

  /** True when there are unsaved changes since the last saveVersion() call. */
  get hasUnsavedChanges(): boolean {
    return hashEdits(this.edits) !== this.lastSavedHash;
  }

  // ── Versioning ─────────────────────────────────────────────────────────────

  configureVersioning(store: PdfVersionStore): void {
    this.versionStore = store;
  }

  saveVersion(label?: string): PdfVersion | null {
    const hash = hashEdits(this.edits);
    if (hash === this.lastSavedHash) return null; // nothing changed
    const version = makeVersion(this.edits, label ?? `Version ${this.listVersions().length + 1}`);
    this.versionStore.save(version);
    this.lastSavedHash = hash;
    return version;
  }

  listVersions(): PdfVersion[] {
    return this.versionStore.list();
  }

  restoreVersion(versionId: string): boolean {
    const version = this.versionStore.get(versionId);
    if (!version) return false;
    restoreVersion(version, this.edits);
    this.history.clear();
    this.lastSavedHash = version.contentHash;
    this.onChange?.(new Map(this.edits));
    return true;
  }

  // ── Export ─────────────────────────────────────────────────────────────────

  /**
   * Export the PDF as a Uint8Array with all edits, annotations, and block
   * style overrides applied. Loads the text block index if not yet loaded.
   *
   * @throws {Error} if the raw bytes were detached or the PDF cannot be serialized.
   */
  async exportBytes(): Promise<Uint8Array> {
    if (this.rawBytes.byteLength === 0) {
      throw new Error(
        'PDF raw bytes are unavailable — the underlying ArrayBuffer was detached. ' +
        'Reload the document and try again.',
      );
    }
    await this.loadAllTextBlocks();
    return exportPdfWithEdits(
      this.rawBytes,
      this.textBlocks,
      this.edits,
      this.annotationsMap,
      this.blockStylesMap,
    );
  }

  async exportBlob(): Promise<Blob> {
    const bytes = await this.exportBytes();
    // Use a fresh ArrayBuffer copy so the Blob constructor never sees a detached buffer.
    const copy = new Uint8Array(bytes.byteLength);
    copy.set(bytes);
    return new Blob([copy.buffer], { type: 'application/pdf' });
  }

  // ── Cleanup ────────────────────────────────────────────────────────────────

  dispose(): void {
    this.proxy.destroy();
  }
}
