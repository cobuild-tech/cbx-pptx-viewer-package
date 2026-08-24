/**
 * DocxEditSession — the seam between a UI and the XML.
 *
 * Everything funnels through {@link commitParagraph}: snapshot the part, mutate
 * the cached XML, mark it dirty, re-parse. Mirrors pptx/edit/session.ts, with
 * two differences that follow from the format:
 *
 *  - the edit unit is one paragraph, not a whole text body, because a DOCX body
 *    is one long flow with no txBody-sized container;
 *  - the re-parse rebuilds styles and numbering from scratch, which is what
 *    keeps ordered lists numbered from 1 (see DocxDocument.parse).
 */
import { History } from '../../pptx/edit/history.js';
import type { DocxDocument } from '../document/document.js';
import type { DocxParagraph, DocxSection } from '../model.js';
import { child, type XmlNode } from '../../oxml/xml.js';
import { writeParagraphs, type DocxParaEdit } from './xmlWrite.js';

export interface DocxEditSessionOptions {
  /** Called after any change to the document (commit, undo or redo). */
  onChange?: () => void;
  /** Cap on retained undo snapshots. */
  historyLimit?: number;
}

export class DocxEditSession {
  private readonly doc: DocxDocument;
  private readonly history: History;
  private readonly onChange: DocxEditSessionOptions['onChange'];

  constructor(doc: DocxDocument, options: DocxEditSessionOptions = {}) {
    this.doc = doc;
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
    return this.doc.hasEdits;
  }

  /** True if this paragraph is in the main document part (not a header/footer). */
  isEditable(para: DocxParagraph): boolean {
    return this.doc.isEditable(para);
  }

  /**
   * Replace one source paragraph with the paragraphs it became, then re-parse.
   *
   * Returns the rebuilt sections, or undefined if the paragraph is not editable
   * or has no recorded source (in which case nothing is mutated).
   */
  commitParagraph(para: DocxParagraph, edits: DocxParaEdit[]): DocxSection[] | undefined {
    const src = this.doc.sourceOf(para);
    if (!src || !this.doc.isEditable(para)) return undefined;

    const body = this.documentBody();
    if (!body) return undefined;

    const before = this.doc.snapshotPart(src.part);
    if (before !== undefined) this.history.push({ part: src.part, xml: before });

    writeParagraphs(body, src.node, edits, (model) => this.doc.sourceOf(model));
    this.doc.markDirty();

    const sections = this.doc.rebuild();
    this.onChange?.();
    return sections;
  }

  undo(): DocxSection[] | undefined {
    return this.restore(this.history.undo((p) => this.doc.snapshotPart(p)));
  }

  redo(): DocxSection[] | undefined {
    return this.restore(this.history.redo((p) => this.doc.snapshotPart(p)));
  }

  private restore(snapshot: { part: string; xml: string } | undefined): DocxSection[] | undefined {
    if (!snapshot) return undefined;
    this.doc.restorePart(snapshot.part, snapshot.xml);
    // setPart drops the cached parse, so this re-reads the restored XML.
    const sections = this.doc.rebuild();
    this.onChange?.();
    return sections;
  }

  private documentBody(): XmlNode | undefined {
    return child(this.doc.documentXml(), 'body');
  }

  /** Re-zip the document with all edits applied. */
  exportBlob(): Blob {
    return this.doc.exportBlob();
  }

  /** Re-zip the document with all edits applied, as raw bytes. */
  exportBytes(): Uint8Array {
    return this.doc.toBytes();
  }
}
