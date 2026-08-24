/**
 * DocxEditContext — the bridge between model objects and the DOM.
 *
 * The renderer needs a string it can put in an attribute; the reconciler needs
 * to turn that string back into the model object. This holds both directions
 * for one render pass, and answers the renderer's questions about what may be
 * edited. Mirrors pptx/edit/context.ts.
 */
import type { DocxDocument } from '../document/document.js';
import type { DocxParagraph, DocxRun } from '../model.js';

/**
 * What the DOCX renderer needs in order to make text editable. Absent for a
 * read-only render, which is the default.
 */
export interface DocxEditRenderContext {
  /** Stable key identifying a model object for this render pass. */
  key(model: object): string;
  /** True if this paragraph is in the main document part and may be edited. */
  editable(para: DocxParagraph): boolean;
  /**
   * True for runs whose text Word generates (`PAGE`, `STYLEREF`, …). They
   * render but must not be typed into, since Word regenerates them.
   */
  isField(run: DocxRun): boolean;
}

export class DocxEditContext implements DocxEditRenderContext {
  private readonly doc: DocxDocument;
  private byKey = new Map<string, object>();
  private keys = new WeakMap<object, string>();
  private next = 0;

  constructor(doc: DocxDocument) {
    this.doc = doc;
  }

  /** Forget every key. Call before re-rendering. */
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

  editable(para: DocxParagraph): boolean {
    return this.doc.isEditable(para);
  }

  isField(run: DocxRun): boolean {
    return run.fieldCode !== undefined;
  }
}
