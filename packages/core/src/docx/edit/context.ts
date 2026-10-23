/**
 * DocxEditContext — the bridge between the drawn pages and the XML.
 *
 * The renderer stamps each paragraph and run with a key; the editor turns a
 * key back into what it stands for. Keys name **XML nodes**, not model
 * objects: a `<w:p>`, and a run's source element plus how far into its text a
 * layout copy starts. XML nodes outlive the re-parse every edit triggers, so a
 * page whose content did not change keeps valid stamps and need not be drawn
 * again — which is what keeps typing fast in a long document.
 */
import type { DocxDocument } from '../document/document.js';
import type { DocxBlock, DocxDrawing, DocxParagraph, DocxRun } from '../model.js';
import type { XmlNode } from '../../oxml/xml.js';
import { originOf } from '../origin.js';

/**
 * What the DOCX renderer needs in order to make text editable. Absent for a
 * read-only render, which is the default.
 */
export interface DocxEditRenderContext {
  /** Stable key identifying a paragraph or run (or a layout copy of one). */
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
  private ids = new WeakMap<XmlNode, number>();
  private nodes = new Map<number, XmlNode>();
  private nextId = 0;
  /** The latest model object drawn under each key, for {@link resolve}. */
  private byKey = new Map<string, object>();

  constructor(doc: DocxDocument) {
    this.doc = doc;
  }

  /** Forget the model objects of the last render (node ids are kept). */
  reset(): void {
    this.byKey = new Map();
  }

  private idOf(node: XmlNode): number {
    let id = this.ids.get(node);
    if (id === undefined) {
      id = this.nextId++;
      this.ids.set(node, id);
      this.nodes.set(id, node);
    }
    return id;
  }

  key(model: object): string {
    const { of, off } = originOf(model);
    const src = this.doc.sourceOf(of);
    let k: string;
    // Generated content with no XML of its own (DrawingML text in a shape)
    // still gets a key, unique to the object, though nothing edits it.
    if (!src) k = `x${this.idOf(of as unknown as XmlNode)}`;
    else if ('kind' in of && (of as DocxParagraph).kind === 'paragraph') k = `p${this.idOf(src.node)}`;
    else k = `r${this.idOf(src.node)}:${off}`;
    this.byKey.set(k, model);
    return k;
  }

  /** The model object last drawn under a key (a layout copy, possibly), if any. */
  resolve(key: string | null | undefined): object | undefined {
    return key ? this.byKey.get(key) : undefined;
  }

  /** The `<w:p>` a paragraph key names. */
  paragraphNode(key: string | null | undefined): XmlNode | undefined {
    const m = key ? /^p(\d+)$/.exec(key) : null;
    return m ? this.node(Number(m[1])) : undefined;
  }

  /** The source element a run key names, and the offset of its text in that run. */
  runNode(key: string | null | undefined): { node: XmlNode; off: number } | undefined {
    const m = key ? /^r(\d+):(\d+)$/.exec(key) : null;
    const node = m ? this.node(Number(m[1])) : undefined;
    return node ? { node, off: Number(m![2]) } : undefined;
  }

  private node(id: number): XmlNode | undefined {
    return this.nodes.get(id);
  }

  /**
   * Which nodes a page's stamps would name: equal for two renders exactly when
   * their stamps would be, so a page can be kept rather than redrawn.
   */
  fingerprint(blocks: DocxBlock[] | undefined): string {
    if (!blocks) return '';
    const out: string[] = [];
    const drawing = (d: DocxDrawing | undefined) => {
      if (!d) return;
      if (d.kind === 'shape') walk(d.content);
      else if (d.kind === 'group') d.items.forEach((i) => drawing(i.drawing));
    };
    const walk = (list: DocxBlock[]) => {
      for (const b of list) {
        if (b.kind === 'paragraph') {
          out.push(this.key(b));
          for (const r of b.runs) {
            out.push(this.key(r));
            drawing(r.drawing);
          }
          for (const a of b.anchors ?? []) drawing(a.drawing);
        } else if (b.kind === 'table') {
          for (const row of b.rows) for (const c of row) if (c) walk(c.content);
        }
      }
    };
    walk(blocks);
    return out.join(' ');
  }

  /** Forget every node id — after an undo replaced the whole XML tree. */
  forgetNodes(): void {
    this.ids = new WeakMap();
    this.nodes = new Map();
  }

  editable(para: DocxParagraph): boolean {
    // A paragraph split across pages is drawn as copies; they edit the original.
    return this.doc.isEditable(originOf(para).of);
  }

  isField(run: DocxRun): boolean {
    return run.fieldCode !== undefined;
  }
}
