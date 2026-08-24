/**
 * DocxDocument loader — the top-level entry point.
 *
 * Reads the OPC package, resolves styles + numbering, parses word/document.xml
 * into a paginated model, and returns a {@link DocxDocument} the renderer/viewer
 * consume. Mirrors the PPTX {@link Deck}: it vends object URLs for embedded
 * media and must be `dispose()`d to free them.
 *
 * DOCX pipeline: read (OPC) -> parse (XML -> model) -> paginate (sections) -> render (DOM)
 */
import { OpcPackage } from '../../oxml/package.js';
import { DocxRelType } from '../relTypes.js';
import { StyleTable } from '../styles/styles.js';
import { Numbering } from '../numbering/numbering.js';
import { parseBody, parseBlocks } from './body.js';
import { child, type XmlNode } from '../../oxml/xml.js';
import type { DocxSection, EmbeddedFont } from '../model.js';
import type { DocxSource, ParseContext } from './context.js';

export type { DocxSource } from './context.js';

const DOCX_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

export class DocxDocument {
  /** Parsed sections; the viewer's paginator flows these into fixed-size pages. */
  sections: DocxSection[];
  /** Fonts embedded in the package (empty in v1 — reserved for parity with Deck). */
  readonly embeddedFonts: EmbeddedFont[];
  private readonly pkg: OpcPackage;
  private readonly docPart: string;
  private readonly urlCache = new Map<string, string>();
  /**
   * Model object -> the XML node it was parsed from. Off the model itself so the
   * render-agnostic types stay pure; a WeakMap means entries die with the
   * sections when the document is rebuilt.
   */
  private sources = new WeakMap<object, DocxSource>();

  private constructor(pkg: OpcPackage, docPart: string, embeddedFonts: EmbeddedFont[]) {
    this.pkg = pkg;
    this.docPart = docPart;
    this.embeddedFonts = embeddedFonts;
    this.sections = [];
  }

  static load(data: ArrayBuffer | Uint8Array): DocxDocument {
    const pkg = OpcPackage.load(data);
    const docPart = pkg.relByType('', DocxRelType.OfficeDocument)?.target;
    if (!docPart) throw new Error('Not a WordprocessingML package: no officeDocument part.');
    if (!pkg.getXml(docPart)) throw new Error('word/document.xml is missing or empty.');

    const doc = new DocxDocument(pkg, docPart, []);
    doc.sections = doc.parse();
    return doc;
  }

  /**
   * Parse (or re-parse) the document part into sections, recording every
   * paragraph's and run's XML source as it goes.
   *
   * Styles and numbering are rebuilt on every call *by design*: `Numbering`
   * carries live counters that `marker()` advances as it walks the body, so
   * reusing one across parses would render an ordered list as "5. 6. 7." the
   * second time round.
   */
  private parse(): DocxSection[] {
    const { pkg, docPart } = this;
    const docXml = pkg.getXml(docPart);
    if (!docXml) return [];

    const stylesPart = pkg.relByType(docPart, DocxRelType.Styles)?.target;
    const numberingPart = pkg.relByType(docPart, DocxRelType.Numbering)?.target;
    const styles = StyleTable.parse(stylesPart ? pkg.getXml(stylesPart) : undefined);
    const numbering = Numbering.parse(numberingPart ? pkg.getXml(numberingPart) : undefined);

    const makeCtx = (partPath: string): ParseContext => {
      const c: ParseContext = {
        styles,
        numbering,
        partPath,
        rel: (relId) => (relId ? pkg.resolveRel(partPath, relId) : undefined),
        getPartXml: (part) => pkg.getXml(part),
        forPart: (part) => makeCtx(part),
        parseBlocks: (container) => parseBlocks(container, c),
        // The context knows its own part, so header/footer content is
        // automatically distinguishable from the main document body.
        recordSource: (model, node, owner) => {
          this.sources.set(model, { node, part: partPath, ...(owner ? { owner } : {}) });
        },
      };
      return c;
    };

    const body = child(docXml, 'body');
    return body ? parseBody(body, makeCtx(docPart)) : [];
  }

  // ─── Editing ───────────────────────────────────────────────────────────────

  /**
   * The XML node a parsed paragraph or run came from. For a run this is its
   * `<w:t>`, with the owning `<w:r>` on `.owner`.
   */
  sourceOf(model: object): DocxSource | undefined {
    return this.sources.get(model);
  }

  /**
   * True if `model` belongs to the main document part. Header and footer
   * content is parsed from its own part and stays read-only — editing it would
   * change every page that shares it.
   */
  isEditable(model: object): boolean {
    return this.sources.get(model)?.part === this.docPart;
  }

  /** The main document part path (`word/document.xml`). */
  get mainPart(): string {
    return this.docPart;
  }

  /** Parsed XML root of the main document part (the live, mutable node). */
  documentXml(): XmlNode | undefined {
    return this.pkg.getXml(this.docPart);
  }

  /**
   * Re-parse the document from its (possibly mutated) XML, replacing
   * `sections`. `pkg.getXml` returns the cached node the edit layer mutated, so
   * this picks up edits without re-reading the zip.
   */
  rebuild(): DocxSection[] {
    this.sources = new WeakMap();
    this.sections = this.parse();
    return this.sections;
  }

  /** Mark the document part as mutated so export re-serializes it. */
  markDirty(): void {
    this.pkg.markDirty(this.docPart);
  }

  /** True if any part has been edited. */
  get hasEdits(): boolean {
    return this.pkg.hasEdits;
  }

  /** Current XML text of a part — used to snapshot for undo. */
  snapshotPart(part: string): string | undefined {
    return this.pkg.serializePart(part);
  }

  /** Restore a part from a snapshot taken by {@link snapshotPart}. */
  restorePart(part: string, xml: string): void {
    this.pkg.setPart(part, xml);
  }

  /** Re-zip the document, edits included, as .docx bytes. */
  toBytes(): Uint8Array {
    return this.pkg.toBytes();
  }

  /** Re-zip the document, edits included, as a .docx Blob. */
  exportBlob(): Blob {
    return this.pkg.toBlob(DOCX_CONTENT_TYPE);
  }

  /** Raw bytes of a part (e.g. an embedded font), for FontFace registration. */
  fontBytes(part: string): Uint8Array | undefined {
    return this.pkg.getBytes(part);
  }

  /** Object URL for an embedded media part (cached). Browser only. */
  imageUrl(part: string): string | undefined {
    const cached = this.urlCache.get(part);
    if (cached) return cached;
    const bytes = this.pkg.getBytes(part);
    if (!bytes || typeof URL === 'undefined' || typeof Blob === 'undefined') return undefined;
    const type = this.pkg.contentType(part) ?? 'application/octet-stream';
    const url = URL.createObjectURL(new Blob([bytes as unknown as BlobPart], { type }));
    this.urlCache.set(part, url);
    return url;
  }

  /** Revoke all media object URLs. Call when the document is no longer displayed. */
  dispose(): void {
    if (typeof URL !== 'undefined') {
      for (const url of this.urlCache.values()) URL.revokeObjectURL(url);
    }
    this.urlCache.clear();
  }
}
