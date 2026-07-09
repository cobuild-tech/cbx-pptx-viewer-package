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
import { parseBody } from './body.js';
import { child } from '../../oxml/xml.js';
import type { DocxSection, EmbeddedFont } from '../model.js';
import type { ParseContext } from './context.js';

export class DocxDocument {
  /** Parsed sections; the viewer's paginator flows these into fixed-size pages. */
  readonly sections: DocxSection[];
  /** Fonts embedded in the package (empty in v1 — reserved for parity with Deck). */
  readonly embeddedFonts: EmbeddedFont[];
  private readonly pkg: OpcPackage;
  private readonly urlCache = new Map<string, string>();

  private constructor(pkg: OpcPackage, sections: DocxSection[], embeddedFonts: EmbeddedFont[]) {
    this.pkg = pkg;
    this.sections = sections;
    this.embeddedFonts = embeddedFonts;
  }

  static load(data: ArrayBuffer | Uint8Array): DocxDocument {
    const pkg = OpcPackage.load(data);
    const docPart = pkg.relByType('', DocxRelType.OfficeDocument)?.target;
    if (!docPart) throw new Error('Not a WordprocessingML package: no officeDocument part.');
    const docXml = pkg.getXml(docPart);
    if (!docXml) throw new Error('word/document.xml is missing or empty.');

    const stylesPart = pkg.relByType(docPart, DocxRelType.Styles)?.target;
    const numberingPart = pkg.relByType(docPart, DocxRelType.Numbering)?.target;
    const styles = StyleTable.parse(stylesPart ? pkg.getXml(stylesPart) : undefined);
    const numbering = Numbering.parse(numberingPart ? pkg.getXml(numberingPart) : undefined);

    const makeCtx = (partPath: string): ParseContext => ({
      styles,
      numbering,
      partPath,
      rel: (relId) => (relId ? pkg.resolveRel(partPath, relId) : undefined),
      getPartXml: (part) => pkg.getXml(part),
      forPart: (part) => makeCtx(part),
    });
    const ctx = makeCtx(docPart);

    const body = child(docXml, 'body');
    const sections = body ? parseBody(body, ctx) : [];
    return new DocxDocument(pkg, sections, []);
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
