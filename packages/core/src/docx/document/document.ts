/**
 * DocxDocument loader — the top-level entry point for DOCX files.
 *
 * Mirrors the PPTX Deck API: load() unzips the OPC package, resolves the
 * document part, parses styles/numbering/theme, and delegates body parsing to
 * buildPages(). The returned DocxDocument vends object URLs for embedded media
 * and must be dispose()d to free them.
 */
import { OpcPackage } from '../../oxml/package.js';
import { DocxRelType } from '../relTypes.js';
import type { DocxPage, EmbeddedFont } from '../model.js';
import { StyleMap } from '../styles/styles.js';
import { NumberingMap } from '../numbering/numbering.js';
import { buildPages } from './body.js';

export class DocxDocument {
  readonly pages: DocxPage[];
  readonly embeddedFonts: EmbeddedFont[];
  private readonly pkg: OpcPackage;
  private readonly urlCache = new Map<string, string>();

  private constructor(pkg: OpcPackage, pages: DocxPage[], embeddedFonts: EmbeddedFont[]) {
    this.pkg = pkg;
    this.pages = pages;
    this.embeddedFonts = embeddedFonts;
  }

  static load(data: ArrayBuffer | Uint8Array): DocxDocument {
    const pkg = OpcPackage.load(data);

    // Find the document.xml part via the root relationship.
    const docPart = pkg.relByType('', DocxRelType.OfficeDocument)?.target;
    if (!docPart) throw new Error('Not a WordprocessingML package: no document part.');
    const docXml = pkg.getXml(docPart);
    if (!docXml) throw new Error('document.xml is missing or empty.');

    // The document body is the first child of <w:document>.
    const bodyEl = findChild(docXml, 'body');
    if (!bodyEl) throw new Error('document.xml has no <w:body>.');

    // Load styles.xml (optional — fall back to empty map).
    const stylesPart = pkg.relByType(docPart, DocxRelType.Styles)?.target;
    const stylesXml = stylesPart ? pkg.getXml(stylesPart) : undefined;
    const styles = StyleMap.parse(stylesXml);

    // Load numbering.xml (optional).
    const numberingPart = pkg.relByType(docPart, DocxRelType.Numbering)?.target;
    const numberingXml = numberingPart ? pkg.getXml(numberingPart) : undefined;
    const numbering = NumberingMap.parse(numberingXml);

    const pages = buildPages(bodyEl, { pkg, docPart, styles, numbering });

    return new DocxDocument(pkg, pages, []);
  }

  /** Object URL for an embedded media part (cached). Browser only. */
  imageUrl(part: string): string | undefined {
    const cached = this.urlCache.get(part);
    if (cached) return cached;
    const bytes = this.pkg.getBytes(part);
    if (!bytes || typeof URL === 'undefined' || typeof Blob === 'undefined') return undefined;
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

/** Find a direct child by local name (namespace-prefix agnostic). */
function findChild(node: import('../../oxml/xml.js').XmlNode, localName: string): import('../../oxml/xml.js').XmlNode | undefined {
  return node.children.find((c) => {
    const n = c.name;
    return n === localName || n.endsWith(`:${localName}`);
  });
}
