/**
 * Deck loader — the top-level entry point.
 *
 * Reads the OPC package, walks presentation -> slides, and delegates per-slide
 * assembly to the slides slice ({@link buildSlide}). Returns a {@link Deck} the
 * renderer/viewer consume. The Deck also vends object URLs for embedded media
 * and must be `dispose()`d to free them.
 */
import { OpcPackage } from '../../oxml/package.js';
import { RelType } from '../relTypes.js';
import { child, children, attr, attrNum, type XmlNode } from '../../oxml/xml.js';
import { emuToPx } from '../../oxml/units.js';
import type { Slide, SlideSize, EmbeddedFont } from '../model.js';
import { buildSlide } from '../slides/slide.js';

export class Deck {
  readonly size: SlideSize;
  readonly slides: Slide[];
  /** Fonts embedded in the package, to install via the FontFace API. */
  readonly embeddedFonts: EmbeddedFont[];
  private readonly pkg: OpcPackage;
  private readonly urlCache = new Map<string, string>();

  private constructor(
    pkg: OpcPackage,
    size: SlideSize,
    slides: Slide[],
    embeddedFonts: EmbeddedFont[],
  ) {
    this.pkg = pkg;
    this.size = size;
    this.slides = slides;
    this.embeddedFonts = embeddedFonts;
  }

  static load(data: ArrayBuffer | Uint8Array): Deck {
    const pkg = OpcPackage.load(data);
    const presPart = pkg.relByType('', RelType.OfficeDocument)?.target;
    if (!presPart) throw new Error('Not a PresentationML package: no presentation part.');
    const presXml = pkg.getXml(presPart);
    if (!presXml) throw new Error('presentation.xml is missing or empty.');

    const size = readSlideSize(presXml);
    const slideParts = readSlideOrder(pkg, presPart, presXml);

    // Presentation-wide default text style — the base list-level formatting
    // (indents, alignment, sizes) for text in non-placeholder text boxes.
    const defaultTextStyle = child(presXml, 'defaultTextStyle');

    const slides: Slide[] = slideParts.map((part, index) =>
      buildSlide(pkg, part, index, defaultTextStyle),
    );
    const embeddedFonts = readEmbeddedFonts(pkg, presPart, presXml);
    return new Deck(pkg, size, slides, embeddedFonts);
  }

  /** Raw bytes of an embedded font part, for FontFace registration. */
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
    const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
    this.urlCache.set(part, url);
    return url;
  }

  /** Release all object URLs created for media. */
  dispose(): void {
    if (typeof URL !== 'undefined') {
      for (const url of this.urlCache.values()) URL.revokeObjectURL(url);
    }
    this.urlCache.clear();
  }
}

function readSlideSize(presXml: XmlNode): SlideSize {
  const sldSz = child(presXml, 'sldSz');
  return {
    wPx: emuToPx(attrNum(sldSz, 'cx') ?? 9144000),
    hPx: emuToPx(attrNum(sldSz, 'cy') ?? 6858000),
  };
}

/**
 * Parse `<p:embeddedFontLst>` into typeface -> font-part mappings. Each
 * embedded font may carry regular/bold/italic/boldItalic variants, referenced
 * by relationship id against presentation.xml's rels.
 */
function readEmbeddedFonts(
  pkg: OpcPackage,
  presPart: string,
  presXml: XmlNode,
): EmbeddedFont[] {
  const variants: Array<[string, number, 'normal' | 'italic']> = [
    ['regular', 400, 'normal'],
    ['bold', 700, 'normal'],
    ['italic', 400, 'italic'],
    ['boldItalic', 700, 'italic'],
  ];
  const out: EmbeddedFont[] = [];
  for (const ef of children(child(presXml, 'embeddedFontLst'), 'embeddedFont')) {
    const typeface = attr(child(ef, 'font'), 'typeface');
    if (!typeface) continue;
    const faces = [];
    for (const [tag, weight, style] of variants) {
      const rId = attr(child(ef, tag), 'r:id');
      const part = rId ? pkg.resolveRel(presPart, rId)?.target : undefined;
      if (part) faces.push({ weight, style, part });
    }
    if (faces.length) out.push({ typeface, faces });
  }
  return out;
}

function readSlideOrder(pkg: OpcPackage, presPart: string, presXml: XmlNode): string[] {
  const parts: string[] = [];
  for (const sldId of children(child(presXml, 'sldIdLst'), 'sldId')) {
    // sldId carries a numeric `id` and the relationship in `r:id`.
    const rId = attr(sldId, 'r:id');
    const rel = rId ? pkg.resolveRel(presPart, rId) : undefined;
    if (rel) parts.push(rel.target);
  }
  return parts;
}
