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

const UNDECODABLE_IMAGE_TYPES = new Set(['image/x-emf', 'image/x-wmf', 'image/emf', 'image/wmf']);

const PPTX_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation';

/** Where a model object was parsed from, for the edit layer. */
export interface ModelSource {
  node: XmlNode;
  /** Package part the node lives in (slide, layout or master). */
  part: string;
}

export class Deck {
  readonly size: SlideSize;
  readonly slides: Slide[];
  /** Fonts embedded in the package, to install via the FontFace API. */
  readonly embeddedFonts: EmbeddedFont[];
  private readonly pkg: OpcPackage;
  private readonly urlCache = new Map<string, string>();
  /**
   * Model object -> the XML node it was parsed from. Kept off the model itself
   * so the render-agnostic types stay pure; a WeakMap means entries die with
   * the slide when it is rebuilt.
   */
  private sources = new WeakMap<object, ModelSource>();
  private readonly slideParts: string[];
  private readonly defaultTextStyle: XmlNode | undefined;

  private constructor(
    pkg: OpcPackage,
    size: SlideSize,
    slides: Slide[],
    embeddedFonts: EmbeddedFont[],
    slideParts: string[],
    defaultTextStyle: XmlNode | undefined,
  ) {
    this.pkg = pkg;
    this.size = size;
    this.slides = slides;
    this.embeddedFonts = embeddedFonts;
    this.slideParts = slideParts;
    this.defaultTextStyle = defaultTextStyle;
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

    const embeddedFonts = readEmbeddedFonts(pkg, presPart, presXml);
    const deck = new Deck(pkg, size, [], embeddedFonts, slideParts, defaultTextStyle);
    slideParts.forEach((part, index) => {
      deck.slides.push(deck.build(part, index));
    });
    return deck;
  }

  // ─── Editing ───────────────────────────────────────────────────────────────

  /** Build one slide, capturing every text node's XML source as it parses. */
  private build(part: string, index: number): Slide {
    return buildSlide(this.pkg, part, index, this.defaultTextStyle, (model, node, srcPart) => {
      this.sources.set(model, { node, part: srcPart });
    });
  }

  /**
   * The XML node a parsed model object (text body, paragraph or run) came from,
   * and the part it lives in. Undefined for anything not recorded.
   */
  sourceOf(model: object): ModelSource | undefined {
    return this.sources.get(model);
  }

  /**
   * True if `model` belongs to the slide's own part rather than being inherited
   * from its layout or master. Only slide-owned text is editable — editing
   * inherited text would change every slide sharing that layout/master.
   */
  isEditable(slideIndex: number, model: object): boolean {
    const slide = this.slides[slideIndex];
    const src = this.sources.get(model);
    return !!slide && !!src && src.part === slide.part;
  }

  /**
   * Re-parse a slide from its (possibly mutated) XML and replace it in place.
   * `pkg.getXml` returns the cached node the edit layer mutated, so this picks
   * up edits without re-reading the zip.
   */
  rebuildSlide(index: number): Slide | undefined {
    const part = this.slideParts[index];
    if (part === undefined) return undefined;
    const rebuilt = this.build(part, index);
    this.slides[index] = rebuilt;
    return rebuilt;
  }

  /** Mark a slide's part as mutated so export re-serializes it. */
  markSlideDirty(index: number): void {
    const part = this.slideParts[index];
    if (part !== undefined) this.pkg.markDirty(part);
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

  /** Re-zip the deck, edits included, as .pptx bytes. */
  toBytes(): Uint8Array {
    return this.pkg.toBytes();
  }

  /** Re-zip the deck, edits included, as a .pptx Blob. */
  exportBlob(): Blob {
    return this.pkg.toBlob(PPTX_CONTENT_TYPE);
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
    // EMF/WMF are vector metafiles no browser can decode via <img>; loading them
    // yields a broken-image box instead of nothing, which shows through crop/overflow
    // clipping as stray artifacts. Skip so callers fall back to rendering nothing.
    if (UNDECODABLE_IMAGE_TYPES.has(type)) return undefined;
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
