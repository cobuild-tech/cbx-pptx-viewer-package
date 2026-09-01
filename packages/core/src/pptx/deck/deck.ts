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
import { deleteSlide as deleteSlideParts, planSlideDeletion } from '../edit/slideOps.js';
import type { Snapshot } from '../../oxml/edit/history.js';

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
  /** Slide part paths in running order — mutated when slides are deleted. */
  private slideParts: string[];
  private defaultTextStyle: XmlNode | undefined;
  private readonly presPart: string;

  private constructor(
    pkg: OpcPackage,
    presPart: string,
    size: SlideSize,
    slides: Slide[],
    embeddedFonts: EmbeddedFont[],
    slideParts: string[],
    defaultTextStyle: XmlNode | undefined,
  ) {
    this.pkg = pkg;
    this.presPart = presPart;
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
    const deck = new Deck(pkg, presPart, size, [], embeddedFonts, slideParts, defaultTextStyle);
    slideParts.forEach((part, index) => {
      deck.slides.push(deck.build(part, index));
    });
    return deck;
  }

  /** Path of the presentation part — the running order lives here. */
  get presentationPart(): string {
    return this.presPart;
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

  /**
   * Re-read the running order from presentation.xml and rebuild every slide.
   *
   * Needed after a *structural* change (a slide deleted), where the slide list
   * itself moved rather than one slide's contents. Like {@link rebuildSlide},
   * this invalidates every model object previously handed out.
   */
  rebuild(): void {
    const presXml = this.pkg.getXml(this.presPart);
    this.defaultTextStyle = presXml ? child(presXml, 'defaultTextStyle') : undefined;
    this.slideParts = presXml ? readSlideOrder(this.pkg, this.presPart, presXml) : [];
    // A fresh source map: the old entries point at nodes of slides that may no
    // longer exist, and stale identity here would silently mis-address edits.
    this.sources = new WeakMap();
    this.slides.length = 0;
    this.slideParts.forEach((part, index) => {
      this.slides.push(this.build(part, index));
    });
  }

  /**
   * True if the slide at `index` can be deleted. A presentation must keep at
   * least one slide — PowerPoint refuses to open a deck with an empty
   * `<p:sldIdLst>`.
   */
  canDeleteSlide(index: number): boolean {
    return this.slides.length > 1 && index >= 0 && index < this.slideParts.length;
  }

  /**
   * Every part a deletion of slide `index` would touch, as snapshots of their
   * current state. Take these *before* calling {@link deleteSlide} to make the
   * deletion undoable; parts that will be removed are snapshotted with their
   * content, parts that do not yet exist are marked absent.
   */
  slideDeletionSnapshots(index: number): Snapshot[] {
    const part = this.slideParts[index];
    if (part === undefined) return [];
    const plan = planSlideDeletion(this.pkg, this.presPart, part);
    if (!plan) return [];
    return [...plan.changed, ...plan.removed].map((p) => this.snapshot(p));
  }

  /**
   * Delete the slide at `index` and rebuild the deck. Returns false (having
   * changed nothing) if the slide cannot be deleted.
   */
  deleteSlide(index: number): boolean {
    if (!this.canDeleteSlide(index)) return false;
    const part = this.slideParts[index]!;
    if (!deleteSlideParts(this.pkg, this.presPart, part)) return false;
    this.rebuild();
    return true;
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

  /**
   * A part's current state as an undo {@link Snapshot}, including the case where
   * the part does not exist — which {@link snapshotPart} cannot express and
   * which redoing a slide deletion depends on.
   */
  snapshot(part: string): Snapshot {
    const xml = this.pkg.serializePart(part);
    return xml === undefined ? { part, xml: '', absent: true } : { part, xml };
  }

  /** Restore a part from a snapshot taken by {@link snapshotPart}. */
  restorePart(part: string, xml: string): void {
    this.pkg.setPart(part, xml);
  }

  /** Apply an undo {@link Snapshot}: write the part back, or delete it again. */
  restore(snapshot: Snapshot): void {
    if (snapshot.absent) this.pkg.deletePart(snapshot.part);
    else this.pkg.setPart(snapshot.part, snapshot.xml);
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
