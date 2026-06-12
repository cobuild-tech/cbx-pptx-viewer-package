/**
 * Deck loader — the top-level entry point.
 *
 * Reads the OPC package, walks presentation -> slides, and for each slide
 * assembles its layout/master/theme context, resolves the background, and builds
 * the shape tree. Returns a {@link Deck} the renderer/viewer consume. The Deck
 * also vends object URLs for embedded media and must be `dispose()`d to free them.
 */
import { OpcPackage } from '../opc/package.js';
import { RelType } from '../opc/relTypes.js';
import { child, children, attr, attrNum, type XmlNode } from '../xml.js';
import { emuToPx } from '../units.js';
import type { Fill, Slide, SlideSize, EmbeddedFont } from '../model.js';
import { parseTheme, type ColorContext, type Theme } from '../resolve/color.js';
import { type ParseScope, parseFill } from '../resolve/fill.js';
import { findColorEl, resolveColorEl } from '../resolve/color.js';
import { indexPlaceholders } from '../resolve/placeholders.js';
import { buildShapes, type SlideBuildCtx, type SlideScopes } from '../resolve/shape.js';

const DEFAULT_CLR_MAP: Record<string, string> = {
  bg1: 'lt1',
  tx1: 'dk1',
  bg2: 'lt2',
  tx2: 'dk2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
};

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

    const slides: Slide[] = slideParts.map((part, index) =>
      buildSlide(pkg, part, index),
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

function buildSlide(pkg: OpcPackage, slidePart: string, index: number): Slide {
  const slideXml = pkg.getXml(slidePart);
  const layoutPart = pkg.relByType(slidePart, RelType.SlideLayout)?.target;
  const layoutXml = layoutPart ? pkg.getXml(layoutPart) : undefined;
  const masterPart = layoutPart
    ? pkg.relByType(layoutPart, RelType.SlideMaster)?.target
    : undefined;
  const masterXml = masterPart ? pkg.getXml(masterPart) : undefined;
  const themePart = masterPart
    ? pkg.relByType(masterPart, RelType.Theme)?.target
    : undefined;

  const theme: Theme = parseTheme(themePart ? pkg.getXml(themePart) : undefined);
  const clrMap = readClrMap(masterXml);
  const colorCtx: ColorContext = { theme, clrMap };

  const scopes: SlideScopes = {
    slide: makeScope(pkg, slidePart, colorCtx),
    layout: makeScope(pkg, layoutPart ?? slidePart, colorCtx),
    master: makeScope(pkg, masterPart ?? slidePart, colorCtx),
  };

  const ctx: SlideBuildCtx = {
    colorCtx,
    theme,
    layoutPhs: indexPlaceholders(spTreeOf(layoutXml)),
    masterPhs: indexPlaceholders(spTreeOf(masterXml)),
    scopes,
  };
  const masterTxStyles = child(masterXml, 'txStyles');
  if (masterTxStyles) ctx.masterTxStyles = masterTxStyles;

  const background = resolveBackground(
    { xml: slideXml, scope: scopes.slide },
    { xml: layoutXml, scope: scopes.layout },
    { xml: masterXml, scope: scopes.master },
    colorCtx,
  );

  // Composite z-order: master decorations, then layout decorations, then the
  // slide's own shapes on top. Placeholders on the layout/master are prompts and
  // are skipped (the slide renders the real placeholder content).
  const shapes: Slide['shapes'] = [];
  const showMaster = attr(layoutXml, 'showMasterSp') !== '0';
  const masterSpTree = spTreeOf(masterXml);
  if (showMaster && masterSpTree) {
    shapes.push(...buildShapes(masterSpTree, ctx, scopes.master, { skipPlaceholders: true }));
  }
  const layoutSpTree = spTreeOf(layoutXml);
  if (layoutSpTree) {
    shapes.push(...buildShapes(layoutSpTree, ctx, scopes.layout, { skipPlaceholders: true }));
  }
  const slideSpTree = spTreeOf(slideXml);
  if (slideSpTree) {
    shapes.push(...buildShapes(slideSpTree, ctx, scopes.slide));
  }

  return { index, background, shapes, part: slidePart };
}

function spTreeOf(xml: XmlNode | undefined): XmlNode | undefined {
  return child(child(xml, 'cSld'), 'spTree');
}

function readClrMap(masterXml: XmlNode | undefined): Record<string, string> {
  const clrMapEl = child(masterXml, 'clrMap');
  if (!clrMapEl) return { ...DEFAULT_CLR_MAP };
  const map: Record<string, string> = {};
  for (const [k, v] of Object.entries(clrMapEl.attrs)) map[k] = v;
  return { ...DEFAULT_CLR_MAP, ...map };
}

function makeScope(pkg: OpcPackage, partPath: string, colorCtx: ColorContext): ParseScope {
  return {
    colorCtx,
    resolveImage: (rId) => pkg.resolveRel(partPath, rId)?.target,
    resolveHyperlink: (rId) => {
      const rel = pkg.resolveRel(partPath, rId);
      if (!rel) return undefined;
      return rel.mode === 'External' ? rel.rawTarget : rel.target;
    },
  };
}

interface BgSource {
  xml: XmlNode | undefined;
  scope: ParseScope;
}

function resolveBackground(
  slide: BgSource,
  layout: BgSource,
  master: BgSource,
  colorCtx: ColorContext,
): Fill {
  for (const src of [slide, layout, master]) {
    const bg = child(child(src.xml, 'cSld'), 'bg');
    if (!bg) continue;
    const bgPr = child(bg, 'bgPr');
    if (bgPr) {
      const fill = parseFill(bgPr, src.scope);
      if (fill) return fill;
    }
    const bgRef = child(bg, 'bgRef');
    if (bgRef) {
      const color = resolveColorEl(findColorEl(bgRef), colorCtx);
      if (color) return { type: 'solid', color };
    }
  }
  // Default to the theme's lt1 (typically white).
  const key = colorCtx.clrMap['bg1'] ?? 'lt1';
  return { type: 'solid', color: { hex: colorCtx.theme.colors[key] ?? 'FFFFFF' } };
}
