/**
 * Per-slide assembly.
 *
 * For one slide part this resolves its layout/master/theme context, builds the
 * color map and resolution scopes, resolves the background, and composites the
 * shape tree (master decorations, then layout decorations, then the slide's own
 * shapes on top). Returns a {@link Slide} for the renderer/viewer.
 */
import { OpcPackage } from '../../oxml/package.js';
import { RelType } from '../relTypes.js';
import { child, attr, type XmlNode } from '../../oxml/xml.js';
import type { Fill, Slide } from '../model.js';
import { parseTheme, type ColorContext, type Theme, findColorEl, resolveColorEl } from '../color.js';
import type { ParseScope } from '../scope.js';
import { parseFill } from '../shapes/fill.js';
import { indexPlaceholders } from '../shapes/placeholders.js';
import { buildShapes, type SlideBuildCtx, type SlideScopes } from '../shapes/shape.js';

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

export function buildSlide(pkg: OpcPackage, slidePart: string, index: number): Slide {
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
