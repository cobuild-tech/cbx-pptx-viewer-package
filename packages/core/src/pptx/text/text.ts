/**
 * Text-body parser: `<p:txBody>` / `<a:txBody>` -> {@link TextBody}.
 *
 * Walks paragraphs and their interleaved runs / breaks / fields in document
 * order, resolving each run and paragraph through the shape's
 * {@link TextStyleChain} (which encodes the full inheritance chain).
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { emuToPx } from '../../oxml/units.js';
import type { Paragraph, TextBody, TextRun, TextVert, VerticalAnchor } from '../model.js';
import type { ColorContext } from '../color.js';
import type { TextStyleChain } from './textStyles.js';
import type { ParseScope } from '../scope.js';

const DEFAULT_INSET = { l: emuToPx(91440), t: emuToPx(45720), r: emuToPx(91440), b: emuToPx(45720) };

export function parseTextBody(
  txBody: XmlNode,
  chain: TextStyleChain,
  ctx: ColorContext,
  scope: ParseScope,
  inheritedBodyPr: (XmlNode | undefined)[] = [],
): TextBody {
  const bodyPr = effectiveBodyPr([child(txBody, 'bodyPr'), ...inheritedBodyPr]);
  const body: TextBody = {
    paragraphs: [],
    anchor: readAnchor(bodyPr),
    wrap: attr(bodyPr, 'wrap') !== 'none',
    insets: {
      l: insetPx(bodyPr, 'lIns', DEFAULT_INSET.l),
      t: insetPx(bodyPr, 'tIns', DEFAULT_INSET.t),
      r: insetPx(bodyPr, 'rIns', DEFAULT_INSET.r),
      b: insetPx(bodyPr, 'bIns', DEFAULT_INSET.b),
    },
  };

  const vert = readVert(bodyPr);
  if (vert) body.vert = vert;

  const normAutofit = child(bodyPr, 'normAutofit');
  if (normAutofit) {
    const fs = attrNum(normAutofit, 'fontScale');
    const lr = attrNum(normAutofit, 'lnSpcReduction');
    if (fs !== undefined) body.fontScale = fs / 100000;
    if (lr !== undefined) body.lnSpcReductionPct = lr / 100000;
  }

  for (const p of children(txBody, 'p')) {
    body.paragraphs.push(parseParagraph(p, chain, ctx, scope));
  }
  scope.recordSource?.(body, txBody);
  return body;
}

const AUTOFITS = new Set(['noAutofit', 'normAutofit', 'spAutoFit']);

/**
 * A placeholder's `<a:bodyPr>` inherits from the layout's and then the
 * master's, attribute by attribute: a slide title stating only `<a:bodyPr/>`
 * still takes its anchor and insets from the master. The autofit mode is one
 * choice, taken from the most specific body that makes it. Most specific first.
 */
function effectiveBodyPr(chain: (XmlNode | undefined)[]): XmlNode | undefined {
  const prs = chain.filter((n): n is XmlNode => n !== undefined);
  if (prs.length <= 1) return prs[0];
  const attrs: Record<string, string> = {};
  for (const n of [...prs].reverse()) Object.assign(attrs, n.attrs);
  const fit = prs.map((n) => n.children.find((c) => AUTOFITS.has(localName(c.name)))).find(Boolean);
  return { name: prs[0]!.name, attrs, children: fit ? [fit] : [], text: '' };
}

export function readAnchor(bodyPr: XmlNode | undefined): VerticalAnchor {
  const a = attr(bodyPr, 'anchor');
  if (a === 'ctr') return 'ctr';
  if (a === 'b') return 'bottom';
  return 'top';
}

const TEXT_VERTS: ReadonlySet<string> = new Set<TextVert>([
  'vert',
  'vert270',
  'eaVert',
  'wordArtVert',
  'wordArtVertRtl',
  'mongolianVert',
]);

/** The `vert` attribute of a `<a:bodyPr>` or `<a:tcPr>`; undefined for horizontal text. */
export function readVert(node: XmlNode | undefined): TextVert | undefined {
  const v = attr(node, 'vert');
  return v && TEXT_VERTS.has(v) ? (v as TextVert) : undefined;
}

function insetPx(bodyPr: XmlNode | undefined, name: string, fallback: number): number {
  const v = attrNum(bodyPr, name);
  return v === undefined ? fallback : emuToPx(v);
}

function parseParagraph(
  p: XmlNode,
  chain: TextStyleChain,
  ctx: ColorContext,
  scope: ParseScope,
): Paragraph {
  const pPr = child(p, 'pPr');
  const level = attrNum(pPr, 'lvl') ?? 0;
  const para = chain.paraProps(level, pPr);

  const runs: TextRun[] = [];
  for (const node of p.children) {
    const tag = localName(node.name);
    if (tag === 'r') {
      const run = parseRun(node, level, chain, ctx, scope);
      if (run.text.length > 0) {
        runs.push(run);
        scope.recordSource?.(run, node);
      }
    } else if (tag === 'br') {
      const rPr = child(node, 'rPr');
      const props = chain.runProps(level, rPr);
      const run: TextRun = { ...toRun(props), text: '\n' };
      runs.push(run);
      scope.recordSource?.(run, node);
    } else if (tag === 'fld') {
      // Fields (slide number, date...). Render the cached text it carries,
      // except a slide number, which PowerPoint recomputes for every slide.
      const text =
        attr(node, 'type') === 'slidenum' && scope.slideNumber !== undefined
          ? String(scope.slideNumber)
          : child(node, 't')?.text ?? '';
      if (text) {
        const parsed = parseRun(node, level, chain, ctx, scope);
        const run: TextRun = { ...parsed, text };
        runs.push(run);
        scope.recordSource?.(run, node);
      }
    }
  }

  // Sets the paragraph block's line-box "strut" (a floor on every line's
  // height). Each line's real height is set per-run by the unitless line-height,
  // so a tall run (e.g. a big number) still makes its own line tall; the strut
  // only needs to be small enough not to inflate the *other* lines. Use the
  // smallest run so a paragraph mixing a large run with small wrapping text
  // (e.g. "216" + a caption) doesn't balloon every wrapped line to the big size.
  // For an empty paragraph fall back to the paragraph mark (endParaRPr), which
  // gives a blank line its height.
  const sized = runs.filter((r) => r.sizePt !== undefined);
  const strutRun = sized.length
    ? sized.reduce((a, b) => (b.sizePt! < a.sizePt! ? b : a))
    : chain.runProps(level, child(p, 'endParaRPr'));
  const defaultSizePt = strutRun.sizePt;
  // The strut's face: the strut run's own, else any run's, else the mark's.
  const defaultFont =
    strutRun.font ?? runs.find((r) => r.font)?.font ?? chain.runProps(level, child(p, 'endParaRPr')).font;

  const result: Paragraph = { runs, level };
  if (defaultSizePt !== undefined && defaultSizePt > 0) result.defaultSizePt = defaultSizePt;
  if (defaultFont) result.defaultFont = defaultFont;
  if (para.align) result.align = para.align;
  if (para.bullet) result.bullet = para.bullet;
  if (para.marginLeftPx !== undefined) result.marginLeftPx = para.marginLeftPx;
  if (para.indentPx !== undefined) result.indentPx = para.indentPx;
  if (para.lineSpacingPct !== undefined) result.lineSpacingPct = para.lineSpacingPct;
  if (para.lineSpacingPt !== undefined) result.lineSpacingPt = para.lineSpacingPt;
  if (para.spaceBeforePt !== undefined) result.spaceBeforePt = para.spaceBeforePt;
  if (para.spaceAfterPt !== undefined) result.spaceAfterPt = para.spaceAfterPt;
  scope.recordSource?.(result, p);
  return result;
}

function parseRun(
  r: XmlNode,
  level: number,
  chain: TextStyleChain,
  ctx: ColorContext,
  scope: ParseScope,
): TextRun {
  const rPr = child(r, 'rPr');
  const props = chain.runProps(level, rPr);
  const run = toRun(props);
  run.text = child(r, 't')?.text ?? '';

  const hlink = child(rPr, 'hlinkClick');
  const rId = attr(hlink, 'id') ?? attr(hlink, 'r:id');
  if (rId && scope.resolveHyperlink) {
    const url = scope.resolveHyperlink(rId);
    if (url) run.hyperlink = url;
  }
  return run;
}

function toRun(props: ReturnType<TextStyleChain['runProps']>): TextRun {
  const run: TextRun = { text: '' };
  if (props.bold !== undefined) run.bold = props.bold;
  if (props.italic !== undefined) run.italic = props.italic;
  if (props.underline !== undefined) run.underline = props.underline;
  if (props.strike !== undefined) run.strike = props.strike;
  if (props.sizePt !== undefined) run.sizePt = props.sizePt;
  if (props.color !== undefined) run.color = props.color;
  if (props.font !== undefined) run.font = props.font;
  if (props.baseline !== undefined) run.baseline = props.baseline;
  if (props.highlight !== undefined) run.highlight = props.highlight;
  if (props.letterSpacingPt !== undefined) run.letterSpacingPt = props.letterSpacingPt;
  if (props.caps !== undefined) run.caps = props.caps;
  return run;
}
