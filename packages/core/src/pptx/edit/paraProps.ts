/**
 * DrawingML encoding of a {@link ParaFormat}: how a paragraph format maps to
 * `<a:pPr>`.
 *
 * The counterpart of `./format.ts`, which does the same for `<a:rPr>`. Three
 * things make paragraph properties harder than character ones:
 *
 *   1. **The element usually does not exist.** Most paragraphs state no
 *      properties at all and inherit everything from the layout, master
 *      `<p:txStyles>` and the presentation's default text style. Writing a
 *      bullet therefore means creating `<a:pPr>` — and it must be the *first*
 *      child of `<a:p>`, since the content model is `pPr?, runs*, endParaRPr?`.
 *   2. **The bullet is a choice group, not a property.** `<a:buNone>`,
 *      `<a:buAutoNum>` and `<a:buChar>` are mutually exclusive, as are the
 *      `buFont`/`buClr`/`buSz` families around them, so every write clears a
 *      whole group rather than replacing one sibling.
 *   3. **Turning a bullet off is an override, not a deletion.** A body
 *      placeholder inherits its bullet from the master, so deleting `<a:buChar>`
 *      would simply re-inherit it and the button would appear not to work.
 *      "Off" has to be *stated* as `<a:buNone>` — which is exactly what
 *      PowerPoint writes, at the cost of severing inheritance for that one
 *      property on paragraphs the user has touched.
 */
import {
  attr,
  attrNum,
  child,
  createElement,
  insertInOrder,
  localName,
  removeAttr,
  setAttr,
  type XmlNode,
} from '../../oxml/xml.js';
import type { ParaFormat } from '../../oxml/edit/format.js';
import type { Paragraph, TextAlign } from '../model.js';

/**
 * Child order of `<a:pPr>` (CT_TextParagraphProperties). Each `bu*` family is a
 * choice of at most one, and the sequence is enforced — a child in the wrong
 * slot makes PowerPoint call the part corrupt.
 */
const PPR_ORDER = [
  'lnSpc',
  'spcBef',
  'spcAft',
  'buClrTx',
  'buClr',
  'buSzTx',
  'buSzPct',
  'buSzPts',
  'buFontTx',
  'buFont',
  'buNone',
  'buAutoNum',
  'buChar',
  'tabLst',
  'defRPr',
  'extLst',
];

/** The three ways a bullet can be stated; at most one may be present. */
const BULLET_TAGS = ['buNone', 'buAutoNum', 'buChar'];
/** The typeface family that goes with a bullet, likewise exclusive. */
const BULLET_FONT_TAGS = ['buFontTx', 'buFont'];

/** Deepest level PowerPoint offers (nine levels, 0-based). */
export const MAX_LEVEL = 8;

/**
 * One level's worth of indent, in EMU (0.375"), matching the step PowerPoint's
 * own body placeholders use. Only ever written when nothing is inherited.
 */
const INDENT_STEP = 342900;

/** The default bullet glyph and the typeface it must be drawn in. */
const BULLET_CHAR = '•';
const BULLET_FONT = 'Arial';

const ALGN_TO_XML: Record<NonNullable<ParaFormat['align']>, TextAlign> = {
  left: 'l',
  center: 'ctr',
  right: 'r',
  justify: 'just',
};
const ALGN_FROM_XML: Record<TextAlign, NonNullable<ParaFormat['align']>> = {
  l: 'left',
  ctr: 'center',
  r: 'right',
  just: 'justify',
};

/** What a paragraph resolved to before the edit — its inherited geometry. */
export interface ResolvedPara {
  /** Level in effect, from the paragraph's own `lvl` or 0. */
  level: number;
  /** Resolved `indent` in px; negative means a hanging indent is inherited. */
  indentPx?: number;
  /** Resolved `marL` in px. */
  marginLeftPx?: number;
}

/** The `<a:pPr>` of a paragraph node, created in its required first slot. */
export function ensureParaProps(pNode: XmlNode, prefix: string): XmlNode {
  const existing = child(pNode, 'pPr');
  if (existing) return existing;
  const created = createElement(prefix ? `${prefix}:pPr` : 'pPr');
  pNode.children.unshift(created);
  return created;
}

/** Remove every child in a choice group, so a new one can take its place. */
function clearGroup(pPr: XmlNode, tags: string[]): void {
  pPr.children = pPr.children.filter((c) => !tags.includes(localName(c.name)));
}

/**
 * Apply a {@link ParaFormat} to an `<a:pPr>` in place.
 *
 * `resolved` describes what the paragraph inherited before the edit, which is
 * what decides whether we may leave the indent alone: a placeholder's per-level
 * `marL`/`indent` from the master are better than anything we can invent, while
 * a plain text box inherits nothing and needs a hanging indent written or the
 * bullet renders inline instead of in a gutter.
 */
export function applyParaFormat(
  pPr: XmlNode,
  format: ParaFormat,
  prefix: string,
  resolved?: ResolvedPara,
): void {
  const q = (n: string) => (prefix ? `${prefix}:${n}` : n);

  if (format.level !== undefined) {
    const from = resolved?.level ?? attrNum(pPr, 'lvl') ?? 0;
    const to = clampLevel(format.level);
    // PowerPoint omits lvl="0", so match it rather than writing the default.
    if (to === 0) removeAttr(pPr, 'lvl');
    else setAttr(pPr, 'lvl', String(to));
    // An explicit marL pins the indent no matter what level says, so it has to
    // travel with the level. Where the paragraph states none, the new level's
    // own `lvlNpPr` supplies the geometry and we must not pre-empt it.
    const ownMarL = attrNum(pPr, 'marL');
    if (ownMarL !== undefined && to !== from) {
      setAttr(pPr, 'marL', String(Math.max(0, ownMarL + (to - from) * INDENT_STEP)));
    }
  }

  if (format.list !== undefined) {
    clearGroup(pPr, BULLET_TAGS);
    clearGroup(pPr, BULLET_FONT_TAGS);

    if (format.list === 'none') {
      insertInOrder(pPr, createElement(q('buNone')), PPR_ORDER);
      // Without this an inherited negative indent leaves a bullet-sized gutter
      // in front of text that no longer has a bullet.
      setAttr(pPr, 'marL', '0');
      setAttr(pPr, 'indent', '0');
    } else {
      if (format.list === 'bullet') {
        // Not optional: a master that states buFont="Wingdings" would draw our
        // '•' as whatever Wingdings has at that code point.
        insertInOrder(
          pPr,
          createElement(q('buFont'), { typeface: BULLET_FONT, pitchFamily: '34', charset: '0' }),
          PPR_ORDER,
        );
        insertInOrder(pPr, createElement(q('buChar'), { char: BULLET_CHAR }), PPR_ORDER);
      } else {
        // Numbers are drawn in the paragraph's own typeface, so buFont stays cleared.
        insertInOrder(pPr, createElement(q('buAutoNum'), { type: 'arabicPeriod' }), PPR_ORDER);
      }
      ensureHangingIndent(pPr, format, resolved);
    }
  }

  if (format.align !== undefined) {
    // Always stated, never removed: the layout may align the placeholder itself,
    // and removing `algn` would inherit that back instead of meaning "left".
    setAttr(pPr, 'algn', ALGN_TO_XML[format.align]);
  }

  if (format.lineSpacingPct !== undefined) {
    clearGroup(pPr, ['lnSpc']);
    insertInOrder(
      pPr,
      createElement(q('lnSpc'), {}, [
        createElement(q('spcPct'), { val: String(Math.round(format.lineSpacingPct * 100000)) }),
      ]),
      PPR_ORDER,
    );
  }

  writeSpacing(pPr, 'spcBef', format.spaceBeforePt, q, PPR_ORDER);
  writeSpacing(pPr, 'spcAft', format.spaceAfterPt, q, PPR_ORDER);
}

/** `<a:spcBef>`/`<a:spcAft>` in hundredths of a point. */
function writeSpacing(
  pPr: XmlNode,
  tag: 'spcBef' | 'spcAft',
  pt: number | undefined,
  q: (n: string) => string,
  order: string[],
): void {
  if (pt === undefined) return;
  clearGroup(pPr, [tag]);
  insertInOrder(
    pPr,
    createElement(q(tag), {}, [createElement(q('spcPts'), { val: String(Math.round(pt * 100)) })]),
    order,
  );
}

/**
 * Give a new list a hanging indent, but only where one is not already inherited.
 *
 * The renderer draws the bullet in its own gutter when the resolved `indent` is
 * negative and inline otherwise, so a plain text box (which inherits nothing)
 * needs this written. A placeholder does not: the master states `marL`/`indent`
 * per level, and leaving those inherited is what makes a later level change
 * move the text correctly.
 */
function ensureHangingIndent(
  pPr: XmlNode,
  format: ParaFormat,
  resolved?: ResolvedPara,
): void {
  const inheritedHang = (resolved?.indentPx ?? 0) < 0;
  if (inheritedHang || attr(pPr, 'indent') !== undefined) return;
  const level = clampLevel(format.level ?? resolved?.level ?? 0);
  setAttr(pPr, 'marL', String(INDENT_STEP * (level + 1)));
  setAttr(pPr, 'indent', String(-INDENT_STEP));
}

export function clampLevel(level: number): number {
  return Math.max(0, Math.min(MAX_LEVEL, Math.round(level)));
}

/**
 * The paragraph format a model paragraph is *in*, inheritance included — what
 * the toolbar shows as active.
 *
 * A paragraph with no bullet anywhere in its chain resolves to `undefined`
 * rather than `{type:'none'}`; both mean "not a list" here, since the user
 * cannot tell them apart and a toggle must behave the same way for both.
 */
export function readParaFormat(para: Paragraph): ParaFormat {
  const bullet = para.bullet;
  const out: ParaFormat = {
    list:
      !bullet || bullet.type === 'none' ? 'none' : bullet.type === 'char' ? 'bullet' : 'number',
    level: para.level,
  };
  if (para.align) out.align = ALGN_FROM_XML[para.align];
  if (para.lineSpacingPct !== undefined) out.lineSpacingPct = para.lineSpacingPct;
  if (para.spaceBeforePt !== undefined) out.spaceBeforePt = para.spaceBeforePt;
  if (para.spaceAfterPt !== undefined) out.spaceAfterPt = para.spaceAfterPt;
  return out;
}

/** What a paragraph inherited, for {@link applyParaFormat}'s indent decisions. */
export function resolvedOf(para: Paragraph): ResolvedPara {
  const out: ResolvedPara = { level: para.level };
  if (para.indentPx !== undefined) out.indentPx = para.indentPx;
  if (para.marginLeftPx !== undefined) out.marginLeftPx = para.marginLeftPx;
  return out;
}
