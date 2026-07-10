/**
 * Run parsing: <w:r> (and hyperlink-wrapped runs) -> DocxRun[].
 *
 * A single <w:r> can contain multiple text/break/tab children in document
 * order; we emit one DocxRun per <w:t>, carrying a break/tab flag from any
 * preceding <w:br>/<w:tab> so the renderer reproduces intra-run layout.
 */
import { children, child, attr, localName, type XmlNode } from '../../oxml/xml.js';
import { halfPtToPt } from '../units.js';
import { rPrFrom, mergeRun, type RunProps } from '../styles/styles.js';
import { parseRunDrawing } from '../images/image.js';
import type { DocxRun, DocxAnchor } from '../model.js';
import type { ParseContext } from '../document/context.js';

export interface FieldState {
  inField: boolean;
  fieldInstr: string;
  inSeparate: boolean;
}

/** Parse the runs inside a paragraph child (<w:r> or <w:hyperlink>). */
export function parseRunContainer(
  node: XmlNode,
  baseRun: RunProps,
  ctx: ParseContext,
  fieldState?: FieldState,
  hyperlink?: string,
  anchors?: DocxAnchor[],
): DocxRun[] {
  const name = localName(node.name);
  if (name === 'hyperlink') {
    const rel = ctx.rel(attr(node, 'id'));
    const href = rel?.mode === 'External' ? rel.target : undefined;
    const anchor = attr(node, 'anchor');
    const target = href ?? (anchor ? `#${anchor}` : undefined);
    const out: DocxRun[] = [];
    for (const r of children(node, 'r')) out.push(...parseRun(r, baseRun, ctx, fieldState, target, anchors));
    return out;
  }
  if (name === 'r') return parseRun(node, baseRun, ctx, fieldState, hyperlink, anchors);
  return [];
}

function parseRun(
  r: XmlNode,
  baseRun: RunProps,
  ctx: ParseContext,
  fieldState?: FieldState,
  hyperlink?: string,
  anchors?: DocxAnchor[],
): DocxRun[] {
  const props = mergeRun(baseRun, resolveRunStyle(r, ctx));
  const out: DocxRun[] = [];
  let pendingBreak = false;
  let pendingTab = false;

  for (const node of r.children) {
    const tagName = localName(node.name);
    switch (tagName) {
      case 'fldChar': {
        const type = attr(node, 'fldCharType');
        if (type === 'begin') {
          if (fieldState) {
            fieldState.inField = true;
            fieldState.fieldInstr = '';
            fieldState.inSeparate = false;
          }
        } else if (type === 'separate') {
          if (fieldState) {
            fieldState.inSeparate = true;
          }
        } else if (type === 'end') {
          if (fieldState) {
            fieldState.inField = false;
            fieldState.inSeparate = false;
            fieldState.fieldInstr = '';
          }
        }
        break;
      }
      case 'instrText': {
        if (fieldState && fieldState.inField) {
          fieldState.fieldInstr += node.text ?? '';
        }
        break;
      }
      case 'br':
        pendingBreak = true;
        break;
      case 'tab':
        pendingTab = true;
        break;
      case 't': {
        const run = makeRun(node.text ?? '', props, hyperlink, pendingBreak, pendingTab);
        if (fieldState && fieldState.inField && fieldState.inSeparate) {
          run.fieldCode = fieldState.fieldInstr;
        }
        out.push(run);
        pendingBreak = false;
        pendingTab = false;
        break;
      }
      case 'cr':
        pendingBreak = true;
        break;
      case 'drawing':
      case 'pict':
      case 'AlternateContent': {
        const res = parseRunDrawing(node, ctx);
        if (res?.inline) {
          const run = makeRun('', props, hyperlink, pendingBreak, pendingTab);
          run.drawing = res.inline;
          out.push(run);
          pendingBreak = false;
          pendingTab = false;
        } else if (res?.anchor) {
          anchors?.push(res.anchor);
        }
        break;
      }
      default:
        break;
    }
  }

  // A run that is only a <w:br>/<w:tab> with no text still needs to emit the break.
  if (out.length === 0 && (pendingBreak || pendingTab)) {
    out.push(makeRun('', props, hyperlink, pendingBreak, pendingTab));
  }
  return out;
}

function resolveRunStyle(r: XmlNode, ctx: ParseContext): RunProps {
  const rPr = child(r, 'rPr');
  const charStyleId = attr(child(rPr, 'rStyle'), 'val');
  const fromStyle = charStyleId ? ctx.styles.resolveCharStyle(charStyleId) : {};
  return mergeRun(fromStyle, rPrFrom(rPr));
}

function makeRun(
  text: string,
  props: RunProps,
  hyperlink: string | undefined,
  breakBefore: boolean,
  tabBefore: boolean,
): DocxRun {
  const run: DocxRun = { text };
  if (props.bold) run.bold = true;
  if (props.italic) run.italic = true;
  if (props.underline) run.underline = true;
  if (props.strike) run.strike = true;
  if (props.sizeHalfPt !== undefined) run.sizePt = halfPtToPt(props.sizeHalfPt);
  if (props.colorHex) run.colorHex = props.colorHex;
  if (props.highlightHex) run.highlightHex = props.highlightHex;
  if (props.font) run.font = props.font;
  if (props.vertAlign) run.vertAlign = props.vertAlign;
  if (props.caps) run.caps = props.caps;
  if (hyperlink) run.hyperlink = hyperlink;
  if (breakBefore) run.breakBefore = true;
  if (tabBefore) run.tabBefore = true;
  return run;
}
