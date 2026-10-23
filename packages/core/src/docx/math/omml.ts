/**
 * Office Math (OMML, the `m:` namespace) -> {@link MathNode}.
 *
 * OMML describes an equation structurally — a fraction has a numerator and a
 * denominator, a radical a degree and a body — which is also how MathML works,
 * so the conversion is a walk over the tree. The renderer then emits MathML,
 * which browsers lay out natively, so no math typesetting happens here.
 *
 * Text inside `<m:t>` is tokenised the way Word displays it: letters as
 * italic identifiers (unless the run is "plain", `<m:sty m:val="p"/>` or
 * `<m:nor/>`), digit runs as numbers, everything else as operators.
 */
import { child, children, attr, localName, type XmlNode } from '../../oxml/xml.js';
import type { MathNode } from '../model.js';

const row = (c: MathNode[]): MathNode => (c.length === 1 ? c[0]! : { t: 'row', c });

/** Convert an `<m:oMath>` (or any OMML container) to a math tree. */
export function ommlToMath(node: XmlNode): MathNode {
  return row(convertChildren(node));
}

function convertChildren(node: XmlNode | undefined): MathNode[] {
  if (!node) return [];
  const out: MathNode[] = [];
  for (const c of node.children) {
    const m = convert(c);
    if (m) out.push(m);
  }
  return out;
}

/** The `<m:e>`-style argument of a structure, as one node. */
const arg = (node: XmlNode | undefined, name: string): MathNode => row(convertChildren(child(node, name)));

/** `<m:xxxPr><m:yyy m:val>` — a property value off a structure's property element. */
function prop(node: XmlNode, pr: string, name: string): string | undefined {
  return attr(child(child(node, pr), name), 'val');
}

function convert(node: XmlNode): MathNode | undefined {
  switch (localName(node.name)) {
    case 'r':
      return run(node);
    case 'oMath':
    case 'e':
      return row(convertChildren(node));
    case 'f': {
      const type = prop(node, 'fPr', 'type');
      const frac: MathNode = { t: 'frac', num: arg(node, 'num'), den: arg(node, 'den') };
      if (type === 'noBar') frac.noBar = true;
      if (type === 'lin' || type === 'skw') frac.linear = true;
      return frac;
    }
    case 'sSup':
      return { t: 'script', base: arg(node, 'e'), sup: arg(node, 'sup') };
    case 'sSub':
      return { t: 'script', base: arg(node, 'e'), sub: arg(node, 'sub') };
    case 'sSubSup':
      return { t: 'script', base: arg(node, 'e'), sub: arg(node, 'sub'), sup: arg(node, 'sup') };
    case 'sPre':
      return { t: 'script', base: arg(node, 'e'), sub: arg(node, 'sub'), sup: arg(node, 'sup'), pre: true };
    case 'rad': {
      const hide = prop(node, 'radPr', 'degHide');
      const deg = child(node, 'deg');
      const root: MathNode = { t: 'root', e: arg(node, 'e') };
      if (deg && deg.children.length && hide !== '1' && hide !== 'on' && hide !== 'true') root.deg = arg(node, 'deg');
      return root;
    }
    case 'd': {
      // A delimiter: an open/close pair around one or more separated items.
      const pr = child(node, 'dPr');
      const open = attr(child(pr, 'begChr'), 'val') ?? '(';
      const close = attr(child(pr, 'endChr'), 'val') ?? ')';
      const sep = attr(child(pr, 'sepChr'), 'val') ?? '|';
      return { t: 'fence', open, close, sep, items: children(node, 'e').map((e) => row(convertChildren(e))) };
    }
    case 'nary': {
      const pr = child(node, 'naryPr');
      const op = attr(child(pr, 'chr'), 'val') ?? '∫';
      const loc = attr(child(pr, 'limLoc'), 'val');
      const hidden = (name: string) => {
        const v = attr(child(pr, name), 'val');
        return child(pr, name) !== undefined && v !== '0' && v !== 'off' && v !== 'false';
      };
      const n: MathNode = { t: 'nary', op, e: arg(node, 'e'), limits: loc ? loc === 'undOvr' : op !== '∫' && op !== '∬' && op !== '∭' && op !== '∮' };
      if (!hidden('subHide') && child(node, 'sub')?.children.length) n.sub = arg(node, 'sub');
      if (!hidden('supHide') && child(node, 'sup')?.children.length) n.sup = arg(node, 'sup');
      return n;
    }
    case 'func':
      // sin x: the name, then the argument (MathML spaces them via the apply op).
      return row([arg(node, 'fName'), { t: 'tok', kind: 'mo', v: '⁡' }, arg(node, 'e')]);
    case 'limLow':
      return { t: 'under', base: arg(node, 'e'), under: arg(node, 'lim') };
    case 'limUpp':
      return { t: 'over', base: arg(node, 'e'), over: arg(node, 'lim') };
    case 'acc': {
      const chr = prop(node, 'accPr', 'chr') ?? '̂';
      return { t: 'over', base: arg(node, 'e'), over: { t: 'tok', kind: 'mo', v: chr }, accent: true };
    }
    case 'bar': {
      const pos = prop(node, 'barPr', 'pos');
      const bar: MathNode = { t: 'tok', kind: 'mo', v: '¯' };
      return pos === 'bot'
        ? { t: 'under', base: arg(node, 'e'), under: { t: 'tok', kind: 'mo', v: '_' } }
        : { t: 'over', base: arg(node, 'e'), over: bar, accent: true };
    }
    case 'groupChr': {
      const chr = prop(node, 'groupChrPr', 'chr') ?? '⏟';
      const pos = prop(node, 'groupChrPr', 'pos') ?? 'bot';
      const mark: MathNode = { t: 'tok', kind: 'mo', v: chr };
      return pos === 'top' ? { t: 'over', base: arg(node, 'e'), over: mark } : { t: 'under', base: arg(node, 'e'), under: mark };
    }
    case 'm':
      return {
        t: 'table',
        rows: children(node, 'mr').map((mr) => children(mr, 'e').map((e) => row(convertChildren(e)))),
      };
    case 'eqArr':
      return { t: 'table', rows: children(node, 'e').map((e) => [row(convertChildren(e))]), align: 'left' };
    case 'box':
    case 'borderBox':
      return { t: 'box', e: arg(node, 'e') };
    case 'phant':
      return undefined;
    default:
      // Property elements (*Pr) and anything unknown contribute no content;
      // an unknown container still shows whatever text it holds.
      if (localName(node.name).endsWith('Pr')) return undefined;
      return node.children.length ? row(convertChildren(node)) : undefined;
  }
}

/** An `<m:r>`: its text, tokenised into identifiers, numbers and operators. */
function run(r: XmlNode): MathNode | undefined {
  const text = children(r, 't')
    .map((t) => t.text ?? '')
    .join('');
  if (!text) return undefined;
  const rPr = child(r, 'rPr');
  const sty = attr(child(rPr, 'sty'), 'val');
  // <m:nor/> is "normal text": shown as prose, not as math.
  if (child(rPr, 'nor')) return { t: 'tok', kind: 'mtext', v: text };
  const plain = sty === 'p' || sty === 'b';
  const out: MathNode[] = [];
  const re = /(\d+(?:[.,]\d+)*)|(\p{L})|(\s+)|([^\d\p{L}\s])/gu;
  for (const m of text.matchAll(re)) {
    if (m[1]) out.push({ t: 'tok', kind: 'mn', v: m[1] });
    else if (m[2]) out.push({ t: 'tok', kind: 'mi', v: m[2], ...(plain ? { normal: true } : {}) });
    else if (m[3]) out.push({ t: 'tok', kind: 'mtext', v: ' ' });
    else if (m[4]) out.push({ t: 'tok', kind: 'mo', v: m[4] });
  }
  // A run of plain letters (a function name like "sin") reads as one word.
  if (plain && out.every((n) => n.t === 'tok' && n.kind === 'mi')) return { t: 'tok', kind: 'mi', v: text, normal: true };
  return row(out);
}
