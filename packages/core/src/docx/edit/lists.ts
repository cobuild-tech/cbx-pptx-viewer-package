/**
 * Lists and paragraph styles, read and written the way Word's Home ribbon
 * does: a bullet or numbering button gives the paragraphs a `<w:numPr>`
 * pointing at a list definition in numbering.xml (creating one when the
 * document has none to reuse), and the style gallery sets `<w:pStyle>`.
 */
import { attr, child, children, localName, parseXml, type XmlNode } from '../../oxml/xml.js';

export const NUMBERING_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';

export const EMPTY_NUMBERING =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>';

function val(n: XmlNode | undefined): string | undefined {
  return attr(n, 'val');
}

/** `<w:abstractNum>` behind a `<w:num>`. */
function abstractFor(numbering: XmlNode, numId: string): XmlNode | undefined {
  const num = children(numbering, 'num').find((n) => attr(n, 'numId') === numId);
  const absId = val(child(num, 'abstractNumId'));
  return children(numbering, 'abstractNum').find((a) => attr(a, 'abstractNumId') === absId);
}

/** Whether a list level draws bullets or numbers. */
export function listKind(numbering: XmlNode | undefined, numId: string, ilvl: number): 'bullet' | 'number' | undefined {
  if (!numbering || numId === '0') return undefined;
  const num = children(numbering, 'num').find((n) => attr(n, 'numId') === numId);
  const override = children(num, 'lvlOverride').find((o) => attr(o, 'ilvl') === String(ilvl));
  const lvl =
    child(override, 'lvl') ??
    children(abstractFor(numbering, numId), 'lvl').find((l) => attr(l, 'ilvl') === String(ilvl));
  const fmt = val(child(lvl, 'numFmt'));
  if (!fmt || fmt === 'none') return undefined;
  return fmt === 'bullet' ? 'bullet' : 'number';
}

/** Bullet glyphs per level, as Word's default bulleted list draws them. */
const BULLETS: Array<[string, string]> = [
  ['', 'Symbol'],
  ['o', 'Courier New'],
  ['', 'Wingdings'],
];

/** Number formats per level, as Word's default numbered list cycles them. */
const NUMBERS = ['decimal', 'lowerLetter', 'lowerRoman'];

/**
 * Append a new nine-level list definition of `kind` and a `<w:num>` using it;
 * returns the new numId. Indents follow Word's defaults: 0.5" per level with a
 * 0.25" hanging marker.
 */
export function createList(numbering: XmlNode, kind: 'bullet' | 'number', prefix: string): string {
  const w = prefix ? `${prefix}:` : '';
  const ids = (name: string, a: string) =>
    children(numbering, name).map((n) => Number(attr(n, a))).filter((n) => Number.isFinite(n));
  const absId = String(Math.max(-1, ...ids('abstractNum', 'abstractNumId')) + 1);
  const numId = String(Math.max(0, ...ids('num', 'numId')) + 1);
  let levels = '';
  for (let i = 0; i < 9; i++) {
    const left = 720 * (i + 1);
    if (kind === 'bullet') {
      const [glyph, font] = BULLETS[i % 3]!;
      levels +=
        `<${w}lvl ${w}ilvl="${i}"><${w}start ${w}val="1"/><${w}numFmt ${w}val="bullet"/>` +
        `<${w}lvlText ${w}val="${glyph}"/><${w}lvlJc ${w}val="left"/>` +
        `<${w}pPr><${w}ind ${w}left="${left}" ${w}hanging="360"/></${w}pPr>` +
        `<${w}rPr><${w}rFonts ${w}ascii="${font}" ${w}hAnsi="${font}" ${w}hint="default"/></${w}rPr></${w}lvl>`;
    } else {
      const fmt = NUMBERS[i % 3]!;
      levels +=
        `<${w}lvl ${w}ilvl="${i}"><${w}start ${w}val="1"/><${w}numFmt ${w}val="${fmt}"/>` +
        `<${w}lvlText ${w}val="%${i + 1}."/><${w}lvlJc ${w}val="${fmt === 'lowerRoman' ? 'right' : 'left'}"/>` +
        `<${w}pPr><${w}ind ${w}left="${left}" ${w}hanging="${fmt === 'lowerRoman' ? 180 : 360}"/></${w}pPr></${w}lvl>`;
    }
  }
  const ns = prefix ? ` xmlns:${prefix}="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` : '';
  const frag = parseXml(
    `<${w}frag${ns}><${w}abstractNum ${w}abstractNumId="${absId}"><${w}multiLevelType ${w}val="hybridMultilevel"/>${levels}</${w}abstractNum>` +
      `<${w}num ${w}numId="${numId}"><${w}abstractNumId ${w}val="${absId}"/></${w}num></${w}frag>`,
  );
  const [abs, num] = frag!.children;
  // Schema order: every <w:abstractNum> before the first <w:num>.
  const firstNum = numbering.children.findIndex((c) => ['num', 'numIdMacAtCleanup'].includes(localName(c.name)));
  numbering.children.splice(firstNum === -1 ? numbering.children.length : firstNum, 0, abs!);
  const cleanup = numbering.children.findIndex((c) => localName(c.name) === 'numIdMacAtCleanup');
  numbering.children.splice(cleanup === -1 ? numbering.children.length : cleanup, 0, num!);
  return numId;
}

/** A paragraph style as the style gallery lists it. */
export interface ParaStyleInfo {
  id: string;
  name: string;
  /** Style for the paragraph Enter makes at the end of one in this style. */
  next?: string;
  /** Shown in Word's style gallery (`<w:qFormat/>`). */
  quick: boolean;
}

/** Paragraph styles in styles.xml, and the default one. */
export function readParaStyles(styles: XmlNode | undefined): { list: ParaStyleInfo[]; defaultId?: string } {
  const list: ParaStyleInfo[] = [];
  let defaultId: string | undefined;
  for (const s of children(styles, 'style')) {
    if (attr(s, 'type') !== 'paragraph') continue;
    const id = attr(s, 'styleId');
    if (!id) continue;
    const def = attr(s, 'default');
    if (def === '1' || def === 'true') defaultId = id;
    const info: ParaStyleInfo = {
      id,
      name: val(child(s, 'name')) ?? id,
      quick: !!child(s, 'qFormat') && !child(s, 'semiHidden'),
    };
    const next = val(child(s, 'next'));
    if (next) info.next = next;
    list.push(info);
  }
  return { list, ...(defaultId ? { defaultId } : {}) };
}

/** Word's names for built-in styles, as its gallery shows them. */
export function displayStyleName(name: string): string {
  const lower = name.toLowerCase();
  if (lower === 'normal') return 'Normal';
  const h = /^heading (\d)$/.exec(lower);
  if (h) return `Heading ${h[1]}`;
  if (lower === 'title') return 'Title';
  if (lower === 'subtitle') return 'Subtitle';
  return name;
}
