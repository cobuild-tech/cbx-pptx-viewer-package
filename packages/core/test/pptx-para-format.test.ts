/**
 * The DrawingML half of paragraph formatting: what a {@link ParaFormat} does to
 * an `<a:pPr>`. Pure XML, no DOM — the DOM half (which paragraphs a click
 * targets, what the toolbar reports) lives in pptx-edit-dom.
 *
 * Two invariants are worth stating out loud, because PowerPoint rejects a file
 * that breaks either: the `bu*` families are exclusive choices, so a write
 * clears the whole group; and `<a:pPr>`'s children are a sequence, so they are
 * inserted by rank rather than appended.
 */
import { describe, it, expect } from 'vitest';
import { parseXml, serializeNode, child, attr, type XmlNode } from '../src/oxml/xml.js';
import {
  applyParaFormat,
  ensureParaProps,
  readParaFormat,
  clampLevel,
} from '../src/pptx/edit/paraProps.js';
import type { Paragraph } from '../src/pptx/model.js';

/** An `<a:pPr>` parsed from its own markup, ready to be rewritten. */
function pPr(inner = '', attrs = ''): XmlNode {
  return parseXml(`<a:pPr xmlns:a="a" ${attrs}>${inner}</a:pPr>`)!;
}

const names = (node: XmlNode) => node.children.map((c) => c.name);

describe('pptx paragraph format — the bullet choice group', () => {
  it('turns a numbered paragraph into a bulleted one, glyph typeface included', () => {
    const node = pPr('<a:buAutoNum type="arabicPeriod"/>');
    applyParaFormat(node, { list: 'bullet' }, 'a', { level: 0, indentPx: -48 });

    expect(names(node)).toEqual(['a:buFont', 'a:buChar']);
    expect(attr(child(node, 'buChar'), 'char')).toBe('•');
    // The typeface is not decoration: a master stating buFont="Wingdings" would
    // otherwise draw '•' as whatever Wingdings has at that code point.
    expect(attr(child(node, 'buFont'), 'typeface')).toBe('Arial');
  });

  it('turns a bulleted paragraph into a numbered one and drops the bullet font', () => {
    const node = pPr('<a:buFont typeface="Wingdings"/><a:buChar char="§"/>');
    applyParaFormat(node, { list: 'number' }, 'a', { level: 0, indentPx: -48 });

    expect(names(node)).toEqual(['a:buAutoNum']);
    expect(attr(child(node, 'buAutoNum'), 'type')).toBe('arabicPeriod');
    // Numbers are drawn in the paragraph's own typeface.
    expect(child(node, 'buFont')).toBeUndefined();
  });

  it('states buNone rather than deleting the bullet, and clears the gutter', () => {
    const node = pPr('<a:buFont typeface="Arial"/><a:buChar char="•"/>', 'marL="457200" indent="-457200"');
    applyParaFormat(node, { list: 'none' }, 'a', { level: 0, indentPx: -48 });

    // Deleting buChar would re-inherit the master's bullet; the "off" state has
    // to be stated to override it.
    expect(names(node)).toEqual(['a:buNone']);
    expect(attr(node, 'marL')).toBe('0');
    expect(attr(node, 'indent')).toBe('0');
  });
});

describe('pptx paragraph format — indent geometry', () => {
  it('writes a hanging indent where nothing is inherited', () => {
    // A plain text box inherits no list geometry, and the renderer only draws
    // the bullet in a gutter when the resolved indent is negative.
    const node = pPr();
    applyParaFormat(node, { list: 'bullet' }, 'a', { level: 0 });
    expect(attr(node, 'marL')).toBe('342900');
    expect(attr(node, 'indent')).toBe('-342900');
  });

  it('leaves a placeholder’s inherited indent alone', () => {
    const node = pPr();
    applyParaFormat(node, { list: 'bullet' }, 'a', { level: 1, indentPx: -36, marginLeftPx: 72 });
    expect(attr(node, 'marL')).toBeUndefined();
    expect(attr(node, 'indent')).toBeUndefined();
  });

  it('moves an explicit marL with the level, leaving indent as it is', () => {
    const node = pPr('<a:buChar char="•"/>', 'marL="342900" indent="-342900"');
    applyParaFormat(node, { level: 3 }, 'a', { level: 0 });

    expect(attr(node, 'lvl')).toBe('3');
    expect(attr(node, 'marL')).toBe(String(342900 * 4));
    expect(attr(node, 'indent')).toBe('-342900');
  });

  it('adds no marL to a paragraph that states none, so the level’s own geometry applies', () => {
    const node = pPr();
    applyParaFormat(node, { level: 3 }, 'a', { level: 0 });
    expect(attr(node, 'lvl')).toBe('3');
    expect(attr(node, 'marL')).toBeUndefined();
  });

  it('omits lvl="0" the way PowerPoint does, and clamps to nine levels', () => {
    const node = pPr('', 'lvl="4"');
    applyParaFormat(node, { level: 0 }, 'a', { level: 4 });
    expect(attr(node, 'lvl')).toBeUndefined();

    applyParaFormat(node, { level: 99 }, 'a', { level: 0 });
    expect(attr(node, 'lvl')).toBe('8');
    expect(clampLevel(-3)).toBe(0);
  });
});

describe('pptx paragraph format — alignment and spacing', () => {
  it('states alignment explicitly', () => {
    const node = pPr();
    applyParaFormat(node, { align: 'justify' }, 'a');
    expect(attr(node, 'algn')).toBe('just');
    applyParaFormat(node, { align: 'left' }, 'a');
    // Never removed: the layout may centre the placeholder, and dropping algn
    // would inherit that back instead of meaning "left".
    expect(attr(node, 'algn')).toBe('l');
  });

  it('replaces a point line spacing with a percentage rather than nesting a second child', () => {
    const node = pPr('<a:lnSpc><a:spcPts val="1800"/></a:lnSpc>');
    applyParaFormat(node, { lineSpacingPct: 1.5 }, 'a');

    expect(names(node)).toEqual(['a:lnSpc']);
    const lnSpc = child(node, 'lnSpc')!;
    expect(names(lnSpc)).toEqual(['a:spcPct']);
    expect(attr(child(lnSpc, 'spcPct'), 'val')).toBe('150000');
  });

  it('writes space before and after in hundredths of a point', () => {
    const node = pPr();
    applyParaFormat(node, { spaceBeforePt: 6, spaceAfterPt: 0 }, 'a');
    expect(attr(child(child(node, 'spcBef')!, 'spcPts'), 'val')).toBe('600');
    expect(attr(child(child(node, 'spcAft')!, 'spcPts'), 'val')).toBe('0');
  });
});

describe('pptx paragraph format — schema order', () => {
  it('inserts every child in its schema slot regardless of write order', () => {
    const node = pPr();
    // Deliberately the wrong order: spacing first, bullet last.
    applyParaFormat(
      node,
      { spaceBeforePt: 6, lineSpacingPct: 1.5, list: 'bullet', align: 'center' },
      'a',
      { level: 0, indentPx: -24 },
    );
    expect(names(node)).toEqual(['a:lnSpc', 'a:spcBef', 'a:buFont', 'a:buChar']);
  });

  it('creates the pPr as the first child of the paragraph, before the runs', () => {
    const p = parseXml(`<a:p xmlns:a="a"><a:r><a:t>hi</a:t></a:r></a:p>`)!;
    const props = ensureParaProps(p, 'a');
    applyParaFormat(props, { list: 'number' }, 'a', { level: 0, indentPx: -24 });

    // <a:p> is pPr?, runs*, endParaRPr? — a trailing pPr makes the part invalid.
    expect(names(p)).toEqual(['a:pPr', 'a:r']);
    expect(serializeNode(p)).toContain('<a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr>');
    // Asking again returns the same element rather than adding a second.
    expect(ensureParaProps(p, 'a')).toBe(props);
  });

  it('keeps properties it does not understand', () => {
    const node = pPr('<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:tabLst><a:tab pos="914400"/></a:tabLst>');
    applyParaFormat(node, { list: 'bullet' }, 'a', { level: 0, indentPx: -24 });
    expect(names(node)).toEqual(['a:lnSpc', 'a:buFont', 'a:buChar', 'a:tabLst']);
  });
});

describe('pptx paragraph format — reading the state back', () => {
  const para = (over: Partial<Paragraph>): Paragraph =>
    ({ runs: [], level: 0, ...over }) as Paragraph;

  it('reports the resolved list kind, treating "no bullet anywhere" as none', () => {
    expect(readParaFormat(para({})).list).toBe('none');
    expect(readParaFormat(para({ bullet: { type: 'none' } })).list).toBe('none');
    expect(readParaFormat(para({ bullet: { type: 'char', char: '•' } })).list).toBe('bullet');
    expect(
      readParaFormat(para({ bullet: { type: 'number', scheme: 'arabicPeriod' } })).list,
    ).toBe('number');
  });

  it('maps alignment and spacing into neutral names', () => {
    const f = readParaFormat(
      para({ level: 2, align: 'ctr', lineSpacingPct: 1.5, spaceBeforePt: 6, spaceAfterPt: 3 }),
    );
    expect(f).toEqual({
      list: 'none',
      level: 2,
      align: 'center',
      lineSpacingPct: 1.5,
      spaceBeforePt: 6,
      spaceAfterPt: 3,
    });
  });
});
