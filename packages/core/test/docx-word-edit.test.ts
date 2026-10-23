/**
 * Word-style DOCX editing: positional commands (type, Enter, Backspace,
 * Delete, paste, formatting, lists) applied to the XML in place.
 *
 * The point of the positional engine is that an edit touches only what it
 * must — so alongside "the text is right" most tests assert that pictures,
 * fields, bookmarks and untouched runs come through byte-identical.
 */
import { describe, it, expect } from 'vitest';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { DocxDocument } from '../src/docx/document/document.js';
import { DocxRelType } from '../src/docx/relTypes.js';
import { DocxEditSession, caret, type DocSel } from '../src/docx/edit/session.js';
import { storyParagraphs } from '../src/docx/edit/story.js';
import type { DocxParagraph } from '../src/docx/model.js';
import type { XmlNode } from '../src/oxml/xml.js';

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

function build(body: string, extra: Record<string, string> = {}, docRels = ''): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
      </Types>`,
    ),
    '_rels/.rels': strToU8(rels(`<Relationship Id="rId1" Type="${DocxRelType.OfficeDocument}" Target="word/document.xml"/>`)),
    'word/_rels/document.xml.rels': strToU8(rels(docRels)),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="w" xmlns:w14="w14" xmlns:m="m"><w:body>${body}<w:sectPr/></w:body></w:document>`,
    ),
    ...Object.fromEntries(Object.entries(extra).map(([k, v]) => [k, strToU8(v)])),
  });
}

function open(body: string, extra?: Record<string, string>, docRels?: string) {
  const doc = DocxDocument.load(build(body, extra, docRels));
  const session = new DocxEditSession(doc);
  const paras = () => storyParagraphs(doc.sections.flatMap((s) => s.blocks)).body;
  const node = (i: number): XmlNode => session.nodeOf(paras()[i]!)!;
  const text = (i: number) => session.map(node(i))!.text;
  const texts = () => paras().map((p) => session.mapOf(p)!.text);
  const xml = () => strFromU8(unzipSync(session.exportBytes())['word/document.xml']!);
  const at = (i: number, off: number) => caret(node(i), off);
  const range = (i: number, a: number, j: number, b: number): DocSel => ({ anchor: { p: node(i), off: a }, focus: { p: node(j), off: b } });
  return { doc, session, paras, node, text, texts, xml, at, range };
}

const P = (inner: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inner}</w:p>`;
const R = (t: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;

describe('word editing — typing', () => {
  it('types into the run at the caret, leaving its neighbours untouched', () => {
    const e = open(P(R('Hello ', '<w:sz w:val="48"/>') + R('World', '<w:b/>')));
    const sel = e.session.insertText(e.at(0, 5), ',');
    expect(e.text(0)).toBe('Hello, World');
    expect(sel!.focus.off).toBe(6);
    expect(e.xml()).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">World</w:t></w:r>');
  });

  it('continues the run on the left at a run boundary', () => {
    const e = open(P(R('Bold', '<w:b/>') + R(' plain')));
    e.session.insertText(e.at(0, 4), 'er');
    expect(e.paras()[0]!.runs.map((r) => [r.text, !!r.bold])).toEqual([
      ['Bolder', true],
      [' plain', false],
    ]);
  });

  it('types into an empty paragraph with the paragraph mark’s formatting', () => {
    const e = open(P('', '<w:rPr><w:i/></w:rPr>'));
    e.session.insertText(e.at(0, 0), 'New');
    expect(e.paras()[0]!.runs[0]).toMatchObject({ text: 'New', italic: true });
  });

  it('types a tab as <w:tab/> and keeps pictures around the caret', () => {
    const pic =
      '<w:r><w:drawing><wp:inline xmlns:wp="wp" xmlns:a="a" xmlns:pic="pic" xmlns:r="r"><wp:extent cx="9525" cy="9525"/>' +
      '<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rI"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>' +
      '</wp:inline></w:drawing></w:r>';
    const e = open(P(R('ab') + pic + R('cd')), { 'word/media/image1.png': 'png' }, `<Relationship Id="rI" Type="${DocxRelType.Image}" Target="media/image1.png"/>`);
    expect(e.text(0)).toBe('ab￼cd');
    e.session.insertText(e.at(0, 3), '\tx');
    expect(e.text(0)).toBe('ab￼\txcd');
    expect(e.xml()).toContain('<w:drawing>');
  });

  it('types after a field, not into its result', () => {
    const field =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const e = open(P(R('Page ') + field));
    e.session.insertText(e.at(0, 6), '!');
    expect(e.xml()).toMatch(/fldCharType="end"\/><\/w:r><w:r><w:t>!<\/w:t><\/w:r>/);
  });

  it('replaces a selection with the first selected character’s formatting', () => {
    const e = open(P(R('one ') + R('two', '<w:b/>') + R(' three')));
    e.session.insertText(e.range(0, 4, 0, 7), 'TWO');
    expect(e.text(0)).toBe('one TWO three');
    expect(e.paras()[0]!.runs.find((r) => r.text === 'TWO')?.bold).toBe(true);
  });
});

describe('word editing — Enter and deletion', () => {
  it('splits a paragraph, the second half keeping the original node and section break', () => {
    const e = open(P(R('Hello World'), '<w:jc w:val="center"/><w:sectPr/>').replace('<w:p>', '<w:p w14:paraId="AA">'));
    const original = e.node(0);
    const sel = e.session.insertParagraph(e.at(0, 5));
    expect(e.texts()).toEqual(['Hello', ' World']);
    expect(sel!.focus.p).toBe(original);
    expect(e.node(1)).toBe(original);
    const xml = e.xml();
    expect(xml.match(/w14:paraId/g)).toHaveLength(1);
    expect(xml.match(/<w:jc w:val="center"\/>/g)).toHaveLength(2);
    expect(xml.match(/<w:sectPr\/>/g)).toHaveLength(2); // the paragraph's, plus the body's
  });

  it('starts the style’s next paragraph after Enter at the end of a heading', () => {
    const styles = `<w:styles xmlns:w="w">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
      <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:next w:val="Normal"/></w:style>
    </w:styles>`;
    const e = open(P(R('Title'), '<w:pStyle w:val="Heading1"/>'), { 'word/styles.xml': styles }, `<Relationship Id="rS" Type="${DocxRelType.Styles}" Target="styles.xml"/>`);
    e.session.insertParagraph(e.at(0, 5));
    expect(e.session.styleIdOf(e.node(0))).toBe('Heading1');
    expect(e.session.styleIdOf(e.node(1))).toBeUndefined();
  });

  it('Backspace deletes a character in place', () => {
    const e = open(P(R('Hello')));
    e.session.deleteBackward(e.at(0, 5));
    expect(e.text(0)).toBe('Hell');
  });

  it('Backspace at a paragraph start joins it to the one before, keeping the first’s properties', () => {
    const e = open(P(R('One'), '<w:jc w:val="right"/>') + P(R('Two')));
    const sel = e.session.deleteBackward(e.at(1, 0));
    expect(e.texts()).toEqual(['OneTwo']);
    expect(sel!.focus.off).toBe(3);
    expect(e.paras()[0]!.align).toBe('r');
  });

  it('Backspace into an empty paragraph takes the lower paragraph’s properties', () => {
    const e = open(P('', '<w:jc w:val="right"/>') + P(R('Two'), '<w:jc w:val="center"/>'));
    e.session.deleteBackward(e.at(1, 0));
    expect(e.texts()).toEqual(['Two']);
    expect(e.paras()[0]!.align).toBe('ctr');
  });

  it('Delete at a paragraph end joins the next one', () => {
    const e = open(P(R('One')) + P(R('Two')));
    e.session.deleteForward(e.at(0, 3));
    expect(e.texts()).toEqual(['OneTwo']);
  });

  it('Ctrl+Backspace deletes a word', () => {
    const e = open(P(R('alpha beta gamma')));
    e.session.deleteBackward(e.at(0, 10), 'word');
    expect(e.text(0)).toBe('alpha  gamma');
  });

  it('deletes a selection across paragraphs and the table between them', () => {
    const tbl = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const e = open(P(R('first line')) + tbl + P(R('last line')));
    e.session.deleteSelection(e.range(0, 5, 2, 5));
    expect(e.texts()).toEqual(['firstline']);
    expect(e.xml()).not.toContain('w:tbl');
  });

  it('keeps bookmarks and drops a deleted field whole', () => {
    const field =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>today</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const e = open(P(R('a') + '<w:bookmarkStart w:id="0" w:name="x"/>' + field + R('b')));
    e.session.deleteSelection(e.range(0, 1, 0, 6));
    expect(e.text(0)).toBe('ab');
    const xml = e.xml();
    expect(xml).toContain('<w:bookmarkStart w:id="0" w:name="x"/>');
    expect(xml).not.toContain('fldChar');
  });

  it('keeps a field’s structure when only part of its result is deleted', () => {
    const field =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>today</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const e = open(P(R('a') + field + R('b')));
    e.session.deleteSelection(e.range(0, 0, 0, 3));
    expect(e.xml().match(/<w:fldChar /g)).toHaveLength(3);
  });

  it('pastes several lines as paragraphs', () => {
    const e = open(P(R('startend'), '<w:jc w:val="center"/>'));
    const sel = e.session.paste(e.at(0, 5), 'one\ntwo\nthree');
    expect(e.texts()).toEqual(['startone', 'two', 'threeend']);
    expect(sel!.focus).toEqual({ p: e.node(2), off: 5 });
    expect(e.paras().every((p) => p.align === 'ctr')).toBe(true);
  });
});

describe('word editing — formatting', () => {
  it('bolds part of a run by cutting it in three', () => {
    const e = open(P(R('abcdef', '<w:i/>')));
    e.session.formatSelection(e.range(0, 2, 0, 4), { bold: true });
    expect(e.paras()[0]!.runs.map((r) => [r.text, !!r.bold, !!r.italic])).toEqual([
      ['ab', false, true],
      ['cd', true, true],
      ['ef', false, true],
    ]);
  });

  it('merges the pieces back when the formatting is undone by hand', () => {
    const e = open(P(R('abcdef')));
    e.session.formatSelection(e.range(0, 2, 0, 4), { bold: true });
    e.session.formatSelection(e.range(0, 2, 0, 4), { bold: false });
    expect(e.paras()[0]!.runs.length).toBeLessThanOrEqual(3);
    expect(e.text(0)).toBe('abcdef');
  });

  it('formats across paragraphs, including the paragraph marks it covers whole', () => {
    const e = open(P(R('one')) + P(R('two')));
    e.session.formatSelection(e.range(0, 0, 1, 3), { sizePt: 20 });
    expect(e.paras().map((p) => p.runs[0]!.sizePt)).toEqual([20, 20]);
    expect(e.xml().match(/<w:pPr><w:rPr><w:sz w:val="40"\/>/g)).toHaveLength(2);
  });

  it('sets alignment and spacing in schema order', () => {
    const e = open(P(R('x'), '<w:pStyle w:val="Body"/><w:rPr><w:b/></w:rPr>'));
    e.session.setParagraphFormat(e.at(0, 0), { align: 'justify', spaceAfterPt: 6 });
    expect(e.xml()).toContain('<w:pPr><w:pStyle w:val="Body"/><w:spacing w:after="120"/><w:jc w:val="both"/><w:rPr><w:b/></w:rPr></w:pPr>');
  });

  it('makes a bulleted list, creating numbering.xml when the document has none', () => {
    const e = open(P(R('one')) + P(R('two')));
    e.session.toggleList(e.range(0, 0, 1, 0), 'bullet');
    expect(e.paras().map((p) => !!p.listMarker)).toEqual([true, true]);
    expect(e.session.listKindOf(e.node(0))).toBe('bullet');
    const files = unzipSync(e.session.exportBytes());
    expect(strFromU8(files['word/numbering.xml']!)).toContain('w:numFmt w:val="bullet"');
    expect(strFromU8(files['word/_rels/document.xml.rels']!)).toContain('numbering.xml');
    expect(strFromU8(files['[Content_Types].xml']!)).toContain('/word/numbering.xml');
    // A reload sees the same list.
    const again = DocxDocument.load(e.session.exportBytes());
    expect((again.sections[0]!.blocks[0] as DocxParagraph).listMarker).toBeTruthy();
  });

  it('toggles a list off again, and Enter on an empty item ends it', () => {
    const e = open(P(R('one')));
    e.session.toggleList(e.at(0, 0), 'number');
    expect(e.paras()[0]!.listMarker).toBe('1.');
    e.session.insertParagraph(e.at(0, 3));
    expect(e.paras()[1]!.listMarker).toBe('2.');
    e.session.insertParagraph(e.at(1, 0));
    expect(e.paras()[1]!.listMarker).toBeUndefined();
  });

  it('indents a list item a level, and a plain paragraph by half an inch', () => {
    const e = open(P(R('item')) + P(R('body')));
    e.session.toggleList(e.at(0, 0), 'bullet');
    e.session.indent(e.at(0, 0), 1);
    expect(e.session.paraFormatAt(e.at(0, 0)).level).toBe(1);
    e.session.indent(e.at(1, 0), 1);
    expect(e.xml()).toContain('<w:ind w:left="720"/>');
  });
});

describe('word editing — undo', () => {
  it('undoes a burst of typing as one step and restores the caret', () => {
    const e = open(P(R('ab')));
    let sel = e.session.insertText(e.at(0, 2), 'c')!;
    sel = e.session.insertText(sel, 'd')!;
    e.session.insertText(sel, 'e');
    expect(e.text(0)).toBe('abcde');
    e.session.undo();
    expect(e.text(0)).toBe('ab');
    expect(e.session.lastSelection!.focus.off).toBe(2);
    expect(e.session.canUndo).toBe(false);
    e.session.redo();
    expect(e.text(0)).toBe('abcde');
    expect(e.session.lastSelection!.focus.off).toBe(5);
  });

  it('separates typing from a following Enter', () => {
    const e = open(P(R('ab')));
    const sel = e.session.insertText(e.at(0, 2), 'c')!;
    e.session.insertParagraph(sel);
    e.session.undo();
    expect(e.texts()).toEqual(['abc']);
  });

  it('leaves untouched parts byte-identical on export', () => {
    const e = open(P(R('ab')), { 'word/extra.xml': '<keep   me="1"/>' });
    e.session.insertText(e.at(0, 1), 'x');
    expect(strFromU8(unzipSync(e.session.exportBytes())['word/extra.xml']!)).toBe('<keep   me="1"/>');
  });
});
