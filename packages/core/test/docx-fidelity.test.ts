/**
 * Layout fidelity of the DOCX parse: the Word rules that decide where text
 * sits and how big it is — theme fonts, the style/table-style/numbering
 * cascade, tab stops, keep and widow rules, footnotes, section inheritance.
 * Each case mirrors a behaviour checked against Word itself.
 */
import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { DocxDocument } from '../src/docx/document/document.js';
import { DocxRelType } from '../src/docx/relTypes.js';
import type { DocxBlock, DocxParagraph, DocxTable } from '../src/docx/model.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

interface Parts {
  body: string;
  styles?: string;
  numbering?: string;
  theme?: string;
  settings?: string;
  footnotes?: string;
  /** Extra parts (e.g. headers) and the relationships pointing at them. */
  extra?: Record<string, string>;
  extraRels?: string;
}

/** A minimal package around the given parts. */
function load(parts: Parts): DocxDocument {
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
      </Types>`,
    ),
    '_rels/.rels': strToU8(rels(`<Relationship Id="rId1" Type="${DocxRelType.OfficeDocument}" Target="word/document.xml"/>`)),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:a="${A}"><w:body>${parts.body}</w:body></w:document>`,
    ),
  };
  let docRels = parts.extraRels ?? '';
  const add = (name: string, xml: string | undefined, type: string, id: string) => {
    if (!xml) return;
    files[`word/${name}`] = strToU8(xml);
    docRels += `<Relationship Id="${id}" Type="${type}" Target="${name}"/>`;
  };
  add('styles.xml', parts.styles && `<w:styles xmlns:w="${W}">${parts.styles}</w:styles>`, DocxRelType.Styles, 'rSt');
  add('numbering.xml', parts.numbering && `<w:numbering xmlns:w="${W}">${parts.numbering}</w:numbering>`, DocxRelType.Numbering, 'rNum');
  add('theme/theme1.xml', parts.theme, DocxRelType.Theme, 'rTheme');
  add('settings.xml', parts.settings && `<w:settings xmlns:w="${W}">${parts.settings}</w:settings>`, DocxRelType.Settings, 'rSet');
  add('footnotes.xml', parts.footnotes && `<w:footnotes xmlns:w="${W}">${parts.footnotes}</w:footnotes>`, DocxRelType.Footnotes, 'rFn');
  for (const [name, xml] of Object.entries(parts.extra ?? {})) files[`word/${name}`] = strToU8(xml);
  files['word/_rels/document.xml.rels'] = strToU8(rels(docRels));
  return DocxDocument.load(zipSync(files));
}

const blocks = (doc: DocxDocument, section = 0): DocxBlock[] => doc.sections[section]!.blocks;
const paras = (doc: DocxDocument): DocxParagraph[] =>
  blocks(doc).filter((b): b is DocxParagraph => b.kind === 'paragraph');
const p = (text: string, pPr = '', rPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const THEME = `<a:theme xmlns:a="${A}"><a:themeElements>
  <a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    <a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme>
  <a:fontScheme name="T"><a:majorFont><a:latin typeface="Cambria"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
</a:themeElements></a:theme>`;

describe('fonts', () => {
  it('resolves theme font references, which win over a literal face in the same rFonts', () => {
    const doc = load({
      theme: THEME,
      body:
        p('body', '', '<w:rFonts w:ascii="Arial" w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>') +
        p('head', '', '<w:rFonts w:asciiTheme="majorHAnsi"/>'),
    });
    const [a, b] = paras(doc);
    expect(a!.runs[0]!.font).toBe('Calibri');
    expect(b!.runs[0]!.font).toBe('Cambria');
  });

  it("falls back to Word's own default — Times New Roman 10pt — when nothing is stated", () => {
    const run = paras(load({ body: p('x') }))[0]!.runs[0]!;
    expect(run.font).toBe('Times New Roman');
    expect(run.sizePt).toBe(10);
  });

  it('formats only the mark with the paragraph mark rPr, which sizes an empty paragraph', () => {
    const doc = load({
      body: p('text', '<w:rPr><w:sz w:val="48"/></w:rPr>') + '<w:p><w:pPr><w:rPr><w:sz w:val="4"/></w:rPr></w:pPr></w:p>',
    });
    const [text, empty] = paras(doc);
    expect(text!.runs[0]!.sizePt).toBe(10);
    expect(empty!.baseFontSizePt).toBe(2);
  });
});

describe('paragraph cascade', () => {
  it('applies a list level indent over the style indent, with direct formatting over both', () => {
    const doc = load({
      styles: `<w:style w:type="paragraph" w:styleId="List"><w:pPr><w:ind w:left="720" w:firstLine="200"/></w:pPr></w:style>`,
      numbering: `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>
          <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
      body:
        p('level', '<w:pStyle w:val="List"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
        p('direct', '<w:pStyle w:val="List"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="567"/>'),
    });
    const [level, direct] = paras(doc);
    expect(level!.indentLeftPx).toBe(96); // 1440 twips, from the level
    expect(level!.indentFirstLinePx).toBe(-24); // the level's hanging displaced the style's firstLine
    expect(direct!.indentLeftPx).toBeCloseTo(37.8); // 567 twips, direct
    expect(direct!.indentFirstLinePx).toBe(-24);
  });

  it('accumulates tab stops down the cascade, a clear removing an inherited one', () => {
    const doc = load({
      styles: `<w:style w:type="paragraph" w:styleId="Tabs"><w:pPr><w:tabs>
        <w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360" w:leader="dot"/></w:tabs></w:pPr></w:style>`,
      settings: '<w:defaultTabStop w:val="360"/>',
      body: p('x', '<w:pStyle w:val="Tabs"/><w:tabs><w:tab w:val="clear" w:pos="4680"/><w:tab w:val="left" w:pos="1440"/></w:tabs>'),
    });
    const para = paras(doc)[0]!;
    expect(para.tabs).toEqual([
      { posPx: 96, align: 'left' },
      { posPx: 624, align: 'right', leader: 'dot' },
    ]);
    expect(para.defaultTabPx).toBe(24);
  });

  it('lets a table style override Normal in a cell, but not a style the paragraph names', () => {
    const doc = load({
      styles: `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr><w:spacing w:before="160" w:after="200"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Body"><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="120"/></w:pPr></w:style>
        <w:style w:type="table" w:styleId="Grid"><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:style>`,
      body: `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr><w:tc>${p('normal')}${p('named', '<w:pStyle w:val="Body"/>')}</w:tc></w:tr></w:tbl>`,
    });
    const cell = (blocks(doc)[0] as DocxTable).rows[0]![0]!;
    const [normal, named] = cell.content as DocxParagraph[];
    expect(normal!.spaceAfterPt).toBe(0);
    expect(normal!.spaceBeforePt).toBe(0);
    expect(named!.spaceAfterPt).toBe(6); // the named style wins
  });

  it('drops spacing between neighbours of the same style under contextualSpacing', () => {
    const doc = load({
      styles: `<w:style w:type="paragraph" w:styleId="Item"><w:pPr><w:spacing w:before="120" w:after="120"/><w:contextualSpacing/></w:pPr></w:style>`,
      body: p('a', '<w:pStyle w:val="Item"/>') + p('b', '<w:pStyle w:val="Item"/>') + p('c'),
    });
    const [a, b] = paras(doc);
    expect(a!.spaceAfterPt).toBe(0);
    expect(b!.spaceBeforePt).toBe(0);
    expect(b!.spaceAfterPt).toBe(6); // the next paragraph is another style
  });

  it('treats an unstated widowControl as on, as Word does', () => {
    const doc = load({ body: p('on') + p('off', '<w:widowControl w:val="0"/>') });
    const [on, off] = paras(doc);
    expect(on!.widowControl).toBe(true);
    expect(off!.widowControl).toBeUndefined();
  });

  it('separates keepNext from keepLines', () => {
    const [a] = paras(load({ body: p('k', '<w:keepNext/>') }));
    expect(a!.keepNext).toBe(true);
    expect(a!.keepLines).toBeUndefined();
  });
});

describe('runs', () => {
  it('marks a page break apart from a line break, and keeps hidden text out of sight', () => {
    const doc = load({
      body: `<w:p><w:r><w:t>one</w:t><w:br w:type="page"/><w:t>two</w:t><w:br/><w:t>three</w:t></w:r>
        <w:r><w:rPr><w:vanish/></w:rPr><w:t>secret</w:t></w:r></w:p>`,
    });
    const runs = paras(doc)[0]!.runs;
    expect(runs[1]).toMatchObject({ text: 'two', breakBefore: true, pageBreakBefore: true });
    expect(runs[2]!.pageBreakBefore).toBeUndefined();
    expect(runs[2]!.breakBefore).toBe(true);
    expect(runs[3]).toMatchObject({ text: 'secret', hidden: true });
  });

  it('tracks nested fields, so a TOC entry\'s PAGEREF result carries its own code', () => {
    const doc = load({
      body: `<w:p>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \\o</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>Intro</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGEREF _Toc1</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
        <w:r><w:t>after</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:p><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>`,
    });
    const [toc, simple] = paras(doc);
    expect(toc!.runs.map((r) => [r.text, r.fieldCode?.trim()])).toEqual([
      ['Intro', 'TOC \\o'],
      ['3', 'PAGEREF _Toc1'],
      ['after', 'TOC \\o'],
    ]);
    expect(simple!.runs[0]!.fieldCode?.trim()).toBe('NUMPAGES');
  });

  it('attributes tracked insertions and deletions to their reviewer', () => {
    const doc = load({
      body: `<w:p><w:r><w:t xml:space="preserve">keep </w:t></w:r><w:ins w:author="Ann"><w:r><w:t>added</w:t></w:r></w:ins><w:del w:author="Bo"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>`,
    });
    const runs = paras(doc)[0]!.runs;
    expect(runs.map((r) => [r.text, r.revision?.kind, r.revision?.author])).toEqual([
      ['keep ', undefined, undefined],
      ['added', 'ins', 'Ann'],
      ['gone', 'del', 'Bo'],
    ]);
  });
});

describe('numbering', () => {
  it('maps a symbol-font bullet to Unicode but keeps the font for line height', () => {
    const doc = load({
      numbering: `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="\uf0b7"/>
          <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl>
          <w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:suff w:val="space"/>
          <w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr></w:lvl></w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
      body:
        p('a', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
        p('b', '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>'),
    });
    const [a, b] = paras(doc);
    expect(a!.listMarker).toBe('•');
    expect(a!.listMarkerFormat?.font).not.toBe('Symbol');
    expect(a!.listMarkerFormat?.metricsFont).toBe('Symbol');
    expect(b!.listMarker).toBe('o');
    expect(b!.listMarkerFormat?.font).toBe('Courier New');
    expect(b!.listSuffix).toBe('space');
  });

  it('restarts a list at a startOverride and repeats letters past z', () => {
    const lvl = (fmt: string) =>
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%1)"/></w:lvl></w:abstractNum>`;
    const doc = load({
      numbering: `${lvl('lowerLetter')}<w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="27"/></w:lvlOverride></w:num>`,
      body: p('x', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') + p('y', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
    });
    expect(paras(doc).map((q) => q.listMarker)).toEqual(['aa)', 'bb)']);
  });
});

describe('tables', () => {
  it("gives an unstyled table Word's default cell margins, from the default table style", () => {
    const doc = load({
      styles: `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:tblPr><w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
        </w:tblCellMar></w:tblPr></w:style>`,
      body: `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcMar><w:top w:w="60" w:type="dxa"/></w:tcMar></w:tcPr>${p('x')}</w:tc></w:tr></w:tbl>`,
    });
    const cell = (blocks(doc)[0] as DocxTable).rows[0]![0]!;
    expect(cell.cellPaddingPx).toEqual({ top: 4, right: 7.2, bottom: 0, left: 7.2 });
  });

  it('reads row heights, repeating header rows, alignment and cantSplit from the style', () => {
    const doc = load({
      styles: `<w:style w:type="table" w:styleId="Keep"><w:trPr><w:cantSplit/></w:trPr></w:style>`,
      body: `<w:tbl><w:tblPr><w:tblStyle w:val="Keep"/><w:jc w:val="center"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
        <w:tr><w:trPr><w:tblHeader/><w:trHeight w:val="400" w:hRule="exact"/></w:trPr><w:tc>${p('h')}</w:tc></w:tr>
        <w:tr><w:trPr><w:cantSplit w:val="0"/><w:trHeight w:val="300"/></w:trPr><w:tc>${p('a')}</w:tc></w:tr>
        <w:tr><w:tc>${p('b')}</w:tc></w:tr></w:tbl>`,
    });
    const table = blocks(doc)[0] as DocxTable;
    expect(table.align).toBe('center');
    expect(table.headerRows).toBe(1);
    expect(table.rowHeights?.[0]).toEqual({ px: 400 / 15, rule: 'exact' });
    expect(table.rowHeights?.[1]).toEqual({ px: 20, rule: 'atLeast' });
    expect(table.cantSplit).toEqual([true, undefined, true]);
  });

  it('keeps a fixed-width table wider than the column instead of shrinking it', () => {
    const doc = load({
      body: `<w:tbl><w:tblGrid><w:gridCol w:w="10000"/></w:tblGrid><w:tr><w:tc>${p('x')}</w:tc></w:tr></w:tbl>
        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`,
    });
    // 10000 twips (666.7px) > the 624px column, but < the 720px to the page edge.
    expect((blocks(doc)[0] as DocxTable).widthPx).toBeCloseTo(666.67, 1);
  });
});

describe('sections and notes', () => {
  it('carries a header on into a section that references none', () => {
    const hdr = `<w:hdr xmlns:w="${W}">${p('Running head')}</w:hdr>`;
    const doc = load({
      extra: { 'header1.xml': hdr },
      extraRels: `<Relationship Id="rH" Type="${DocxRelType.Header}" Target="header1.xml"/>`,
      body: `${p('one', '<w:sectPr><w:headerReference w:type="default" r:id="rH"/></w:sectPr>')}${p('two')}
        <w:sectPr><w:type w:val="continuous"/><w:pgNumType w:start="5" w:fmt="lowerRoman"/></w:sectPr>`,
    });
    expect(doc.sections).toHaveLength(2);
    expect((doc.sections[1]!.header![0] as DocxParagraph).runs[0]!.text).toBe('Running head');
    expect(doc.sections[1]).toMatchObject({ start: 'continuous', pageNumberStart: 5, pageNumberFormat: 'lowerRoman' });
  });

  it('numbers footnotes in reference order and parses their content with the matching mark', () => {
    const note = (id: string, text: string) =>
      `<w:footnote w:id="${id}"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:footnote>`;
    const doc = load({
      footnotes: `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>${note('7', 'seven')}${note('3', 'three')}`,
      body: `<w:p><w:r><w:t>a</w:t></w:r><w:r><w:footnoteReference w:id="7"/></w:r><w:r><w:t>b</w:t></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p>`,
    });
    const refs = paras(doc)[0]!.runs.filter((r) => r.note);
    expect(refs.map((r) => r.text)).toEqual(['1', '2']);
    const first = refs[0]!.note!.blocks[0] as DocxParagraph;
    expect(first.runs.map((r) => r.text).join('')).toBe('1 seven');
    const sep = doc.sections[0]!.footnoteSeparator![0] as DocxParagraph;
    expect(sep.runs[0]!.separator).toBe(true);
  });
});

describe('drawings', () => {
  it('resolves a picture duotone through the theme, including satMod', () => {
    const doc = load({
      theme: THEME,
      extraRels: `<Relationship Id="rImg" Type="${DocxRelType.Image}" Target="media/p.png"/>`,
      body: `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="952500" cy="952500"/><a:graphic><a:graphicData>
        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rImg">
          <a:duotone><a:prstClr val="black"/><a:schemeClr val="accent1"><a:satMod val="50000"/></a:schemeClr></a:duotone>
        </a:blip></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    });
    const img = paras(doc)[0]!.runs[0]!.drawing!;
    expect(img.kind).toBe('image');
    if (img.kind !== 'image') return;
    expect(img.duotone?.[0]).toBe('000000');
    // accent1 4472C4 at half saturation is greyer, same lightness.
    expect(img.duotone?.[1]).not.toBe('4472C4');
    expect(img.duotone?.[1]).toMatch(/^[0-9A-F]{6}$/);
  });
});

describe('layout features', () => {
  it('reads odd/even headers only together with the evenAndOddHeaders setting', () => {
    const hdr = (t: string) => `<w:hdr xmlns:w="${W}">${p(t)}</w:hdr>`;
    const doc = load({
      settings: '<w:evenAndOddHeaders/>',
      extra: { 'header1.xml': hdr('odd'), 'header2.xml': hdr('even') },
      extraRels: `<Relationship Id="rH1" Type="${DocxRelType.Header}" Target="header1.xml"/>
        <Relationship Id="rH2" Type="${DocxRelType.Header}" Target="header2.xml"/>`,
      body: `${p('x')}<w:sectPr><w:headerReference w:type="default" r:id="rH1"/><w:headerReference w:type="even" r:id="rH2"/></w:sectPr>`,
    });
    const s = doc.sections[0]!;
    expect((s.header![0] as DocxParagraph).runs[0]!.text).toBe('odd');
    expect((s.evenHeader![0] as DocxParagraph).runs[0]!.text).toBe('even');
    expect(s.evenAndOddHeaders).toBe(true);
  });

  it('lays out equal-width columns with their gap and separator', () => {
    const doc = load({
      body: `${p('x')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        <w:cols w:num="2" w:space="720" w:sep="1"/></w:sectPr>`,
    });
    const cols = doc.sections[0]!.columns!;
    expect(cols.separator).toBe(true);
    // 6.5in text width: two 3in columns either side of a 0.5in gap.
    expect(cols.cols.map((c) => [Math.round(c.xPx), Math.round(c.widthPx)])).toEqual([
      [0, 288],
      [336, 288],
    ]);
  });

  it('marks right-to-left paragraphs and runs, sizing rtl text by its complex-script props', () => {
    const doc = load({
      body: `<w:p><w:pPr><w:bidi/><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:rtl/><w:sz w:val="20"/><w:szCs w:val="32"/></w:rPr><w:t>שלום</w:t></w:r></w:p>`,
    });
    const para = paras(doc)[0]!;
    expect(para.rtl).toBe(true);
    expect(para.runs[0]!.rtl).toBe(true);
    expect(para.runs[0]!.sizePt).toBe(16);
  });

  it('gives East Asian text its own font', () => {
    const doc = load({ body: p('日本語', '', '<w:rFonts w:ascii="Arial" w:eastAsia="MS Mincho"/>') });
    expect(paras(doc)[0]!.runs[0]!.eastAsiaFont).toBe('MS Mincho');
  });

  it('converts an equation to a math tree', () => {
    const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
    const doc = load({
      body: `<w:p><m:oMath xmlns:m="${M}"><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>
        <m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></w:p>`,
    });
    expect(paras(doc)[0]!.runs.find((r) => r.math)!.math).toEqual({
      t: 'row',
      c: [
        { t: 'frac', num: { t: 'tok', kind: 'mi', v: 'a' }, den: { t: 'tok', kind: 'mn', v: '2' } },
        { t: 'script', base: { t: 'tok', kind: 'mi', v: 'x' }, sup: { t: 'tok', kind: 'mn', v: '2' } },
      ],
    });
  });

  it('ties commented runs to their comment and carries its author and text', () => {
    const doc = load({
      extra: {
        'comments.xml': `<w:comments xmlns:w="${W}"><w:comment w:id="4" w:author="Ann" w:initials="A">${p('Check this')}</w:comment></w:comments>`,
      },
      extraRels: `<Relationship Id="rC" Type="${DocxRelType.Comments}" Target="comments.xml"/>`,
      body: `<w:p><w:r><w:t xml:space="preserve">plain </w:t></w:r><w:commentRangeStart w:id="4"/><w:r><w:t>noted</w:t></w:r>
        <w:commentRangeEnd w:id="4"/><w:r><w:commentReference w:id="4"/></w:r></w:p>`,
    });
    const runs = paras(doc)[0]!.runs;
    expect(runs.find((r) => r.text === 'plain ')!.commentIds).toBeUndefined();
    expect(runs.find((r) => r.text === 'noted')!.commentIds).toEqual(['4']);
    const c = runs.find((r) => r.comment)!.comment!;
    expect(c.author).toBe('Ann');
    expect((c.blocks[0] as DocxParagraph).runs[0]!.text).toBe('Check this');
  });

  it('reads a square-wrapped floating picture as a float on the given side', () => {
    const doc = load({
      extraRels: `<Relationship Id="rImg" Type="${DocxRelType.Image}" Target="media/p.png"/>`,
      body: `<w:p><w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
          distT="0" distB="0" distL="114300" distR="114300" behindDoc="0" relativeHeight="1">
        <wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
        <wp:extent cx="952500" cy="952500"/><wp:wrapSquare wrapText="left"/><a:graphic><a:graphicData>
        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rImg"/></pic:blipFill></pic:pic>
        </a:graphicData></a:graphic></wp:anchor></w:drawing></w:r><w:r><w:t>text</w:t></w:r></w:p>`,
    });
    const a = paras(doc)[0]!.anchors![0]!;
    expect(a.wrapText).toBe('left');
    expect(a.drawing.kind).toBe('image');
  });

  it('parses an embedded chart from its chart part', () => {
    const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
    const chart = `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:title/><c:plotArea>
      <c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Y</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:xVal><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:xVal>
        <c:yVal><c:numRef><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser>
        <c:axId val="1"/><c:axId val="2"/></c:scatterChart>
      <c:valAx><c:axId val="1"/><c:majorGridlines/></c:valAx><c:valAx><c:axId val="2"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;
    const doc = load({
      theme: THEME,
      extra: { 'charts/chart1.xml': chart },
      extraRels: `<Relationship Id="rCh" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>`,
      body: `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="4572000" cy="2743200"/><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="rCh"/></a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p>`,
    });
    const d = paras(doc)[0]!.runs[0]!.drawing!;
    expect(d.kind).toBe('chart');
    if (d.kind !== 'chart') return;
    expect(d.widthPx).toBe(480);
    // An empty <c:title> on a one-series chart is titled with the series name.
    expect(d.chart).toMatchObject({ kind: 'scatter', title: 'Y', xGridlines: true });
    expect(d.chart.series[0]).toMatchObject({ values: [3, 4], xValues: [1, 2], line: true });
  });

  it('places the shapes of a group through its child coordinate space', () => {
    const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
    const WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
    const sp = (x: number) =>
      `<wps:wsp><wps:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="ellipse"/></wps:spPr><wps:bodyPr/></wps:wsp>`;
    const doc = load({
      body: `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="1905000" cy="952500"/><a:graphic><a:graphicData uri="${WPG}">
        <wpg:wgp xmlns:wpg="${WPG}" xmlns:wps="${WPS}"><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/>
          <a:chOff x="0" y="0"/><a:chExt cx="200" cy="100"/></a:xfrm></wpg:grpSpPr>${sp(0)}${sp(100)}</wpg:wgp>
        </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    });
    const g = paras(doc)[0]!.runs[0]!.drawing!;
    expect(g.kind).toBe('group');
    if (g.kind !== 'group') return;
    // 200 × 100 child units mapped onto a 200 × 100 px box.
    expect(g.items.map((i) => [Math.round(i.xPx), Math.round(i.wPx)])).toEqual([
      [0, 100],
      [100, 100],
    ]);
    expect(g.items[0]!.drawing).toMatchObject({ kind: 'shape', preset: 'ellipse' });
  });
});
