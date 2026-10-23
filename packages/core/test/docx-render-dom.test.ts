/**
 * @vitest-environment jsdom
 *
 * The DOM side of DOCX viewing fidelity: what the renderer draws for tabs,
 * list markers, hidden text, note marks, fields and footnotes — and that the
 * pieces which are not content survive an edit round trip.
 */
import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { DocxDocument } from '../src/docx/document/document.js';
import { DocxRelType } from '../src/docx/relTypes.js';
import { DocxEditContext } from '../src/docx/edit/context.js';
import { reconcileParagraph } from '../src/docx/edit/reconcile.js';
import { renderBlock, renderPage } from '../src/docx/render/dom.js';
import { renderChart } from '../src/oxml/drawingml/chart/render.js';
import { EDIT_ATTR } from '../src/oxml/edit/attrs.js';
import { canonicalFont, fontStack, lineHeightRatio, fallbackFontFaceCss } from '../src/oxml/fontMetrics.js';
import type { DocxPage, DocxParagraph } from '../src/docx/model.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

function load(body: string, footnotes?: string): DocxDocument {
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
      </Types>`,
    ),
    '_rels/.rels': strToU8(rels(`<Relationship Id="rId1" Type="${DocxRelType.OfficeDocument}" Target="word/document.xml"/>`)),
    'word/document.xml': strToU8(`<w:document xmlns:w="${W}" xmlns:m="${M}"><w:body>${body}</w:body></w:document>`),
    'word/numbering.xml': strToU8(
      `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    ),
  };
  let docRels = `<Relationship Id="rNum" Type="${DocxRelType.Numbering}" Target="numbering.xml"/>`;
  if (footnotes) {
    files['word/footnotes.xml'] = strToU8(`<w:footnotes xmlns:w="${W}">${footnotes}</w:footnotes>`);
    docRels += `<Relationship Id="rFn" Type="${DocxRelType.Footnotes}" Target="footnotes.xml"/>`;
  }
  files['word/_rels/document.xml.rels'] = strToU8(rels(docRels));
  return DocxDocument.load(zipSync(files));
}

const firstPara = (doc: DocxDocument) => doc.sections[0]!.blocks[0] as DocxParagraph;
const deps = { imageUrl: () => undefined };

describe('docx render — paragraphs', () => {
  it('puts the list marker and its tab in one uneditable span, apart from content tabs', () => {
    const doc = load(
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
        <w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>`,
    );
    const el = renderBlock(firstPara(doc), deps);
    const marker = el.firstElementChild as HTMLElement;
    expect(marker.getAttribute('contenteditable')).toBe('false');
    expect(marker.textContent).toBe('1.');
    // The marker's tab is layout only; just the run's tab is content.
    expect(marker.querySelectorAll('[data-cbx-tabstop]')).toHaveLength(1);
    expect(marker.querySelector(`[${EDIT_ATTR.tab}]`)).toBeNull();
    expect(el.querySelectorAll(`[${EDIT_ATTR.tab}]`)).toHaveLength(1);
  });

  it('keeps every space Word was given and sizes lines by the font, not a fixed multiple', () => {
    const doc = load(
      `<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>
        <w:r><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr><w:t xml:space="preserve">two  spaces</w:t></w:r></w:p>`,
    );
    const el = renderBlock(firstPara(doc), deps);
    expect(el.style.whiteSpace).toBe('pre-wrap');
    // 1.5 × Times New Roman's own single line (the paragraph mark's font).
    expect(Number(el.style.lineHeight)).toBeCloseTo(1.5 * lineHeightRatio('Times New Roman'), 5);
  });

  it('renders hidden text invisibly but keeps it for editing', () => {
    const doc = load(`<w:p><w:r><w:t xml:space="preserve">shown </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r></w:p>`);
    const ctx = new DocxEditContext(doc);
    const el = renderBlock(firstPara(doc), { ...deps, edit: ctx });
    const spans = el.querySelectorAll(`[${EDIT_ATTR.run}]`);
    expect((spans[1] as HTMLElement).style.display).toBe('none');
    const segs = reconcileParagraph(el, (k) => ctx.resolve(k))[0]!.segments;
    expect(segs.map((s) => s.text).join('')).toBe('shown hidden');
  });

  it('reads a content tab back as a tab after editing', () => {
    const doc = load(`<w:p><w:r><w:t>left</w:t><w:tab/><w:t>right</w:t></w:r></w:p>`);
    const ctx = new DocxEditContext(doc);
    const el = renderBlock(firstPara(doc), { ...deps, edit: ctx });
    const segs = reconcileParagraph(el, (k) => ctx.resolve(k))[0]!.segments;
    expect(segs.map((s) => [s.text, !!s.tabBefore])).toEqual([
      ['left', false],
      ['right', true],
    ]);
  });

  it('draws a footnote reference mark that is never typed into', () => {
    const doc = load(
      `<w:p><w:r><w:t>text</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p>`,
      `<w:footnote w:id="2"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:footnote>`,
    );
    const ctx = new DocxEditContext(doc);
    const el = renderBlock(firstPara(doc), { ...deps, edit: ctx });
    const mark = [...el.querySelectorAll('span')].find((s) => s.textContent === '1')!;
    expect(mark.getAttribute('contenteditable')).toBe('false');
    // Mapped as one object atom: the caret steps over it, never into it.
    expect(mark.getAttribute(EDIT_ATTR.part)).toBe('obj');
    expect(mark.style.verticalAlign).toBe('super');
  });
});

describe('docx render — pages', () => {
  const page = (doc: DocxDocument, extra: Partial<DocxPage> = {}): DocxPage => {
    const s = doc.sections[0]!;
    return { index: 4, size: s.size, margins: s.margins, elements: s.blocks, ...extra };
  };
  const field = (code: string, cached: string) =>
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>${code}</w:instrText></w:r>
     <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${cached}</w:t></w:r><w:r><w:t>X</w:t></w:r>
     <w:r><w:fldChar w:fldCharType="end"/></w:r>`;

  it('fills PAGE, NUMPAGES and SECTIONPAGES for the page, once per field', () => {
    const doc = load(
      `<w:p>${field(' PAGE \\* roman ', '1')}<w:r><w:t xml:space="preserve"> of </w:t></w:r>${field(' NUMPAGES ', '1')}<w:r><w:t xml:space="preserve"> / </w:t></w:r>${field(' SECTIONPAGES ', '1')}</w:p>`,
    );
    const sheet = renderPage(page(doc, { pageNumber: '5', pageCount: 12, sectionPageCount: 3 }), deps);
    // The field's second cached run renders empty: the value replaces the result.
    expect(sheet.textContent).toContain('v of 12 / 3');
  });

  it('is its own stacking context, so behind-text drawings stay on the page', () => {
    const doc = load(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(renderPage(page(doc), deps).style.isolation).toBe('isolate');
  });

  it("draws the page's footnotes in a band above the bottom margin", () => {
    const doc = load(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`);
    const note = load(`<w:p><w:r><w:t>the note</w:t></w:r></w:p>`).sections[0]!.blocks;
    const sheet = renderPage(page(doc, { footnotes: note, contentBottomPx: 96 }), deps);
    const band = [...sheet.children].find((c) => c.textContent === 'the note') as HTMLElement;
    expect(band.style.position).toBe('absolute');
    expect(band.style.bottom).toBe('96px');
  });
});

describe('font metrics', () => {
  it("maps legacy names to the face Word draws and stacks a metric-matched fallback", () => {
    expect(canonicalFont('Tms Rmn')).toBe('Times New Roman');
    expect(fontStack('Calibri')).toBe('Calibri, cbx-fb-calibri, sans-serif');
    expect(fontStack('Cambria')).toBe('Cambria, cbx-fb-cambria, serif');
    expect(fontStack('Some Unknown Face')).toBe('"Some Unknown Face", sans-serif');
  });

  it("uses each font's own single line height", () => {
    expect(lineHeightRatio('Calibri')).toBeCloseTo(1.2205);
    expect(lineHeightRatio('Arial')).toBeCloseTo(1.15);
    expect(lineHeightRatio(undefined)).toBeGreaterThan(1);
  });

  it('scales the fallback face to the original width, metrics divided back out', () => {
    const css = fallbackFontFaceCss();
    const calibri = css.split('\n').find((r) => r.includes('"cbx-fb-calibri"') && r.includes('font-weight:normal;font-style:normal'))!;
    expect(calibri).toContain('size-adjust:90.17%');
    expect(calibri).toContain('local("Arial")');
    // 1.2205 / 0.9017 → the face's line comes out at Calibri's 1.2205em.
    expect(calibri).toMatch(/ascent-override:10[45]\.\d+%/);
  });
});

describe('docx render — review and special text', () => {
  const tracked = `<w:p><w:r><w:t xml:space="preserve">keep </w:t></w:r><w:ins w:author="Ann"><w:r><w:t>added</w:t></w:r></w:ins>
    <w:del w:author="Bo"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>`;
  const visible = (el: HTMLElement) => el.textContent!.replace(/\u200b/g, '');

  it('shows insertions underlined and deletions struck through in markup view', () => {
    const el = renderBlock(firstPara(load(tracked)), deps);
    expect(visible(el)).toBe('keep addedgone');
    const ins = [...el.querySelectorAll('span')].find((s) => s.textContent === 'added')!;
    const del = [...el.querySelectorAll('span')].find((s) => s.textContent === 'gone')!;
    expect(ins.style.textDecorationLine || ins.style.textDecoration).toContain('underline');
    expect(del.style.textDecorationLine || del.style.textDecoration).toContain('line-through');
    expect(del.getAttribute('contenteditable')).toBe('false');
  });

  it('shows the final text, or the original, when asked', () => {
    const doc = load(tracked);
    expect(visible(renderBlock(firstPara(doc), { ...deps, revisions: 'final' }))).toBe('keep added');
    expect(visible(renderBlock(firstPara(doc), { ...deps, revisions: 'original' }))).toBe('keep gone');
  });

  it('draws an equation as MathML', () => {
    const doc = load(`<w:p><m:oMath><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad></m:oMath></w:p>`);
    const math = renderBlock(firstPara(doc), deps).querySelector('math')!;
    expect(math).not.toBeNull();
    expect(math.querySelector('msqrt mi')!.textContent).toBe('x');
  });
});

describe('charts', () => {
  it('joins scatter points with a line and labels both value axes on Excel steps', () => {
    const svg = renderChart(
      { kind: 'scatter', categories: [], series: [{ values: [2.7, 3.2, 0.8], xValues: [0.7, 1.8, 2.6], line: true }], xGridlines: true },
      480,
      280,
    );
    expect(svg.querySelectorAll('polyline')).toHaveLength(1);
    const labels = [...svg.querySelectorAll('text')].map((t) => t.textContent);
    expect(labels).toEqual(expect.arrayContaining(['0', '0.5', '1', '3', '3.5']));
    expect(labels).not.toContain('4');
  });
});
