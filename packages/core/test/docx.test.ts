import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { DocxDocument } from '../src/docx/document/document.js';
import { DocxRelType } from '../src/docx/relTypes.js';
import { twipToPx, halfPtToPt, borderSzToPx } from '../src/docx/units.js';
import type { DocxParagraph, DocxTable } from '../src/docx/model.js';

const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function rels(entries: string): string {
  return `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;
}

/** A structurally complete minimal .docx exercising styles, lists, and tables. */
function buildDocx(): Uint8Array {
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
      </Types>`,
    ),
    '_rels/.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${DocxRelType.OfficeDocument}" Target="word/document.xml"/>`),
    ),
    'word/_rels/document.xml.rels': strToU8(
      rels(
        `<Relationship Id="rId1" Type="${DocxRelType.Styles}" Target="styles.xml"/>
         <Relationship Id="rId2" Type="${DocxRelType.Numbering}" Target="numbering.xml"/>
         <Relationship Id="rHdr" Type="${DocxRelType.Header}" Target="header1.xml"/>
         <Relationship Id="rFtr" Type="${DocxRelType.Footer}" Target="footer1.xml"/>
         <Relationship Id="rLink" Type="${OFFICE}/hyperlink" Target="https://example.com" TargetMode="External"/>`,
      ),
    ),
    'word/header1.xml': strToU8(
      `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:p><w:r><w:t>Header text</w:t></w:r></w:p>
      </w:hdr>`,
    ),
    'word/footer1.xml': strToU8(
      `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:p>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> STYLEREF Heading1 </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>Cached Title</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
          <w:r><w:tab/></w:r>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>1</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>
      </w:ftr>`,
    ),
    'word/styles.xml': strToU8(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:docDefaults>
          <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
        </w:docDefaults>
        <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
          <w:name w:val="Normal"/>
        </w:style>
        <w:style w:type="paragraph" w:styleId="Heading1">
          <w:name w:val="heading 1"/>
          <w:basedOn w:val="Normal"/>
          <w:pPr><w:jc w:val="center"/></w:pPr>
          <w:rPr><w:b/><w:sz w:val="48"/><w:color w:val="2E74B5"/></w:rPr>
        </w:style>
      </w:styles>`,
    ),
    'word/numbering.xml': strToU8(
      `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:abstractNum w:abstractNumId="0">
          <w:lvl w:ilvl="0">
            <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
            <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
          </w:lvl>
        </w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
      </w:numbering>`,
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:r="${OFFICE}">
        <w:body>
          <w:p>
            <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
            <w:r><w:t>Title</w:t></w:r>
          </w:p>
          <w:p>
            <w:r><w:t xml:space="preserve">Hello </w:t></w:r>
            <w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>
            <w:hyperlink r:id="rLink"><w:r><w:t>link</w:t></w:r></w:hyperlink>
          </w:p>
          <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>
          <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>
          <w:tbl>
            <w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>
            <w:tr>
              <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Wide</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
              <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
              <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
            </w:tr>
          </w:tbl>
          <w:sectPr>
            <w:headerReference w:type="default" r:id="rHdr"/>
            <w:footerReference w:type="default" r:id="rFtr"/>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
          </w:sectPr>
        </w:body>
      </w:document>`,
    ),
  };
  return zipSync(files);
}

describe('docx units', () => {
  it('converts Word units to px/pt', () => {
    expect(twipToPx(1440)).toBeCloseTo(96); // 1 inch
    expect(halfPtToPt(24)).toBe(12);
    expect(borderSzToPx(8)).toBeCloseTo(96 / 72); // 1 pt
  });
});

describe('DocxDocument.load', () => {
  const doc = DocxDocument.load(buildDocx());
  const section = doc.sections[0]!;

  it('produces one section with the section page size and margins', () => {
    expect(doc.sections).toHaveLength(1);
    expect(section.size.wPx).toBeCloseTo(twipToPx(12240));
    expect(section.size.hPx).toBeCloseTo(twipToPx(15840));
    expect(section.margins.leftPx).toBeCloseTo(96);
  });

  it('resolves the header and footer from the section references', () => {
    const header = section.header![0] as DocxParagraph;
    const footer = section.footer![0] as DocxParagraph;
    expect(header.runs.map((r) => r.text).join('')).toBe('Header text');
    expect(footer.runs.map((r) => r.text).join('')).toBe('Cached Title1');
  });

  it('parses field characters and attaches field codes to runs', () => {
    const footer = section.footer![0] as DocxParagraph;
    const stylerefRun = footer.runs.find((r) => r.text === 'Cached Title')!;
    const pageRun = footer.runs.find((r) => r.text === '1')!;
    expect(stylerefRun.fieldCode?.trim()).toBe('STYLEREF Heading1');
    expect(pageRun.fieldCode?.trim()).toBe('PAGE');
  });

  it('resolves the style cascade (Heading1: centered, bold, larger, colored)', () => {
    const title = section.blocks[0] as DocxParagraph;
    expect(title.kind).toBe('paragraph');
    expect(title.styleName).toBe('heading 1');
    expect(title.align).toBe('ctr');
    expect(title.baseBold).toBe(true);
    expect(title.baseFontSizePt).toBe(24);
    expect(title.baseColorHex).toBe('2e74b5');
    expect(title.runs.map((r) => r.text).join('')).toBe('Title');
  });

  it('parses runs with direct formatting and hyperlinks, preserving whitespace', () => {
    const p = section.blocks[1] as DocxParagraph;
    expect(p.runs.map((r) => r.text)).toEqual(['Hello ', 'bold', 'link']);
    expect(p.runs[1]!.bold).toBe(true);
    expect(p.runs[2]!.hyperlink).toBe('https://example.com');
  });

  it('numbers an ordered list with running counters', () => {
    const list = section.blocks.filter(
      (e): e is DocxParagraph => e.kind === 'paragraph' && e.listMarker !== undefined,
    );
    expect(list.map((p) => p.listMarker)).toEqual(['1.', '2.']);
    expect(list[0]!.indentLeftPx).toBeCloseTo(twipToPx(720));
  });

  it('builds a table grid with a horizontal span', () => {
    const table = section.blocks.find((e): e is DocxTable => e.kind === 'table')!;
    expect(table.colWidths).toHaveLength(2);
    expect(table.rows).toHaveLength(2);
    // First row: one cell spanning 2 columns, plus a null placeholder.
    expect(table.rows[0]![0]!.colSpan).toBe(2);
    expect(table.rows[0]![1]).toBeNull();
    // Second row: two independent cells.
    expect(table.rows[1]![0]!.colSpan).toBe(1);
    expect(table.rows[1]![1]!.colSpan).toBe(1);
  });
});
