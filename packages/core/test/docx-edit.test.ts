/**
 * DOCX text editing: source addressing, XML write-back, export round-trip.
 *
 * No DOM here — edits are expressed directly as DocxParaEdit lists, so the XML
 * and packaging halves are tested independently of contentEditable.
 */
import { describe, it, expect } from 'vitest';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { DocxDocument } from '../src/docx/document/document.js';
import { DocxRelType } from '../src/docx/relTypes.js';
import { DocxEditSession } from '../src/docx/edit/session.js';
import type { DocxParaEdit } from '../src/docx/edit/xmlWrite.js';
import type { DocxParagraph, DocxSection } from '../src/docx/model.js';

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

/**
 * A document exercising the cases that make DOCX different: a two-run
 * paragraph, a single `<w:r>` holding TWO `<w:t>` children, an ordered list
 * (numbering counters), and a header (a separate part — read-only).
 */
function buildDocx(): Uint8Array {
  return zipSync({
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
        `<Relationship Id="rId10" Type="${DocxRelType.Numbering}" Target="numbering.xml"/>` +
          `<Relationship Id="rId11" Type="${DocxRelType.Header}" Target="header1.xml"/>`,
      ),
    ),
    'word/numbering.xml': strToU8(
      `<w:numbering xmlns:w="w">
        <w:abstractNum w:abstractNumId="0">
          <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
        </w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
      </w:numbering>`,
    ),
    'word/header1.xml': strToU8(
      `<w:hdr xmlns:w="w"><w:p><w:r><w:t>From header</w:t></w:r></w:p></w:hdr>`,
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="w"><w:body>
        <w:p w14:paraId="AAAA1111">
          <w:pPr><w:jc w:val="center"/></w:pPr>
          <w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t xml:space="preserve">Hello </w:t></w:r>
          <w:r><w:rPr><w:b/><w:i/><w:color w:val="FF0000"/><w:sz w:val="48"/></w:rPr><w:t>World</w:t></w:r>
        </w:p>
        <w:p>
          <w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">One </w:t><w:t>Two</w:t></w:r>
        </w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Beta</w:t></w:r></w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Gamma</w:t></w:r></w:p>
        <w:sectPr><w:headerReference w:type="default" r:id="rId11" xmlns:r="r"/></w:sectPr>
      </w:body></w:document>`,
    ),
  });
}

/** Paragraphs of the first section, in order. */
function paras(doc: DocxDocument): DocxParagraph[] {
  return doc.sections[0]!.blocks.filter((b): b is DocxParagraph => b.kind === 'paragraph');
}

const textOf = (p: DocxParagraph) => p.runs.map((r) => r.text).join('');

/** A ParaEdit that keeps every run exactly as parsed. */
function identity(p: DocxParagraph): DocxParaEdit[] {
  return [
    {
      src: p,
      segments: p.runs.map((r) => ({
        text: r.text,
        src: r,
        ...(r.breakBefore ? { breakBefore: true } : {}),
        ...(r.tabBefore ? { tabBefore: true } : {}),
      })),
    },
  ];
}

const docXml = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)['word/document.xml']!);

describe('docx edit — addressing', () => {
  it('records the <w:p> of a paragraph and the <w:t> of each run', () => {
    const doc = DocxDocument.load(buildDocx());
    const p = paras(doc)[0]!;
    expect(doc.sourceOf(p)?.node.name).toBe('w:p');

    const run = doc.sourceOf(p.runs[0]!);
    expect(run?.node.name).toBe('w:t');
    // The owning <w:r> comes along because that is where <w:rPr> lives.
    expect(run?.owner?.name).toBe('w:r');
  });

  it('gives the two runs of a single <w:r> distinct <w:t> sources', () => {
    const doc = DocxDocument.load(buildDocx());
    const p = paras(doc)[1]!;
    expect(p.runs.map((r) => r.text)).toEqual(['One ', 'Two']);

    const a = doc.sourceOf(p.runs[0]!);
    const b = doc.sourceOf(p.runs[1]!);
    expect(a?.node).not.toBe(b?.node);
    // ...but they share one owning run element.
    expect(a?.owner).toBe(b?.owner);
  });

  it('marks body text editable and header text read-only', () => {
    const doc = DocxDocument.load(buildDocx());
    expect(doc.isEditable(paras(doc)[0]!)).toBe(true);

    const headerPara = doc.sections[0]!.header![0] as DocxParagraph;
    expect(textOf(headerPara)).toBe('From header');
    expect(doc.isEditable(headerPara)).toBe(false);
  });
});

describe('docx edit — round trip', () => {
  it('writes edited text back and survives export/reload', () => {
    const doc = DocxDocument.load(buildDocx());
    const session = new DocxEditSession(doc);
    const p = paras(doc)[0]!;

    const edits = identity(p);
    edits[0]!.segments[0]!.text = 'Goodbye ';
    expect(session.commitParagraph(p, edits)).toBeTruthy();

    const reloaded = DocxDocument.load(session.exportBytes());
    expect(textOf(paras(reloaded)[0]!)).toBe('Goodbye World');
  });

  it('leaves the untouched sibling run’s formatting intact', () => {
    const doc = DocxDocument.load(buildDocx());
    const session = new DocxEditSession(doc);
    const p = paras(doc)[0]!;

    const edits = identity(p);
    edits[0]!.segments[0]!.text = 'Goodbye ';
    session.commitParagraph(p, edits);

    const world = paras(DocxDocument.load(session.exportBytes()))[0]!.runs[1]!;
    expect(world.text).toBe('World');
    expect(world.bold).toBe(true);
    expect(world.italic).toBe(true);
    expect(world.colorHex).toBe('ff0000');
    expect(world.sizePt).toBe(24);
  });

  it('edits one <w:t> of a shared run without disturbing the other', () => {
    const doc = DocxDocument.load(buildDocx());
    const session = new DocxEditSession(doc);
    const p = paras(doc)[1]!;

    const edits = identity(p);
    edits[0]!.segments[0]!.text = 'Uno ';
    session.commitParagraph(p, edits);

    const after = paras(DocxDocument.load(session.exportBytes()))[1]!;
    expect(after.runs.map((r) => r.text)).toEqual(['Uno ', 'Two']);
    // Both halves keep the original run's size.
    expect(after.runs.every((r) => r.sizePt === 12)).toBe(true);
  });

  it('rewrites only word/document.xml, leaving every other part byte-identical', () => {
    const original = buildDocx();
    const doc = DocxDocument.load(original);
    const session = new DocxEditSession(doc);

    const p = paras(doc)[0]!;
    const edits = identity(p);
    edits[0]!.segments[0]!.text = 'Changed ';
    session.commitParagraph(p, edits);

    const before = unzipSync(original);
    const after = unzipSync(session.exportBytes());
    expect(Object.keys(after).sort()).toEqual(Object.keys(before).sort());
    for (const name of Object.keys(before)) {
      if (name === 'word/document.xml') continue;
      expect(after[name], `${name} should be untouched`).toEqual(before[name]);
    }
  });

  it('preserves paragraph properties and adds xml:space when needed', () => {
    const doc = DocxDocument.load(buildDocx());
    const session = new DocxEditSession(doc);
    const p = paras(doc)[0]!;

    const edits = identity(p);
    edits[0]!.segments[0]!.text = '  padded  ';
    session.commitParagraph(p, edits);

    const xml = docXml(session.exportBytes());
    expect(xml).toContain('<w:pPr><w:jc w:val="center"/></w:pPr>');
    expect(xml).toContain('<w:t xml:space="preserve">  padded  </w:t>');
  });

  it('refuses to edit header text', () => {
    const doc = DocxDocument.load(buildDocx());
    const session = new DocxEditSession(doc);
    const headerPara = doc.sections[0]!.header![0] as DocxParagraph;

    const edits = identity(headerPara);
    edits[0]!.segments[0]!.text = 'Tampered';
    expect(session.commitParagraph(headerPara, edits)).toBeUndefined();

    const reloaded = DocxDocument.load(session.exportBytes());
    expect(textOf(reloaded.sections[0]!.header![0] as DocxParagraph)).toBe('From header');
  });
});

describe('docx edit — numbering counters', () => {
  const markers = (sections: DocxSection[]) =>
    sections[0]!.blocks
      .filter((b): b is DocxParagraph => b.kind === 'paragraph')
      .map((p) => p.listMarker)
      .filter(Boolean);

  it('numbers an ordered list from 1 on the initial parse', () => {
    const doc = DocxDocument.load(buildDocx());
    expect(markers(doc.sections)).toEqual(['1.', '2.', '3.']);
  });

  it('still numbers from 1 after repeated rebuilds', () => {
    // Numbering carries live counters that marker() advances while walking the
    // body, so a rebuild that reused it would render "4. 5. 6." and then
    // "7. 8. 9." — this is the trap the fresh-Numbering rebuild exists to avoid.
    const doc = DocxDocument.load(buildDocx());
    doc.rebuild();
    doc.rebuild();
    expect(markers(doc.sections)).toEqual(['1.', '2.', '3.']);
  });

  it('keeps list numbering correct after an edit', () => {
    const doc = DocxDocument.load(buildDocx());
    const session = new DocxEditSession(doc);
    const list = paras(doc).filter((p) => p.listMarker);

    const edits = identity(list[0]!);
    edits[0]!.segments[0]!.text = 'Alpha edited';
    const sections = session.commitParagraph(list[0]!, edits)!;

    expect(markers(sections)).toEqual(['1.', '2.', '3.']);
    expect(markers(DocxDocument.load(session.exportBytes()).sections)).toEqual(['1.', '2.', '3.']);
  });
});

describe('docx edit — structure and formatting', () => {
  it('splits a paragraph, cloning its properties onto the new one', () => {
    const doc = DocxDocument.load(buildDocx());
    const session = new DocxEditSession(doc);
    const p = paras(doc)[0]!;

    session.commitParagraph(p, [
      { src: p, segments: [{ text: 'Hello ', src: p.runs[0]! }] },
      { src: p, segments: [{ text: 'World', src: p.runs[1]! }] },
    ]);

    const after = paras(DocxDocument.load(session.exportBytes()));
    expect(after.slice(0, 2).map(textOf)).toEqual(['Hello ', 'World']);
    // Both halves keep the centred alignment from the source <w:pPr>.
    expect(after[0]!.align).toBe('ctr');
    expect(after[1]!.align).toBe('ctr');
    // A split must not duplicate the source paragraph's unique id.
    expect(docXml(session.exportBytes()).match(/AAAA1111/g)).toBeNull();
  });

  it('merges two paragraphs into one', () => {
    const doc = DocxDocument.load(buildDocx());
    const session = new DocxEditSession(doc);
    const [first, second] = [paras(doc)[0]!, paras(doc)[1]!];

    // Fold the second paragraph's runs into the first, then empty the second.
    session.commitParagraph(first, [
      {
        src: first,
        segments: [
          ...first.runs.map((r) => ({ text: r.text, src: r })),
          ...second.runs.map((r) => ({ text: r.text, src: r })),
        ],
      },
    ]);
    const doc2 = DocxDocument.load(session.exportBytes());
    expect(textOf(paras(doc2)[0]!)).toBe('Hello WorldOne Two');
  });

  it('writes formatting as <w:rPr> child elements in half-points', () => {
    const doc = DocxDocument.load(buildDocx());
    const session = new DocxEditSession(doc);
    const p = paras(doc)[1]!;

    session.commitParagraph(p, [
      {
        src: p,
        segments: [
          { text: 'One ', src: p.runs[0]!, format: { bold: true, sizePt: 20, colorHex: '00FF00' } },
          { text: 'Two', src: p.runs[1]! },
        ],
      },
    ]);

    const xml = docXml(session.exportBytes());
    // Half-points, and correct CT_RPr child order (b before color before sz).
    expect(xml).toContain('<w:b/><w:color w:val="00FF00"/><w:sz w:val="40"/><w:szCs w:val="40"/>');

    const after = paras(DocxDocument.load(session.exportBytes()))[1]!;
    expect(after.runs[0]!.bold).toBe(true);
    expect(after.runs[0]!.sizePt).toBe(20);
    expect(after.runs[0]!.colorHex).toBe('00ff00');
    // The untouched run keeps its original size and no bold.
    expect(after.runs[1]!.sizePt).toBe(12);
    expect(after.runs[1]!.bold).toBeUndefined();
  });

  it('turns bold off explicitly rather than by omission', () => {
    const doc = DocxDocument.load(buildDocx());
    const session = new DocxEditSession(doc);
    const p = paras(doc)[0]!;

    session.commitParagraph(p, [
      {
        src: p,
        segments: [
          { text: 'Hello ', src: p.runs[0]! },
          { text: 'World', src: p.runs[1]!, format: { bold: false } },
        ],
      },
    ]);

    // A style could switch bold on, so "off" needs w:val="0", not a missing tag.
    expect(docXml(session.exportBytes())).toContain('<w:b w:val="0"/>');
    expect(paras(DocxDocument.load(session.exportBytes()))[0]!.runs[1]!.bold).toBeUndefined();
  });
});

describe('docx edit — history', () => {
  it('undoes and redoes an edit', () => {
    const doc = DocxDocument.load(buildDocx());
    const session = new DocxEditSession(doc);

    const edits = identity(paras(doc)[0]!);
    edits[0]!.segments[0]!.text = 'Goodbye ';
    session.commitParagraph(paras(doc)[0]!, edits);
    expect(textOf(paras(doc)[0]!)).toBe('Goodbye World');

    expect(session.canUndo).toBe(true);
    session.undo();
    expect(textOf(paras(doc)[0]!)).toBe('Hello World');

    expect(session.canRedo).toBe(true);
    session.redo();
    expect(textOf(paras(doc)[0]!)).toBe('Goodbye World');
  });

  it('drops the redo branch after a fresh edit', () => {
    const doc = DocxDocument.load(buildDocx());
    const session = new DocxEditSession(doc);

    const first = identity(paras(doc)[0]!);
    first[0]!.segments[0]!.text = 'A ';
    session.commitParagraph(paras(doc)[0]!, first);
    session.undo();
    expect(session.canRedo).toBe(true);

    const second = identity(paras(doc)[0]!);
    second[0]!.segments[0]!.text = 'B ';
    session.commitParagraph(paras(doc)[0]!, second);
    expect(session.canRedo).toBe(false);
    expect(textOf(paras(doc)[0]!)).toBe('B World');
  });
});
