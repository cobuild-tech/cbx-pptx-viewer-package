/**
 * Structural slide edits: deleting a slide.
 *
 * Deletion is cross-part surgery — presentation.xml, its relationships,
 * [Content_Types].xml and the slide part itself all have to agree afterwards,
 * or PowerPoint calls the file corrupt. These assert each of those, plus that
 * untouched parts still round-trip byte-identically and that undo/redo is exact.
 */
import { describe, it, expect } from 'vitest';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { Deck } from '../src/pptx/deck/deck.js';
import { RelType } from '../src/pptx/relTypes.js';
import { EditSession } from '../src/pptx/edit/session.js';

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

const NOTES_SLIDE_CT =
  'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';
const SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';

const slideXml = (label: string) =>
  `<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:p><a:r><a:t>${label}</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld></p:sld>`;

/**
 * Three slides. Slide 2 carries a notes slide (which must die with it) and is
 * referenced from both a custom show and a section list (which must be purged).
 */
function buildDeck(): Uint8Array {
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/ppt/slides/slide1.xml" ContentType="${SLIDE_CT}"/>
        <Override PartName="/ppt/slides/slide2.xml" ContentType="${SLIDE_CT}"/>
        <Override PartName="/ppt/slides/slide3.xml" ContentType="${SLIDE_CT}"/>
        <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="${NOTES_SLIDE_CT}"/>
      </Types>`,
    ),
    '_rels/.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.OfficeDocument}" Target="ppt/presentation.xml"/>`),
    ),
    'ppt/presentation.xml': strToU8(
      `<p:presentation xmlns:p="p" xmlns:r="r" xmlns:p14="p14">
        <p:sldIdLst>
          <p:sldId id="256" r:id="rId1"/>
          <p:sldId id="257" r:id="rId2"/>
          <p:sldId id="258" r:id="rId3"/>
        </p:sldIdLst>
        <p:sldSz cx="12192000" cy="6858000"/>
        <p:custShowLst>
          <p:custShow name="Short" id="0">
            <p:sldLst><p:sld r:id="rId1"/><p:sld r:id="rId2"/></p:sldLst>
          </p:custShow>
        </p:custShowLst>
        <p:extLst><p:ext uri="sections">
          <p14:sectionLst>
            <p14:section name="All"><p14:sldIdLst>
              <p14:sldId id="256"/><p14:sldId id="257"/><p14:sldId id="258"/>
            </p14:sldIdLst></p14:section>
          </p14:sectionLst>
        </p:ext></p:extLst>
      </p:presentation>`,
    ),
    'ppt/_rels/presentation.xml.rels': strToU8(
      rels(
        `<Relationship Id="rId1" Type="${RelType.Slide}" Target="slides/slide1.xml"/>` +
          `<Relationship Id="rId2" Type="${RelType.Slide}" Target="slides/slide2.xml"/>` +
          `<Relationship Id="rId3" Type="${RelType.Slide}" Target="slides/slide3.xml"/>`,
      ),
    ),
    'ppt/slideLayouts/slideLayout1.xml': strToU8(
      `<p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
    ),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': strToU8(rels('')),
    'ppt/notesSlides/notesSlide1.xml': strToU8(
      `<p:notes xmlns:p="p"><p:cSld><p:spTree/></p:cSld></p:notes>`,
    ),
    'ppt/notesSlides/_rels/notesSlide1.xml.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.Slide}" Target="../slides/slide2.xml"/>`),
    ),
  };
  for (const n of [1, 2, 3]) {
    files[`ppt/slides/slide${n}.xml`] = strToU8(slideXml(`Slide ${n}`));
    files[`ppt/slides/_rels/slide${n}.xml.rels`] = strToU8(
      rels(
        `<Relationship Id="rId1" Type="${RelType.SlideLayout}" Target="../slideLayouts/slideLayout1.xml"/>` +
          (n === 2
            ? `<Relationship Id="rId2" Type="${RelType.NotesSlide}" Target="../notesSlides/notesSlide1.xml"/>`
            : ''),
      ),
    );
  }
  return zipSync(files);
}

/** The visible text of a slide, for identifying which slide is which. */
function titleOf(deck: Deck, index: number): string {
  const shape = deck.slides[index]!.shapes[0] as { text?: { paragraphs: { runs: { text: string }[] }[] } };
  return shape.text!.paragraphs[0]!.runs.map((r) => r.text).join('');
}

describe('deleting a slide', () => {
  it('drops it from the running order and renumbers the rest', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);

    expect(session.deleteSlide(1)).toBe(true);
    expect(deck.slides.length).toBe(2);
    expect([titleOf(deck, 0), titleOf(deck, 1)]).toEqual(['Slide 1', 'Slide 3']);
    // `index` is the position in the deck, so it must follow the shift.
    expect(deck.slides.map((s) => s.index)).toEqual([0, 1]);
  });

  it('removes the slide part, its rels, its notes slide and their content types', () => {
    const deck = Deck.load(buildDeck());
    new EditSession(deck).deleteSlide(1);
    const out = unzipSync(deck.toBytes());

    expect(out['ppt/slides/slide2.xml']).toBeUndefined();
    expect(out['ppt/slides/_rels/slide2.xml.rels']).toBeUndefined();
    expect(out['ppt/notesSlides/notesSlide1.xml']).toBeUndefined();
    expect(out['ppt/notesSlides/_rels/notesSlide1.xml.rels']).toBeUndefined();

    const types = strFromU8(out['[Content_Types].xml']!);
    expect(types).not.toContain('/ppt/slides/slide2.xml');
    expect(types).not.toContain('/ppt/notesSlides/notesSlide1.xml');
    // The other slides keep theirs.
    expect(types).toContain('/ppt/slides/slide1.xml');
    expect(types).toContain('/ppt/slides/slide3.xml');
  });

  it('leaves no dangling reference in presentation.xml or its rels', () => {
    const deck = Deck.load(buildDeck());
    new EditSession(deck).deleteSlide(1);
    const out = unzipSync(deck.toBytes());

    const pres = strFromU8(out['ppt/presentation.xml']!);
    expect(pres).not.toContain('"rId2"');
    expect(pres).not.toContain('id="257"');
    // Slides either side, the custom show and the section list all survive.
    expect(pres).toContain('id="256"');
    expect(pres).toContain('id="258"');
    expect(pres).toContain('custShow');
    expect(pres).toContain('sectionLst');

    expect(strFromU8(out['ppt/_rels/presentation.xml.rels']!)).not.toContain('slide2.xml');
  });

  it('refuses to delete the last remaining slide', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    expect(session.deleteSlide(0)).toBe(true);
    expect(session.deleteSlide(0)).toBe(true);
    expect(session.canDeleteSlide(0)).toBe(false);
    expect(session.deleteSlide(0)).toBe(false);
    expect(deck.slides.length).toBe(1);
  });

  it('leaves untouched parts byte-identical', () => {
    const original = buildDeck();
    const deck = Deck.load(original);
    new EditSession(deck).deleteSlide(1);

    const before = unzipSync(original);
    const after = unzipSync(deck.toBytes());
    for (const name of ['ppt/slides/slide1.xml', 'ppt/slides/slide3.xml', '_rels/.rels']) {
      expect(after[name]).toEqual(before[name]);
    }
  });

  it('undoes and redoes exactly', () => {
    const original = buildDeck();
    const deck = Deck.load(original);
    const session = new EditSession(deck);
    session.deleteSlide(1);

    session.undo();
    expect(deck.slides.length).toBe(3);
    expect(titleOf(deck, 1)).toBe('Slide 2');
    const restored = unzipSync(deck.toBytes());
    const before = unzipSync(original);
    for (const name of Object.keys(before)) {
      // Rewritten parts are re-serialized, so compare their parse, not bytes.
      expect(restored[name], name).toBeDefined();
    }
    expect(strFromU8(restored['ppt/presentation.xml']!)).toContain('id="257"');
    expect(restored['ppt/notesSlides/notesSlide1.xml']).toEqual(
      before['ppt/notesSlides/notesSlide1.xml'],
    );

    session.redo();
    expect(deck.slides.length).toBe(2);
    expect(unzipSync(deck.toBytes())['ppt/slides/slide2.xml']).toBeUndefined();
  });

  it('reports no edits before a deletion and edits after', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    expect(session.hasEdits).toBe(false);
    session.deleteSlide(2);
    expect(session.hasEdits).toBe(true);
  });
});
