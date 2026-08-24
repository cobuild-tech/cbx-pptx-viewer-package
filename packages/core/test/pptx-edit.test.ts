/**
 * PPTX text editing: source addressing, XML write-back, export round-trip.
 *
 * These exercise the edit pipeline without a DOM — the reconciler's output
 * (a ParaEdit list) is constructed directly, so the XML and packaging halves
 * are tested independently of contentEditable.
 */
import { describe, it, expect } from 'vitest';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { Deck } from '../src/pptx/deck/deck.js';
import { RelType } from '../src/pptx/relTypes.js';
import { EditSession } from '../src/pptx/edit/session.js';
import type { ParaEdit } from '../src/pptx/edit/xmlWrite.js';
import type { PresetShape, TextBody } from '../src/pptx/model.js';

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

/**
 * A deck whose slide has a two-run title (so we can check a sibling run's
 * formatting survives an edit) and a bulleted body, plus a layout that
 * contributes its own non-placeholder text — the read-only case.
 */
function buildDeck(): Uint8Array {
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
      </Types>`,
    ),
    '_rels/.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.OfficeDocument}" Target="ppt/presentation.xml"/>`),
    ),
    'ppt/presentation.xml': strToU8(
      `<p:presentation xmlns:p="p" xmlns:r="r">
        <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
        <p:sldSz cx="12192000" cy="6858000"/>
      </p:presentation>`,
    ),
    'ppt/_rels/presentation.xml.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.Slide}" Target="slides/slide1.xml"/>`),
    ),
    'ppt/slides/_rels/slide1.xml.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.SlideLayout}" Target="../slideLayouts/slideLayout1.xml"/>`),
    ),
    'ppt/slideLayouts/slideLayout1.xml': strToU8(
      `<p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="9" name="LayoutNote"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/><a:p><a:r><a:t>From layout</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld></p:sldLayout>`,
    ),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': strToU8(rels('')),
    'ppt/slides/slide1.xml': strToU8(
      `<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
          <p:spPr/>
          <p:txBody><a:bodyPr/>
            <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="4000"/><a:t>Hello </a:t></a:r><a:r><a:rPr lang="en-US" sz="4000" b="1" i="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>World</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/>
            <a:p><a:pPr lvl="1" marL="457200"><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>One</a:t></a:r></a:p>
            <a:p><a:endParaRPr sz="1800"/></a:p>
          </p:txBody>
        </p:sp>
      </p:spTree></p:cSld></p:sld>`,
    ),
  };
  return zipSync(files);
}

/** The text body of the shape named by its 1-based position in the slide tree. */
function bodyAt(deck: Deck, shapeIndex: number): TextBody {
  const shape = deck.slides[0]!.shapes[shapeIndex] as PresetShape;
  return shape.text!;
}

/** Flatten a text body back to a plain string, paragraphs joined by \n. */
function textOf(body: TextBody): string {
  return body.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
}

/** A ParaEdit list that keeps every paragraph and run exactly as parsed. */
function identityEdit(body: TextBody): ParaEdit[] {
  return body.paragraphs.map((p) => ({
    src: p,
    segments: p.runs.map((r) => ({ text: r.text, src: r, isBreak: r.text === '\n' })),
  }));
}

describe('pptx edit — addressing', () => {
  it('records the XML source of bodies, paragraphs and runs', () => {
    const deck = Deck.load(buildDeck());
    const body = bodyAt(deck, 1); // index 0 is the layout shape
    expect(deck.sourceOf(body)?.node.name).toBe('p:txBody');
    expect(deck.sourceOf(body.paragraphs[0]!)?.node.name).toBe('a:p');
    expect(deck.sourceOf(body.paragraphs[0]!.runs[0]!)?.node.name).toBe('a:r');
  });

  it('marks slide-owned text editable and layout-inherited text read-only', () => {
    const deck = Deck.load(buildDeck());
    const layoutBody = bodyAt(deck, 0);
    const slideBody = bodyAt(deck, 1);
    expect(textOf(layoutBody)).toBe('From layout');
    expect(deck.isEditable(0, layoutBody)).toBe(false);
    expect(deck.isEditable(0, slideBody)).toBe(true);
  });
});

describe('pptx edit — round trip', () => {
  it('writes edited text back and survives export/reload', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const body = bodyAt(deck, 1);

    const paras = identityEdit(body);
    paras[0]!.segments[0]!.text = 'Goodbye ';
    expect(session.commitTextBody(0, body, paras)).toBeTruthy();

    const reloaded = Deck.load(session.exportBytes());
    const after = bodyAt(reloaded, 1);
    expect(textOf(after)).toBe('Goodbye World');
  });

  it('leaves the untouched sibling run’s formatting intact', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const body = bodyAt(deck, 1);

    const paras = identityEdit(body);
    paras[0]!.segments[0]!.text = 'Goodbye ';
    session.commitTextBody(0, body, paras);

    const after = bodyAt(Deck.load(session.exportBytes()), 1);
    const world = after.paragraphs[0]!.runs[1]!;
    expect(world.text).toBe('World');
    expect(world.bold).toBe(true);
    expect(world.italic).toBe(true);
    expect(world.color?.hex).toBe('FF0000');
    expect(world.sizePt).toBe(40);
  });

  it('rewrites only the edited part, leaving every other part byte-identical', () => {
    const original = buildDeck();
    const deck = Deck.load(original);
    const session = new EditSession(deck);
    const body = bodyAt(deck, 1);

    const paras = identityEdit(body);
    paras[0]!.segments[0]!.text = 'Changed ';
    session.commitTextBody(0, body, paras);

    const before = unzipSync(original);
    const after = unzipSync(session.exportBytes());
    expect(Object.keys(after).sort()).toEqual(Object.keys(before).sort());
    for (const name of Object.keys(before)) {
      if (name === 'ppt/slides/slide1.xml') continue;
      expect(after[name], `${name} should be untouched`).toEqual(before[name]);
    }
  });

  it('preserves paragraph properties and the endParaRPr of an empty paragraph', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const body = bodyAt(deck, 2);

    const paras = identityEdit(body);
    paras[0]!.segments[0]!.text = 'Uno';
    session.commitTextBody(0, body, paras);

    const xml = strFromU8(unzipSync(session.exportBytes())['ppt/slides/slide1.xml']!);
    expect(xml).toContain('<a:pPr lvl="1" marL="457200"><a:buChar char="•"/></a:pPr>');
    expect(xml).toContain('<a:endParaRPr sz="1800"/>');

    const after = bodyAt(Deck.load(session.exportBytes()), 2);
    expect(after.paragraphs[0]!.level).toBe(1);
    expect(after.paragraphs[0]!.bullet).toEqual({ type: 'char', char: '•' });
    expect(after.paragraphs).toHaveLength(2);
  });

  it('refuses to edit layout-inherited text', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const layoutBody = bodyAt(deck, 0);
    const paras = identityEdit(layoutBody);
    paras[0]!.segments[0]!.text = 'Tampered';

    expect(session.commitTextBody(0, layoutBody, paras)).toBeUndefined();
    expect(textOf(bodyAt(Deck.load(session.exportBytes()), 0))).toBe('From layout');
  });
});

describe('pptx edit — structure and formatting', () => {
  it('splits a paragraph, cloning the source paragraph properties', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const body = bodyAt(deck, 2);
    const para = body.paragraphs[0]!;
    const run = para.runs[0]!;

    // "One" -> "On" / "e" as two paragraphs. A commit replaces the whole body,
    // so the trailing empty paragraph is passed through unchanged.
    session.commitTextBody(0, body, [
      { src: para, segments: [{ text: 'On', src: run }] },
      { src: para, segments: [{ text: 'e', src: run }] },
      { src: body.paragraphs[1]!, segments: [] },
    ]);

    const after = bodyAt(Deck.load(session.exportBytes()), 2);
    expect(after.paragraphs.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['On', 'e', '']);
    // The split-off paragraph keeps the bullet and indent level.
    expect(after.paragraphs[1]!.level).toBe(1);
    expect(after.paragraphs[1]!.bullet).toEqual({ type: 'char', char: '•' });
    expect(after.paragraphs[1]!.runs[0]!.sizePt).toBe(18);
  });

  it('merges two paragraphs into one', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const body = bodyAt(deck, 1);
    const para = body.paragraphs[0]!;

    session.commitTextBody(0, body, [
      { src: para, segments: para.runs.map((r) => ({ text: r.text, src: r })) },
    ]);
    const after = bodyAt(Deck.load(session.exportBytes()), 1);
    expect(after.paragraphs).toHaveLength(1);
    expect(textOf(after)).toBe('Hello World');
  });

  it('applies a formatting override as a new rPr on a split run', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const body = bodyAt(deck, 1);
    const para = body.paragraphs[0]!;
    const [hello, world] = [para.runs[0]!, para.runs[1]!];

    // Bold + recolour just "Hel", leaving "lo " as it was.
    session.commitTextBody(0, body, [
      {
        src: para,
        segments: [
          { text: 'Hel', src: hello, format: { bold: true, colorHex: '00FF00', sizePt: 20 } },
          { text: 'lo ', src: hello },
          { text: 'World', src: world },
        ],
      },
    ]);

    const after = bodyAt(Deck.load(session.exportBytes()), 1);
    const runs = after.paragraphs[0]!.runs;
    expect(runs.map((r) => r.text)).toEqual(['Hel', 'lo ', 'World']);
    expect(runs[0]!.bold).toBe(true);
    expect(runs[0]!.color?.hex).toBe('00FF00');
    expect(runs[0]!.sizePt).toBe(20);
    // The untouched remainder keeps the original size and no bold.
    expect(runs[1]!.sizePt).toBe(40);
    expect(runs[1]!.bold).toBeUndefined();
  });

  it('emits rPr before the text and orders solidFill correctly', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const body = bodyAt(deck, 2);
    const para = body.paragraphs[0]!;

    session.commitTextBody(0, body, [
      {
        src: para,
        segments: [{ text: 'One', src: para.runs[0]!, format: { colorHex: '123456', font: 'Arial' } }],
      },
    ]);

    const xml = strFromU8(unzipSync(session.exportBytes())['ppt/slides/slide1.xml']!);
    expect(xml).toContain(
      '<a:rPr sz="1800"><a:solidFill><a:srgbClr val="123456"/></a:solidFill>' +
        '<a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>One</a:t>',
    );
  });

  it('adds xml:space="preserve" only when whitespace would be lost', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const body = bodyAt(deck, 2);
    const para = body.paragraphs[0]!;

    session.commitTextBody(0, body, [
      { src: para, segments: [{ text: '  padded  ', src: para.runs[0]! }] },
    ]);

    const xml = strFromU8(unzipSync(session.exportBytes())['ppt/slides/slide1.xml']!);
    expect(xml).toContain('<a:t xml:space="preserve">  padded  </a:t>');
    expect(textOf(bodyAt(Deck.load(session.exportBytes()), 2))).toContain('  padded  ');
  });
});

describe('pptx edit — history', () => {
  it('undoes and redoes an edit', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const body = bodyAt(deck, 1);

    const paras = identityEdit(body);
    paras[0]!.segments[0]!.text = 'Goodbye ';
    session.commitTextBody(0, body, paras);
    expect(textOf(bodyAt(deck, 1))).toBe('Goodbye World');

    expect(session.canUndo).toBe(true);
    session.undo();
    expect(textOf(bodyAt(deck, 1))).toBe('Hello World');

    expect(session.canRedo).toBe(true);
    session.redo();
    expect(textOf(bodyAt(deck, 1))).toBe('Goodbye World');
  });

  it('restores the original XML exactly on undo', () => {
    const original = strFromU8(unzipSync(buildDeck())['ppt/slides/slide1.xml']!);
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const body = bodyAt(deck, 1);

    const paras = identityEdit(body);
    paras[0]!.segments[0]!.text = 'Goodbye ';
    session.commitTextBody(0, body, paras);
    session.undo();

    const restored = strFromU8(unzipSync(session.exportBytes())['ppt/slides/slide1.xml']!);
    // The snapshot is a re-serialization of the original parse, so compare the
    // parsed text rather than raw bytes (the declaration/whitespace normalizes).
    expect(restored).toContain('<a:t>Hello </a:t>');
    expect(restored).not.toContain('Goodbye');
    expect(original).toContain('<a:t>Hello </a:t>');
  });

  it('drops the redo branch after a fresh edit', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);

    const first = identityEdit(bodyAt(deck, 1));
    first[0]!.segments[0]!.text = 'A ';
    session.commitTextBody(0, bodyAt(deck, 1), first);
    session.undo();
    expect(session.canRedo).toBe(true);

    const second = identityEdit(bodyAt(deck, 1));
    second[0]!.segments[0]!.text = 'B ';
    session.commitTextBody(0, bodyAt(deck, 1), second);
    expect(session.canRedo).toBe(false);
    expect(textOf(bodyAt(deck, 1))).toBe('B World');
  });
});
