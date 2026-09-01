/**
 * PPTX shape editing: addressing a shape's XML, writing geometry back into the
 * right element for each shape kind, deletion, z-order, and undo.
 *
 * No DOM here — the drag math is covered in pptx-shape-geometry, so these start
 * from a finished {@link Transform} and check what lands in the XML.
 */
import { describe, it, expect } from 'vitest';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { Deck } from '../src/pptx/deck/deck.js';
import { RelType } from '../src/pptx/relTypes.js';
import { EditSession } from '../src/pptx/edit/session.js';
import { child, children, attr, localName, serializeXml } from '../src/oxml/xml.js';

import type { Shape, Transform } from '../src/pptx/model.js';

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

/**
 * A deck covering every place a transform can live: a placeholder that states
 * none (inheriting the layout's), a plain shape, a group with its own child
 * coordinate space, a graphic frame, and a layout shape that is not the
 * slide's to move.
 */
function buildDeck(): Uint8Array {
  return zipSync({
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
          <p:nvSpPr><p:cNvPr id="9" name="LayoutDecoration"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="10" name="Title PH"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="476250"/><a:ext cx="1905000" cy="952500"/></a:xfrm></p:spPr>
        </p:sp>
      </p:spTree></p:cSld></p:sldLayout>`,
    ),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': strToU8(rels('')),
    'ppt/slides/slide1.xml': strToU8(
      `<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
        <p:nvGrpSpPr><p:cNvPr id="1" name="Shape Tree"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
          <p:spPr/>
          <p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="3" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm rot="2700000" flipH="1"><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
          </p:spPr>
        </p:sp>
        <p:grpSp>
          <p:nvGrpSpPr><p:cNvPr id="4" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr>
            <a:xfrm>
              <a:off x="0" y="0"/><a:ext cx="1905000" cy="1905000"/>
              <a:chOff x="0" y="0"/><a:chExt cx="3810000" cy="3810000"/>
            </a:xfrm>
          </p:grpSpPr>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="5" name="InGroup"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm></p:spPr>
          </p:sp>
        </p:grpSp>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="6" name="Frame"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="4762500" y="0"/><a:ext cx="952500" cy="952500"/></p:xfrm>
          <a:graphic><a:graphicData uri="urn:unknown"/></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld></p:sld>`,
    ),
  });
}

/** Shapes by their `<p:cNvPr name>`, which the parser now carries on the model. */
function shapeNamed(deck: Deck, name: string): Shape {
  const found = deck.slides[0]!.shapes.find((s) => s.name === name);
  if (!found) throw new Error(`no shape named ${name}`);
  return found;
}

function slideXml(deck: Deck): string {
  const root = deck.slideXml(0);
  return root ? serializeXml(root) : '';
}

const at = (t: Transform | undefined) => [t?.x, t?.y, t?.w, t?.h];

describe('pptx shapes — addressing', () => {
  it('records each shape against the XML node it was built from', () => {
    const deck = Deck.load(buildDeck());
    expect(deck.sourceOf(shapeNamed(deck, 'Box'))?.node.name).toBe('p:sp');
    expect(deck.sourceOf(shapeNamed(deck, 'Group'))?.node.name).toBe('p:grpSp');
    expect(deck.sourceOf(shapeNamed(deck, 'Frame'))?.node.name).toBe('p:graphicFrame');
  });

  it('carries the non-visual id and name onto the model', () => {
    const deck = Deck.load(buildDeck());
    expect(shapeNamed(deck, 'Box').id).toBe('3');
    expect(shapeNamed(deck, 'Title').id).toBe('2');
  });

  it('treats a shape composited in from the layout as not the slide’s to move', () => {
    const deck = Deck.load(buildDeck());
    expect(deck.isEditable(0, shapeNamed(deck, 'LayoutDecoration'))).toBe(false);
    expect(deck.isEditable(0, shapeNamed(deck, 'Box'))).toBe(true);
  });
});

describe('pptx shapes — geometry write-back', () => {
  it('moves a shape and survives export and reload', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const box = shapeNamed(deck, 'Box');

    session.commitShapeTransform(0, box, { ...box.transform!, x: 200, y: 100 });
    expect(at(shapeNamed(deck, 'Box').transform)).toEqual([200, 100, 200, 100]);

    const reloaded = Deck.load(deck.toBytes());
    expect(at(shapeNamed(reloaded, 'Box').transform)).toEqual([200, 100, 200, 100]);
  });

  it('keeps rotation and mirroring, and drops them when they go back to zero', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const box = shapeNamed(deck, 'Box');
    expect(box.transform?.rot).toBe(45);
    expect(box.transform?.flipH).toBe(true);

    session.commitShapeTransform(0, box, { ...box.transform!, rot: 90 });
    const xfrm = child(child(deck.sourceOf(shapeNamed(deck, 'Box'))!.node, 'spPr'), 'xfrm')!;
    expect(attr(xfrm, 'rot')).toBe('5400000');
    expect(attr(xfrm, 'flipH')).toBe('1');

    const rotated = shapeNamed(deck, 'Box');
    session.commitShapeTransform(0, rotated, { ...rotated.transform!, rot: 0, flipH: false });
    const after = child(child(deck.sourceOf(shapeNamed(deck, 'Box'))!.node, 'spPr'), 'xfrm')!;
    expect(attr(after, 'rot')).toBeUndefined();
    expect(attr(after, 'flipH')).toBeUndefined();
  });

  it('gives an inheriting placeholder an xfrm of its own, at the head of spPr', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const title = shapeNamed(deck, 'Title');
    // Position came from the layout placeholder, with nothing in the slide.
    expect(at(title.transform)).toEqual([100, 50, 200, 100]);

    session.commitShapeTransform(0, title, { ...title.transform!, x: 300 });

    const spPr = child(deck.sourceOf(shapeNamed(deck, 'Title'))!.node, 'spPr')!;
    // The schema sequences <a:xfrm> before geometry and fill, so it has to be
    // created first, not appended.
    expect(localName(spPr.children[0]!.name)).toBe('xfrm');
    expect(at(shapeNamed(deck, 'Title').transform)).toEqual([300, 50, 200, 100]);
    // The layout is untouched — the point of writing into the slide.
    expect(deck.slides[0]!.shapes.find((s) => s.name === 'Title PH')).toBeUndefined();
  });

  it('writes a graphic frame’s xfrm as its own child, not into spPr', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const frame = shapeNamed(deck, 'Frame');
    session.commitShapeTransform(0, frame, { ...frame.transform!, x: 10, y: 20 });

    const node = deck.sourceOf(shapeNamed(deck, 'Frame'))!.node;
    expect(node.children.map((c) => localName(c.name))).toEqual([
      'nvGraphicFramePr',
      'xfrm',
      'graphic',
    ]);
    expect(child(node, 'spPr')).toBeUndefined();
    expect(at(shapeNamed(deck, 'Frame').transform)).toEqual([10, 20, 100, 100]);
  });

  it('resizes a group without touching its child coordinate space', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const group = shapeNamed(deck, 'Group');
    session.commitShapeTransform(0, group, { ...group.transform!, w: 400, h: 400 });

    const xfrm = child(child(deck.sourceOf(shapeNamed(deck, 'Group'))!.node, 'grpSpPr'), 'xfrm')!;
    // chOff/chExt untouched is what makes the children scale with the group.
    expect(attr(child(xfrm, 'chExt'), 'cx')).toBe('3810000');
    expect(attr(child(xfrm, 'ext'), 'cx')).toBe(String(400 * 9525));
    expect(xfrm.children.map((c) => localName(c.name))).toEqual(['off', 'ext', 'chOff', 'chExt']);
  });

  it('refuses to move a shape inherited from the layout', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const before = slideXml(deck);
    const decoration = shapeNamed(deck, 'LayoutDecoration');

    expect(session.commitShapeTransform(0, decoration, { x: 9, y: 9, w: 9, h: 9 })).toBeUndefined();
    expect(slideXml(deck)).toBe(before);
    expect(deck.hasEdits).toBe(false);
  });

  it('rounds to whole EMU without drifting over repeated moves', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    for (let i = 0; i < 50; i++) {
      const box = shapeNamed(deck, 'Box');
      session.commitShapeTransform(0, box, { ...box.transform!, x: box.transform!.x + 1 });
    }
    // 100 + 50 exactly: each commit re-reads the EMU the last one wrote, so the
    // rounding error cannot accumulate.
    expect(shapeNamed(deck, 'Box').transform!.x).toBeCloseTo(150, 3);
  });
});

describe('pptx shapes — deletion and z-order', () => {
  it('deletes a shape and leaves the rest of the slide alone', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    session.deleteShapes(0, [shapeNamed(deck, 'Box')]);

    const names = deck.slides[0]!.shapes.map((s) => s.name);
    expect(names).not.toContain('Box');
    expect(names).toEqual(expect.arrayContaining(['Title', 'Group', 'Frame']));

    // The shape tree's own properties are not shapes and must survive.
    const spTree = child(child(deck.slideXml(0)!, 'cSld'), 'spTree')!;
    expect(children(spTree, 'nvGrpSpPr')).toHaveLength(1);
    expect(children(spTree, 'grpSpPr')).toHaveLength(1);
  });

  it('deletes several shapes as one undoable change', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    session.deleteShapes(0, [shapeNamed(deck, 'Box'), shapeNamed(deck, 'Frame')]);
    expect(deck.slides[0]!.shapes.map((s) => s.name)).not.toContain('Frame');

    session.undo();
    const names = deck.slides[0]!.shapes.map((s) => s.name);
    expect(names).toEqual(expect.arrayContaining(['Box', 'Frame']));
    expect(session.canUndo).toBe(false);
  });

  it('moves a shape through the z-order without disturbing the tree header', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const order = () =>
      deck.slides[0]!.shapes.filter((s) => s.name !== 'LayoutDecoration').map((s) => s.name);
    expect(order()).toEqual(['Title', 'Box', 'Group', 'Frame']);

    session.reorderShape(0, shapeNamed(deck, 'Box'), 'front');
    expect(order()).toEqual(['Title', 'Group', 'Frame', 'Box']);

    session.reorderShape(0, shapeNamed(deck, 'Box'), 'backward');
    expect(order()).toEqual(['Title', 'Group', 'Box', 'Frame']);

    const spTree = child(child(deck.slideXml(0)!, 'cSld'), 'spTree')!;
    expect(spTree.children.map((c) => localName(c.name))).toEqual([
      'nvGrpSpPr',
      'grpSpPr',
      'sp',
      'grpSp',
      'sp',
      'graphicFrame',
    ]);
  });

  it('reports a no-op restack without leaving an undo entry', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    expect(session.reorderShape(0, shapeNamed(deck, 'Title'), 'backward')).toBeUndefined();
    expect(session.canUndo).toBe(false);
    expect(deck.hasEdits).toBe(false);
  });
});

describe('pptx shapes — undo and export', () => {
  it('undoes and redoes a move', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const box = shapeNamed(deck, 'Box');
    session.commitShapeTransform(0, box, { ...box.transform!, x: 500 });
    expect(shapeNamed(deck, 'Box').transform!.x).toBe(500);

    session.undo();
    expect(shapeNamed(deck, 'Box').transform!.x).toBe(100);
    session.redo();
    expect(shapeNamed(deck, 'Box').transform!.x).toBe(500);
  });

  it('collapses a multi-shape drag into a single undo step', () => {
    const deck = Deck.load(buildDeck());
    const session = new EditSession(deck);
    const box = shapeNamed(deck, 'Box');
    const frame = shapeNamed(deck, 'Frame');
    session.commitShapeTransforms(0, [
      { shape: box, transform: { ...box.transform!, x: box.transform!.x + 10 } },
      { shape: frame, transform: { ...frame.transform!, x: frame.transform!.x + 10 } },
    ]);
    expect(shapeNamed(deck, 'Box').transform!.x).toBe(110);
    expect(shapeNamed(deck, 'Frame').transform!.x).toBe(510);

    session.undo();
    expect(shapeNamed(deck, 'Box').transform!.x).toBe(100);
    expect(shapeNamed(deck, 'Frame').transform!.x).toBe(500);
    expect(session.canUndo).toBe(false);
  });

  it('re-emits untouched parts byte-for-byte', () => {
    const original = buildDeck();
    const deck = Deck.load(original);
    const session = new EditSession(deck);
    const box = shapeNamed(deck, 'Box');
    session.commitShapeTransform(0, box, { ...box.transform!, x: 42 });

    const before = unzipSync(original);
    const after = unzipSync(deck.toBytes());
    for (const [name, bytes] of Object.entries(before)) {
      if (name === 'ppt/slides/slide1.xml') continue;
      expect(strFromU8(after[name]!), name).toBe(strFromU8(bytes));
    }
    // …and the part that did change says exactly what was written.
    expect(shapeNamed(Deck.load(deck.toBytes()), 'Box').transform!.x).toBe(42);
  });
});
