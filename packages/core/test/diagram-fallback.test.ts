import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { Deck } from '../src/pptx/deck/deck.js';
import { RelType } from '../src/pptx/relTypes.js';
import type { FrameShape, PresetShape } from '../src/pptx/model.js';

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

const pt = (id: string, text: string, type?: string) =>
  `<dgm:pt modelId="${id}"${type ? ` type="${type}"` : ''}>` +
  `<dgm:prSet${type === 'doc' ? ' loCatId="process"' : ''}/>` +
  `<dgm:t><a:bodyPr/><a:p>${text ? `<a:r><a:t>${text}</a:t></a:r>` : ''}</a:p></dgm:t></dgm:pt>`;

const cxn = (src: string, dest: string, ord: number) =>
  `<dgm:cxn srcId="${src}" destId="${dest}" srcOrd="${ord}"/>`;

/** Deck with a diagram frame whose data model has NO cached drawing -> fallback. */
function buildDeck(): Uint8Array {
  const data =
    `<dgm:dataModel xmlns:dgm="dgm" xmlns:a="a"><dgm:ptLst>` +
    pt('D', '', 'doc') +
    pt('A', 'Plan') +
    pt('B', 'Build') +
    pt('C', 'Ship') +
    `</dgm:ptLst><dgm:cxnLst>` +
    cxn('D', 'A', 0) + cxn('D', 'B', 1) + cxn('D', 'C', 2) +
    `</dgm:cxnLst></dgm:dataModel>`;
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/></Types>`,
    ),
    '_rels/.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.OfficeDocument}" Target="ppt/presentation.xml"/>`),
    ),
    'ppt/presentation.xml': strToU8(
      `<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
        <p:sldSz cx="12192000" cy="6858000"/></p:presentation>`,
    ),
    'ppt/_rels/presentation.xml.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.Slide}" Target="slides/slide1.xml"/>`),
    ),
    'ppt/slides/slide1.xml': strToU8(
      `<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="5" name="Diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="3000000"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
            <dgm:relIds xmlns:dgm="dgm" r:dm="rIdData"/>
          </a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld></p:sld>`,
    ),
    // Note: only a diagramData rel, NO diagramDrawing -> forces the fallback.
    'ppt/slides/_rels/slide1.xml.rels': strToU8(
      rels(`<Relationship Id="rIdData" Type="${RelType.Diagram}" Target="../diagrams/data1.xml"/>`),
    ),
    'ppt/diagrams/data1.xml': strToU8(data),
  };
  return zipSync(files);
}

describe('SmartArt data-model fallback', () => {
  const deck = Deck.load(buildDeck());
  const frame = deck.slides[0]!.shapes.find((s) => s.kind === 'frame') as FrameShape;

  it('lays out data-model nodes when no cached drawing exists', () => {
    expect(frame.frameType).toBe('diagram');
    expect(frame.diagram).toBeDefined();
    const boxes = frame.diagram!.filter((s) => s.kind === 'shape') as PresetShape[];
    expect(boxes).toHaveLength(3);
    const texts = boxes.map((b) => b.text?.paragraphs[0]?.runs[0]?.text);
    expect(texts).toEqual(['Plan', 'Build', 'Ship']);
  });

  it('arranges a process diagram as a horizontal sequence with arrows', () => {
    const boxes = frame.diagram!.filter((s) => s.kind === 'shape');
    const arrows = frame.diagram!.filter((s) => s.kind === 'connector');
    // process layout: boxes laid left-to-right, arrows between consecutive boxes
    expect(arrows.length).toBe(2);
    const xs = boxes.map((b) => b.transform!.x);
    expect(xs[0]!).toBeLessThan(xs[1]!);
    expect(xs[1]!).toBeLessThan(xs[2]!);
  });
});
