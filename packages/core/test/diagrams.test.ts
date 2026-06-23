import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { Deck } from '../src/pptx/deck/deck.js';
import { RelType } from '../src/pptx/relTypes.js';
import type { FrameShape, PresetShape } from '../src/pptx/model.js';

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

/** Minimal deck: one slide with a SmartArt graphicFrame + its drawing part. */
function buildDiagramDeck(): Uint8Array {
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
      `<p:presentation xmlns:p="p" xmlns:r="r">
        <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
        <p:sldSz cx="12192000" cy="6858000"/></p:presentation>`,
    ),
    'ppt/_rels/presentation.xml.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.Slide}" Target="slides/slide1.xml"/>`),
    ),
    'ppt/slides/slide1.xml': strToU8(
      `<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="5" name="Diagram 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="1000000" y="500000"/><a:ext cx="3000000" cy="2000000"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
            <dgm:relIds xmlns:dgm="dgm" r:dm="rIdData" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/>
          </a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld></p:sld>`,
    ),
    'ppt/slides/_rels/slide1.xml.rels': strToU8(
      rels(
        `<Relationship Id="rIdData" Type="${RelType.Diagram}" Target="../diagrams/data1.xml"/>` +
          `<Relationship Id="rIdDraw" Type="${RelType.DiagramDrawing}" Target="../diagrams/drawing1.xml"/>`,
      ),
    ),
    'ppt/diagrams/drawing1.xml': strToU8(
      `<dsp:drawing xmlns:dsp="dsp" xmlns:a="a"><dsp:spTree>
        <dsp:nvGrpSpPr/><dsp:grpSpPr/>
        <dsp:sp>
          <dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>
          <dsp:spPr>
            <a:xfrm><a:off x="100000" y="200000"/><a:ext cx="500000" cy="400000"/></a:xfrm>
            <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
          </dsp:spPr>
          <dsp:txBody><a:bodyPr/><a:p><a:r><a:t>Node</a:t></a:r></a:p></dsp:txBody>
        </dsp:sp>
      </dsp:spTree></dsp:drawing>`,
    ),
  };
  return zipSync(files);
}

describe('SmartArt diagram resolution', () => {
  const deck = Deck.load(buildDiagramDeck());
  const frame = deck.slides[0]!.shapes.find((s) => s.kind === 'frame') as FrameShape;

  it('produces a diagram frame at the graphicFrame transform', () => {
    expect(frame.frameType).toBe('diagram');
    expect(Math.round(frame.transform!.x)).toBe(Math.round(1000000 / 9525));
  });

  it('resolves the pre-laid-out drawing into renderable shapes', () => {
    expect(frame.diagram).toBeDefined();
    expect(frame.diagram).toHaveLength(1);
    const node = frame.diagram![0] as PresetShape;
    expect(node.kind).toBe('shape');
    expect(node.geom).toMatchObject({ type: 'preset', preset: 'ellipse' });
    expect(node.fill).toEqual({ type: 'solid', color: { hex: '00FF00' } });
    // Drawing-canvas coordinates are frame-relative (off x=100000 EMU ~= 10.5px).
    expect(Math.round(node.transform!.x)).toBe(Math.round(100000 / 9525));
    expect(node.text!.paragraphs[0]!.runs[0]!.text).toBe('Node');
  });
});
