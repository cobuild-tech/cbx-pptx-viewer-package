import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { Deck } from '../src/parse/deck.js';
import { RelType } from '../src/opc/relTypes.js';
import type { PresetShape } from '../src/model.js';

/** A structurally complete (if minimal) single-slide deck for pipeline tests. */
function buildDeck(): Uint8Array {
  const rels = (entries: string) =>
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

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
        <p:embeddedFontLst>
          <p:embeddedFont>
            <p:font typeface="BrandFont"/>
            <p:regular r:id="rIdF1"/>
            <p:bold r:id="rIdF2"/>
          </p:embeddedFont>
        </p:embeddedFontLst>
      </p:presentation>`,
    ),
    'ppt/_rels/presentation.xml.rels': strToU8(
      rels(
        `<Relationship Id="rId1" Type="${RelType.Slide}" Target="slides/slide1.xml"/>` +
          `<Relationship Id="rIdF1" Type="${RelType.Font}" Target="fonts/font1.fntdata"/>` +
          `<Relationship Id="rIdF2" Type="${RelType.Font}" Target="fonts/font2.fntdata"/>`,
      ),
    ),
    'ppt/fonts/font1.fntdata': new Uint8Array([0x00, 0x01, 0x00, 0x00]),
    'ppt/fonts/font2.fntdata': new Uint8Array([0x00, 0x01, 0x00, 0x00]),
    'ppt/slides/slide1.xml': strToU8(
      `<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
          <p:spPr/>
          <p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="3" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
          </p:spPr>
          <p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1800" b="1"/><a:t>Box</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree></p:cSld></p:sld>`,
    ),
    'ppt/slides/_rels/slide1.xml.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.SlideLayout}" Target="../slideLayouts/slideLayout1.xml"/>`),
    ),
    'ppt/slideLayouts/slideLayout1.xml': strToU8(
      `<p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="11277600" cy="1325563"/></a:xfrm></p:spPr>
        </p:sp>
      </p:spTree></p:cSld></p:sldLayout>`,
    ),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.SlideMaster}" Target="../slideMasters/slideMaster1.xml"/>`),
    ),
    'ppt/slideMasters/slideMaster1.xml': strToU8(
      `<p:sldMaster xmlns:p="p" xmlns:a="a">
        <p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></p:bgPr></p:bg>
          <p:spTree>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="9" name="MasterLogo"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                <a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr>
            </p:sp>
          </p:spTree>
        </p:cSld>
        <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"/>
        <p:txStyles>
          <p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:titleStyle>
        </p:txStyles>
      </p:sldMaster>`,
    ),
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.Theme}" Target="../theme/theme1.xml"/>`),
    ),
    'ppt/theme/theme1.xml': strToU8(
      `<a:theme xmlns:a="a"><a:themeElements>
        <a:clrScheme name="t">
          <a:dk1><a:srgbClr val="111111"/></a:dk1>
          <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
          <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
        </a:clrScheme>
        <a:fontScheme name="f">
          <a:majorFont><a:latin typeface="Georgia"/></a:majorFont>
          <a:minorFont><a:latin typeface="Verdana"/></a:minorFont>
        </a:fontScheme>
      </a:themeElements></a:theme>`,
    ),
  };
  return zipSync(files);
}

describe('Deck pipeline', () => {
  const deck = Deck.load(buildDeck());

  it('reads slide size in px (16:9 at 96dpi)', () => {
    expect(Math.round(deck.size.wPx)).toBe(1280);
    expect(Math.round(deck.size.hPx)).toBe(720);
  });

  it('loads exactly one slide', () => {
    expect(deck.slides).toHaveLength(1);
  });

  it('parses embedded fonts with their variants and resolved parts', () => {
    expect(deck.embeddedFonts).toHaveLength(1);
    const font = deck.embeddedFonts[0]!;
    expect(font.typeface).toBe('BrandFont');
    expect(font.faces).toEqual([
      { weight: 400, style: 'normal', part: 'ppt/fonts/font1.fntdata' },
      { weight: 700, style: 'normal', part: 'ppt/fonts/font2.fntdata' },
    ]);
    expect(deck.fontBytes('ppt/fonts/font1.fntdata')).toBeInstanceOf(Uint8Array);
  });

  it('resolves the slide background through the master scheme color', () => {
    const bg = deck.slides[0]!.background;
    expect(bg.type).toBe('solid');
    if (bg.type === 'solid') expect(bg.color.hex).toBe('FFFFFF'); // bg1 -> lt1 -> white
  });

  const shapes = () => deck.slides[0]!.shapes as PresetShape[];
  const title = () => shapes().find((s) => s.placeholder?.type === 'title')!;

  it('composites master/layout decorations beneath the slide shapes', () => {
    // The master's non-placeholder logo renders first (lowest z-order).
    expect(shapes()[0]!.fill).toEqual({ type: 'solid', color: { hex: '4472C4' } });
    // Slide content (title + box) comes after the inherited decorations.
    expect(shapes().some((s) => s.placeholder?.type === 'title')).toBe(true);
  });

  it('inherits the title placeholder geometry from the layout', () => {
    const t = title();
    expect(t.kind).toBe('shape');
    // Layout off x=457200 EMU = 48px, ext cx=11277600 EMU = 1184px.
    expect(Math.round(t.transform!.x)).toBe(48);
    expect(Math.round(t.transform!.w)).toBe(1184);
  });

  it('applies the master title style to the title text (size + color + align)', () => {
    const run = title().text!.paragraphs[0]!.runs[0]!;
    expect(run.sizePt).toBe(44); // sz=4400
    expect(run.color?.hex).toBe('111111'); // tx1 -> dk1 -> 111111
    expect(title().text!.paragraphs[0]!.align).toBe('ctr');
  });

  it('parses an explicit solid fill and run formatting on a non-placeholder shape', () => {
    const box = shapes().find((s) => !s.placeholder && s.fill.type === 'solid' && s.fill.color.hex === 'FF0000')!;
    expect(box.fill).toEqual({ type: 'solid', color: { hex: 'FF0000' } });
    const run = box.text!.paragraphs[0]!.runs[0]!;
    expect(run.sizePt).toBe(18);
    expect(run.bold).toBe(true);
  });
});
