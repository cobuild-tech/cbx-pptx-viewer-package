import { describe, it, expect } from 'vitest';
import { parseXml, child } from '../src/oxml/xml.js';
import { blipEmbed } from '../src/pptx/shapes/fill.js';

const blipOf = (xml: string) => child(parseXml(xml)!, 'blip');

describe('blipEmbed', () => {
  it('reads a direct r:embed', () => {
    expect(blipEmbed(blipOf(`<a:blipFill xmlns:a="a" xmlns:r="r"><a:blip r:embed="rId7"/></a:blipFill>`))).toBe('rId7');
  });

  it('falls back to the svgBlip alternative when there is no raster embed', () => {
    const xml =
      `<a:blipFill xmlns:a="a" xmlns:r="r" xmlns:asvg="asvg">
         <a:blip><a:extLst><a:ext uri="{x}"><asvg:svgBlip r:embed="rId3"/></a:ext></a:extLst></a:blip>
       </a:blipFill>`;
    expect(blipEmbed(blipOf(xml))).toBe('rId3');
  });

  it('returns undefined when nothing is referenced', () => {
    expect(blipEmbed(blipOf(`<a:blipFill xmlns:a="a"><a:blip/></a:blipFill>`))).toBeUndefined();
  });
});
