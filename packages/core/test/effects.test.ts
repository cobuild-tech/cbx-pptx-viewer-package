import { describe, it, expect } from 'vitest';
import { parseXml } from '../src/oxml/xml.js';
import { parseEffectLst } from '../src/pptx/effects/effects.js';
import { parseTheme, type ColorContext } from '../src/pptx/color.js';

const ctx: ColorContext = { theme: parseTheme(undefined), clrMap: {} };

function effects(xml: string) {
  const node = parseXml(xml);
  if (!node) throw new Error('parse failed');
  return parseEffectLst(node, ctx);
}

describe('parseEffectLst', () => {
  it('parses an outer shadow with offset, blur and alpha', () => {
    const [e] = effects(
      `<a:effectLst xmlns:a="a">
         <a:outerShdw blurRad="50800" dist="38100" dir="2700000">
           <a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr>
         </a:outerShdw>
       </a:effectLst>`,
    );
    expect(e?.type).toBe('outerShadow');
    if (e?.type !== 'outerShadow') return;
    expect(e.blur).toBeCloseTo(50800 / 9525, 2);
    // dist 4px at 45deg -> equal dx/dy, both positive (right + down)
    expect(e.dx).toBeCloseTo(2.83, 1);
    expect(e.dy).toBeCloseTo(2.83, 1);
    expect(e.color.hex).toBe('000000');
    expect(e.color.alpha).toBeCloseTo(0.4, 2);
  });

  it('parses glow and soft edge, preserving document order', () => {
    const list = effects(
      `<a:effectLst xmlns:a="a">
         <a:glow rad="63500"><a:srgbClr val="FF0000"/></a:glow>
         <a:softEdge rad="12700"/>
       </a:effectLst>`,
    );
    expect(list.map((e) => e.type)).toEqual(['glow', 'softEdge']);
    expect(list[0]).toMatchObject({ type: 'glow', color: { hex: 'FF0000' } });
  });

  it('returns an empty list when there is no effectLst', () => {
    expect(parseEffectLst(undefined, ctx)).toEqual([]);
  });
});
