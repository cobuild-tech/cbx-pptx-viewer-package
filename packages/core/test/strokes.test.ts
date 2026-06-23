import { describe, it, expect } from 'vitest';
import { parseXml } from '../src/oxml/xml.js';
import { strokeFromLn } from '../src/pptx/shapes/fill.js';
import { parseTheme } from '../src/pptx/color.js';
import type { ParseScope } from '../src/pptx/scope.js';

const scope: ParseScope = {
  colorCtx: { theme: parseTheme(undefined), clrMap: {} },
  resolveImage: () => undefined,
};

function ln(xml: string) {
  return strokeFromLn(parseXml(xml)!, scope);
}

describe('strokeFromLn line ends', () => {
  it('parses a tail arrowhead with its size categories', () => {
    const s = ln(
      `<a:ln xmlns:a="a" w="28575"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
         <a:tailEnd type="arrow" w="med" len="lg"/></a:ln>`,
    );
    expect(s?.color.hex).toBe('FF0000');
    expect(s?.width).toBeCloseTo(28575 / 9525, 2);
    expect(s?.tailEnd).toEqual({ type: 'arrow', w: 'med', len: 'lg' });
    expect(s?.headEnd).toBeUndefined();
  });

  it('ignores a "none" end and defaults missing sizes to med', () => {
    const s = ln(
      `<a:ln xmlns:a="a"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>
         <a:headEnd type="none"/><a:tailEnd type="triangle"/></a:ln>`,
    );
    expect(s?.headEnd).toBeUndefined();
    expect(s?.tailEnd).toEqual({ type: 'triangle', w: 'med', len: 'med' });
  });
});
