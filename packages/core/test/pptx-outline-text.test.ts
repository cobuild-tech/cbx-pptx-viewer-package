/**
 * @vitest-environment jsdom
 *
 * Gradient outlines, the clip around custom geometry, and the paragraph
 * strut. jsdom does no layout, so these pin down the markup that the browser
 * measurements depend on: a gradient stroke laid out in user space (a flat
 * line's bounding box has no height), a clip that leaves the outline whole,
 * and a paragraph whose strut is in its text's own face.
 */
import { describe, it, expect } from 'vitest';
import { parseXml } from '../src/oxml/xml.js';
import { parseTheme } from '../src/pptx/color.js';
import { strokeFromLn } from '../src/pptx/shapes/fill.js';
import { customGeomSvg, strokeOverlay } from '../src/pptx/shapes/render.js';
import { parseTextBody } from '../src/pptx/text/text.js';
import { renderTextBody } from '../src/pptx/text/render.js';
import { TextStyleChain } from '../src/pptx/text/textStyles.js';
import type { ParseScope } from '../src/pptx/scope.js';
import type { RenderDeps } from '../src/pptx/render/primitives.js';
import type { Stroke } from '../src/pptx/model.js';

const colorCtx = { theme: parseTheme(undefined), clrMap: {} };
const scope: ParseScope = { colorCtx, resolveImage: () => undefined };
const deps = { imageUrl: () => undefined } as unknown as RenderDeps;

const GRADIENT_LN = `<a:ln xmlns:a="a" w="38100">
  <a:gradFill><a:gsLst>
    <a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs>
    <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
  </a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>
  <a:headEnd type="triangle"/><a:tailEnd type="triangle"/>
</a:ln>`;

const ln = (xml: string) => strokeFromLn(parseXml(xml)!, scope);

describe('gradient outlines', () => {
  it('parses a gradFill outline, keeping its first stop as the plain colour', () => {
    const s = ln(GRADIENT_LN)!;
    expect(s).toBeDefined();
    expect(s.color.hex).toBe('00FF00');
    expect(s.gradient?.stops.map((g) => g.color.hex)).toEqual(['00FF00', '0000FF']);
    expect(s.gradient?.angle).toBe(90); // ang 0 = left to right
  });

  it('still prefers a solid fill, and still drops a noFill outline', () => {
    expect(ln(`<a:ln xmlns:a="a"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>`)?.gradient).toBeUndefined();
    expect(ln(`<a:ln xmlns:a="a"><a:noFill/></a:ln>`)).toBeUndefined();
  });

  /** The gradient a stroked path paints with. */
  function strokeGradient(svg: SVGSVGElement): Element {
    const ref = svg.querySelector(':scope > path')!.getAttribute('stroke')!;
    expect(ref).toMatch(/^url\(#.+\)$/);
    return svg.querySelector(ref.slice(4, -1))!;
  }

  it('strokes a preset outline with a gradient spanning the box in user space', () => {
    const svg = strokeOverlay('M0,0 L200,0', 200, 0, ln(GRADIENT_LN)!, false);
    const grad = strokeGradient(svg);
    expect(grad.tagName).toBe('linearGradient');
    // A horizontal line's bounding box has zero height, which would paint
    // nothing in objectBoundingBox units.
    expect(grad.getAttribute('gradientUnits')).toBe('userSpaceOnUse');
    expect([grad.getAttribute('x1'), grad.getAttribute('x2')]).toEqual(['0.00', '200.00']);
    expect([...grad.querySelectorAll('stop')].map((s) => s.getAttribute('stop-color'))).toEqual([
      '#00FF00',
      '#0000FF',
    ]);
  });

  it('gives each arrowhead the colour of its own end of the line', () => {
    const svg = strokeOverlay('M0,0 L200,0', 200, 0, ln(GRADIENT_LN)!, false);
    const fillOf = (attr: string) => {
      const id = svg.querySelector(':scope > path')!.getAttribute(attr)!.slice(4, -1);
      return svg.querySelector(`${id} path`)!.getAttribute('fill');
    };
    expect(fillOf('marker-start')).toBe('#00FF00');
    expect(fillOf('marker-end')).toBe('#0000FF');
  });

  it('strokes custom geometry in the path coordinate space', () => {
    const svg = customGeomSvg('M0,0 L1000,500', 1000, 500, 100, 50, undefined, ln(GRADIENT_LN)!, deps);
    const grad = strokeGradient(svg);
    expect(grad.getAttribute('x2')).toBe('1000.00');
    expect(grad.getAttribute('y1')).toBe('250.00');
  });
});

describe('custom geometry clip', () => {
  it('clips near the box, not flush with it, so the outline survives', () => {
    const stroke: Stroke = { color: { hex: '000000' }, width: 4 } as Stroke;
    const svg = customGeomSvg('M0,0 L100,100', 100, 100, 200, 100, undefined, stroke, deps);
    expect(svg.style.overflow).toBe('visible');
    // The stroke's width plus 10% of the larger side.
    expect(svg.style.clipPath).toBe('inset(-24.00px)');
  });
});

describe('paragraph strut', () => {
  const body = (paras: string) =>
    parseTextBody(
      parseXml(`<p:txBody xmlns:p="p" xmlns:a="a"><a:bodyPr/>${paras}</p:txBody>`)!,
      new TextStyleChain([], colorCtx),
      colorCtx,
      scope,
    );
  const run = (sz: number, font: string, text: string) =>
    `<a:r><a:rPr sz="${sz}"><a:latin typeface="${font}"/></a:rPr><a:t>${text}</a:t></a:r>`;

  it('takes its face from the run that sets its size', () => {
    const b = body(`<a:p>${run(2400, 'Georgia', 'Big')}${run(1100, 'Open Sans', 'small')}</a:p>`);
    expect(b.paragraphs[0]!.defaultSizePt).toBe(11);
    expect(b.paragraphs[0]!.defaultFont).toBe('Open Sans');
  });

  it("falls back to the paragraph mark's face for an empty paragraph", () => {
    const b = body(`<a:p><a:endParaRPr sz="1100"><a:latin typeface="Open Sans"/></a:endParaRPr></a:p>`);
    expect(b.paragraphs[0]!.defaultFont).toBe('Open Sans');
  });

  it('renders the strut in the text face, minus any weight suffix', () => {
    const el = renderTextBody(body(`<a:p>${run(1100, 'Open Sans Light', 'text')}</a:p>`), deps);
    expect((el.firstElementChild as HTMLElement).style.fontFamily).toMatch(/^"Open Sans",/);
  });

  it('keeps a marker no bigger than its text out of the line height', () => {
    const bullet = (sz: string) =>
      `<a:p><a:pPr marL="171450" indent="-171450"><a:buFont typeface="Arial"/>${sz}<a:buChar char="•"/></a:pPr>${run(1100, 'Open Sans', 'item')}</a:p>`;
    const marker = (xml: string) =>
      renderTextBody(body(xml), deps).querySelector('span[contenteditable="false"]') as HTMLElement;
    expect(marker(bullet('')).style.lineHeight).toBe('0');
    // A bigger marker may legitimately make its line taller.
    expect(marker(bullet('<a:buSzPct val="150000"/>')).style.lineHeight).toBe('');
  });
});
