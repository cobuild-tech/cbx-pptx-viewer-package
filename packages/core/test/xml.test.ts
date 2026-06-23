import { describe, it, expect } from 'vitest';
import { parseXml, child, children, path, attr, attrNum, attrBool } from '../src/oxml/xml.js';

describe('xml normalization', () => {
  const xml = parseXml(`
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title 1"/></p:nvSpPr>
      <p:spPr>
        <a:xfrm rot="5400000">
          <a:off x="100" y="200"/>
          <a:ext cx="3000" cy="4000"/>
        </a:xfrm>
      </p:spPr>
      <p:txBody>
        <a:p><a:r><a:t>Hello </a:t></a:r><a:br/><a:r><a:t>world</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>`)!;

  it('parses the root element', () => {
    expect(xml.name).toBe('p:sp');
  });

  it('matches by local name when no prefix is given', () => {
    expect(child(xml, 'spPr')?.name).toBe('p:spPr');
    expect(child(xml, 'p:spPr')?.name).toBe('p:spPr');
  });

  it('reads attributes by local or qualified name', () => {
    const off = path(xml, 'spPr/xfrm/off')!;
    expect(attrNum(off, 'x')).toBe(100);
    expect(attrNum(off, 'y')).toBe(200);
    expect(attrNum(child(xml, 'spPr')!.children[0], 'rot')).toBe(5400000);
  });

  it('reads namespaced attributes prefix-robustly', () => {
    const n = parseXml('<p:sldId id="256" r1:id="rId1"/>')!;
    expect(attr(n, 'r:id')).toBe('rId1');
    expect(attr(n, 'id')).toBe('256');
  });

  it('preserves run/break order inside a paragraph (z-order semantics)', () => {
    const p = path(xml, 'txBody/p')!;
    expect(p.children.map((c) => c.name)).toEqual(['a:r', 'a:br', 'a:r']);
  });

  it('preserves significant whitespace in text runs', () => {
    const runs = children(path(xml, 'txBody/p')!, 'r');
    expect(child(runs[0], 't')?.text).toBe('Hello ');
    expect(child(runs[1], 't')?.text).toBe('world');
  });

  it('parses OOXML booleans', () => {
    const n = parseXml('<a:rPr b="1" i="0"/>')!;
    expect(attrBool(n, 'b')).toBe(true);
    expect(attrBool(n, 'i')).toBe(false);
    expect(attrBool(n, 'u', true)).toBe(true);
    expect(attr(n, 'missing')).toBeUndefined();
  });
});
