import { describe, it, expect } from 'vitest';
import {
  parseXml,
  serializeXml,
  serializeNode,
  child,
  children,
  path,
  attr,
  attrNum,
  attrBool,
  createElement,
  cloneNode,
  setText,
  setAttr,
  removeAttr,
  insertChildAt,
  insertInOrder,
  removeChildAt,
  resolveIndexPath,
} from '../src/oxml/xml.js';

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

describe('xml serialization (round-trip)', () => {
  /** parse → serialize → parse must yield a structurally identical tree. */
  const roundTrips = (xml: string) => {
    const a = parseXml(xml)!;
    const b = parseXml(serializeXml(a))!;
    expect(b).toEqual(a);
  };

  it('round-trips a realistic WordprocessingML body', () => {
    roundTrips(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body>' +
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
        '<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>Title</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r>' +
        '<w:r><w:br/><w:t>world</w:t></w:r></w:p>' +
        '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>' +
        '</w:body></w:document>',
    );
  });

  it('round-trips even with insignificant inter-element whitespace', () => {
    roundTrips(`
      <w:p>
        <w:r><w:t>Hi</w:t></w:r>
      </w:p>`);
  });

  it('preserves child + break order through a round-trip', () => {
    const a = parseXml('<w:r><w:t>a</w:t><w:br/><w:t>b</w:t></w:r>')!;
    const b = parseXml(serializeXml(a))!;
    expect(b.children.map((c) => c.name)).toEqual(['w:t', 'w:br', 'w:t']);
  });

  it('escapes entities in text and attributes', () => {
    const node = parseXml('<w:t>a &amp; b &lt; c &gt; d</w:t>')!;
    expect(node.text).toBe('a & b < c > d');
    const out = serializeNode(node);
    expect(out).toContain('&amp;');
    expect(out).toContain('&lt;');
    expect(parseXml(out)!.text).toBe('a & b < c > d');

    const attrNode = parseXml('<x a="x &amp; &quot;y&quot;"/>')!;
    expect(attr(attrNode, 'a')).toBe('x & "y"');
    expect(parseXml(serializeNode(attrNode))!).toEqual(attrNode);
  });

  it('emits empty elements as self-closing', () => {
    expect(serializeNode(parseXml('<w:br/>')!)).toBe('<w:br/>');
    expect(serializeNode(createElement('w:b'))).toBe('<w:b/>');
  });

  it('preserves significant whitespace with xml:space', () => {
    roundTrips('<w:t xml:space="preserve">  spaced  </w:t>');
  });
});

describe('xml mutation helpers', () => {
  it('createElement / cloneNode are independent', () => {
    const el = createElement('w:t', { 'xml:space': 'preserve' }, [], 'hi');
    const copy = cloneNode(el);
    setText(copy, 'bye');
    setAttr(copy, 'w:val', '1');
    expect(el.text).toBe('hi');
    expect(attr(el, 'w:val')).toBeUndefined();
    expect(copy.text).toBe('bye');
    expect(attr(copy, 'w:val')).toBe('1');
  });

  it('removeAttr matches by qualified or local name', () => {
    const n = createElement('w:p', { 'w:rsidR': '00AB', 'w:rsidRDefault': '00CD' });
    removeAttr(n, 'rsidR');
    expect(attr(n, 'w:rsidR')).toBeUndefined();
    expect(attr(n, 'w:rsidRDefault')).toBe('00CD');
  });

  it('inserts, removes, and resolves child paths by index', () => {
    const body = createElement('w:body', {}, [
      createElement('w:p', {}, [], 'one'),
      createElement('w:p', {}, [], 'three'),
    ]);
    insertChildAt(body, 1, createElement('w:p', {}, [], 'two'));
    expect(body.children.map((c) => c.text)).toEqual(['one', 'two', 'three']);

    expect(resolveIndexPath(body, [2])?.text).toBe('three');
    expect(resolveIndexPath(body, [9])).toBeUndefined();

    const removed = removeChildAt(body, 0);
    expect(removed?.text).toBe('one');
    expect(body.children.map((c) => c.text)).toEqual(['two', 'three']);
  });

  it('insertInOrder places a child before the first one that must follow it', () => {
    // OOXML content models are sequences: a child in the wrong slot makes Office
    // call the part corrupt, so the order table decides the slot, not the caller.
    const order = ['lnSpc', 'spcBef', 'buFont', 'buChar'];
    const pPr = createElement('a:pPr', {}, [createElement('a:buChar', { char: '-' })]);

    insertInOrder(pPr, createElement('a:buFont', { typeface: 'Arial' }), order);
    insertInOrder(pPr, createElement('a:spcBef'), order);
    expect(pPr.children.map((c) => c.name)).toEqual(['a:spcBef', 'a:buFont', 'a:buChar']);

    // A name the table does not mention has nothing to be placed by, so it goes last.
    insertInOrder(pPr, createElement('a:extLst'), order);
    insertInOrder(pPr, createElement('a:lnSpc'), order);
    expect(pPr.children.map((c) => c.name)).toEqual([
      'a:lnSpc',
      'a:spcBef',
      'a:buFont',
      'a:buChar',
      'a:extLst',
    ]);
  });
});
