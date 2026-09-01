/**
 * @vitest-environment jsdom
 *
 * Reaching inside a group, and typing in a table cell — the two "box inside a
 * box" cases. A group's children are marked like any other shape, but a press
 * picks the group until the user steps into it, and a child's box is stated in
 * the group's own child space, so both the handles and the committed geometry
 * have to come back through that mapping.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { Deck } from '../src/pptx/deck/deck.js';
import { RelType } from '../src/pptx/relTypes.js';
import { Viewer } from '../src/pptx/viewer/viewer.js';
import { EDIT_ATTR } from '../src/oxml/edit/attrs.js';
import type { Shape } from '../src/pptx/model.js';

vi.stubGlobal(
  'ResizeObserver',
  class {
    observe() {}
    unobserve() {}
    disconnect() {}
  },
);
Object.assign(HTMLElement.prototype, {
  setPointerCapture() {},
  releasePointerCapture() {},
  hasPointerCapture: () => false,
});

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

/**
 * One group whose child space is half the size of its box (so everything inside
 * is drawn at 2×), plus a one-cell table.
 *
 *   group   off 96,96   ext 192×96     chOff 0,0   chExt 96×48
 *   inner   off 48,24    ext 48×24     -> slide 192,144 96×48
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
        <p:sldSz cx="9144000" cy="6858000"/>
      </p:presentation>`,
    ),
    'ppt/_rels/presentation.xml.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.Slide}" Target="slides/slide1.xml"/>`),
    ),
    'ppt/slides/_rels/slide1.xml.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.SlideLayout}" Target="../slideLayouts/slideLayout1.xml"/>`),
    ),
    'ppt/slideLayouts/slideLayout1.xml': strToU8(
      `<p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
    ),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': strToU8(rels('')),
    'ppt/slides/slide1.xml': strToU8(
      `<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
        <p:nvGrpSpPr><p:cNvPr id="1" name="Shape Tree"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:grpSp>
          <p:nvGrpSpPr><p:cNvPr id="2" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr><a:xfrm>
            <a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/>
            <a:chOff x="0" y="0"/><a:chExt cx="914400" cy="457200"/>
          </a:xfrm></p:grpSpPr>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="3" name="Inner"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="457200" y="228600"/><a:ext cx="457200" cy="228600"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1800"/><a:t>Nested</a:t></a:r></a:p></p:txBody>
          </p:sp>
        </p:grpSp>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="4" name="Grid"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="3810000" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
            <a:tr h="914400"><a:tc>
              <a:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1800"/><a:t>Cell</a:t></a:r></a:p></a:txBody>
            </a:tc></a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld></p:sld>`,
    ),
  });
}

let container: HTMLElement;

beforeEach(() => {
  document.body.innerHTML = '';
  container = document.createElement('div');
  document.body.appendChild(container);
});

function mount(): Viewer {
  return new Viewer(Deck.load(buildDeck()), container, {
    editable: true,
    filmstrip: false,
    webFonts: false,
  });
}

/** jsdom has no PointerEvent; a MouseEvent of the same type dispatches fine. */
function pointer(type: string, target: EventTarget, x: number, y: number, init: MouseEventInit = {}) {
  target.dispatchEvent(
    new MouseEvent(type, { bubbles: true, cancelable: true, clientX: x, clientY: y, ...init }),
  );
}

const host = () => container.querySelector<HTMLElement>('.cbx-sel')!.parentElement!;

/** The rendered element of the shape with this name, at any depth. */
function shapeEl(viewer: Viewer, name: string): HTMLElement {
  const shape = find(viewer, name);
  const ctx = (viewer as unknown as { editCtx: { resolve(k: string): object | undefined } }).editCtx;
  return [...container.querySelectorAll<HTMLElement>(`[${EDIT_ATTR.shape}]`)].find(
    (el) => ctx.resolve(el.getAttribute(EDIT_ATTR.shape)!) === shape,
  )!;
}

/** A shape by name, searching inside groups too. */
function find(viewer: Viewer, name: string): Shape {
  const deck = (viewer as unknown as { deck: Deck }).deck;
  const walk = (shapes: readonly Shape[]): Shape | undefined => {
    for (const s of shapes) {
      if (s.name === name) return s;
      if (s.kind === 'group') {
        const hit = walk(s.children);
        if (hit) return hit;
      }
    }
    return undefined;
  };
  return walk(deck.slides[0]!.shapes)!;
}

function click(viewer: Viewer, name: string, x = 200, y = 150) {
  pointer('pointerdown', shapeEl(viewer, name), x, y);
  pointer('pointerup', host(), x, y);
}

describe('pptx shapes — inside a group', () => {
  it('marks a group and its children alike', () => {
    const viewer = mount();
    // Group, its child, and the table frame.
    expect(container.querySelectorAll(`[${EDIT_ATTR.shape}]`)).toHaveLength(3);
    viewer.destroy();
  });

  it('picks the group first, then the child on a second click', () => {
    const viewer = mount();
    click(viewer, 'Inner');
    expect(viewer.selectedShapes.map((s) => s.name)).toEqual(['Group']);

    click(viewer, 'Inner');
    expect(viewer.selectedShapes.map((s) => s.name)).toEqual(['Inner']);
    viewer.destroy();
  });

  it('draws the child’s handles at its place on the slide, not in the group’s space', () => {
    const viewer = mount();
    click(viewer, 'Inner');
    click(viewer, 'Inner');

    const frame = container.querySelector<HTMLElement>('.cbx-sel-frame')!;
    expect(frame.hidden).toBe(false);
    // off 48,24 in a space drawn at 2x, group at 96,96 -> 192,144 and 96x48.
    expect(frame.style.left).toBe('192px');
    expect(frame.style.top).toBe('144px');
    expect(frame.style.width).toBe('96px');
    expect(frame.style.height).toBe('48px');
    viewer.destroy();
  });

  it('commits a child drag in the group’s space, not the slide’s', () => {
    const viewer = mount();
    click(viewer, 'Inner');
    click(viewer, 'Inner');
    expect(find(viewer, 'Inner').transform!.x).toBe(48);

    const el = shapeEl(viewer, 'Inner');
    pointer('pointerdown', el, 200, 150);
    pointer('pointermove', host(), 240, 150);
    pointer('pointerup', host(), 240, 150);

    // 40px across the slide is 20px in a space drawn at 2x.
    expect(find(viewer, 'Inner').transform!.x).toBe(68);
    expect(find(viewer, 'Group').transform!.x).toBe(96);
    // The child is still selected after the rebuild the commit does.
    expect(viewer.selectedShapes.map((s) => s.name)).toEqual(['Inner']);
    viewer.destroy();
  });

  it('nudges the child in its own space too', () => {
    const viewer = mount();
    click(viewer, 'Inner');
    click(viewer, 'Inner');
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
    expect(find(viewer, 'Inner').transform!.x).toBe(49);
    viewer.destroy();
  });

  it('Escape steps back out to the group before clearing', () => {
    const viewer = mount();
    click(viewer, 'Inner');
    click(viewer, 'Inner');

    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(viewer.selectedShapes.map((s) => s.name)).toEqual(['Group']);

    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(viewer.selectedShapes).toHaveLength(0);

    // …and the next click starts from the group again.
    click(viewer, 'Inner');
    expect(viewer.selectedShapes.map((s) => s.name)).toEqual(['Group']);
    viewer.destroy();
  });

  it('opens the child’s text on a third click, once it is the selection', () => {
    const viewer = mount();
    click(viewer, 'Inner');
    click(viewer, 'Inner');
    click(viewer, 'Inner');
    expect(viewer.isEditingText).toBe(true);
    const box = container.querySelector<HTMLElement>(`[${EDIT_ATTR.body}][contenteditable="true"]`);
    expect(box?.textContent).toContain('Nested');
    viewer.destroy();
  });

  it('rubber-bands the slide’s own shapes, not what is inside a group', () => {
    const viewer = mount();
    const slide = container.querySelector<HTMLElement>('.pptx-slide')!;
    pointer('pointerdown', slide, 2, 2);
    pointer('pointermove', host(), 900, 700);
    pointer('pointerup', host(), 900, 700);
    expect(viewer.selectedShapes.map((s) => s.name)).toEqual(['Group', 'Grid']);
    viewer.destroy();
  });
});

describe('pptx shapes — a table cell', () => {
  it('types in a cell once the table is selected', () => {
    const viewer = mount();
    click(viewer, 'Grid', 450, 150);
    expect(viewer.selectedShapes.map((s) => s.name)).toEqual(['Grid']);

    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true }));
    expect(viewer.isEditingText).toBe(true);
    const box = container.querySelector<HTMLElement>(`[${EDIT_ATTR.body}][contenteditable="true"]`);
    expect(box?.textContent).toContain('Cell');
    viewer.destroy();
  });
});
