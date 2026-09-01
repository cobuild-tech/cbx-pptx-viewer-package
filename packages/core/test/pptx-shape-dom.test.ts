/**
 * @vitest-environment jsdom
 *
 * The interaction half of shape editing: what the renderer marks as selectable,
 * and what clicking, dragging, nudging and double-clicking actually do to the
 * deck. The geometry itself is covered in pptx-shape-geometry and the XML in
 * pptx-shape-edit, so these tests are about wiring.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { Deck } from '../src/pptx/deck/deck.js';
import { RelType } from '../src/pptx/relTypes.js';
import { Viewer } from '../src/pptx/viewer/viewer.js';
import { EDIT_ATTR } from '../src/oxml/edit/attrs.js';
import type { Shape } from '../src/pptx/model.js';

// jsdom has no ResizeObserver; the viewer only uses it to re-fit on resize.
vi.stubGlobal(
  'ResizeObserver',
  class {
    observe() {}
    unobserve() {}
    disconnect() {}
  },
);
// …nor pointer capture, which the drag uses to keep receiving events.
Object.assign(HTMLElement.prototype, {
  setPointerCapture() {},
  releasePointerCapture() {},
  hasPointerCapture: () => false,
});

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

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
      `<p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="9" name="LayoutDecoration"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
        </p:sp>
      </p:spTree></p:cSld></p:sldLayout>`,
    ),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': strToU8(rels('')),
    'ppt/slides/slide1.xml': strToU8(
      `<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
        <p:nvGrpSpPr><p:cNvPr id="1" name="Shape Tree"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1800"/><a:t>Hello</a:t></a:r></a:p></p:txBody>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="3" name="Other"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="4762500" y="952500"/><a:ext cx="952500" cy="952500"/></a:xfrm></p:spPr>
        </p:sp>
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

function mount(editable = true): Viewer {
  return new Viewer(Deck.load(buildDeck()), container, {
    editable,
    filmstrip: false,
    webFonts: false,
  });
}

/** The rendered element of the shape with this name. */
function shapeEl(name: string): HTMLElement {
  const el = [...container.querySelectorAll<HTMLElement>(`[${EDIT_ATTR.shape}]`)].find(
    (e) => e.getAttribute('data-name') === name,
  );
  // The renderer does not emit names, so fall back to document order: Box then
  // Other, matching the shape tree.
  return el ?? container.querySelectorAll<HTMLElement>(`[${EDIT_ATTR.shape}]`)[name === 'Box' ? 0 : 1]!;
}

/** jsdom has no PointerEvent; a MouseEvent of the same type dispatches fine. */
function pointer(type: string, target: EventTarget, x: number, y: number, init: MouseEventInit = {}) {
  target.dispatchEvent(
    new MouseEvent(type, { bubbles: true, cancelable: true, clientX: x, clientY: y, ...init }),
  );
}

/** A press and release on a shape with no travel in between. */
function click(name: string, x: number, y: number) {
  const host = container.querySelector<HTMLElement>('.cbx-sel')!.parentElement!;
  pointer('pointerdown', shapeEl(name), x, y);
  pointer('pointerup', host, x, y);
}

function drag(viewer: Viewer, name: string, from: [number, number], to: [number, number]) {
  const host = container.querySelector<HTMLElement>('.cbx-sel')!.parentElement!;
  pointer('pointerdown', shapeEl(name), from[0], from[1]);
  pointer('pointermove', host, to[0], to[1]);
  pointer('pointerup', host, to[0], to[1]);
  return viewer;
}

const named = (viewer: Viewer, name: string): Shape =>
  viewer.deckShapes().find((s) => s.name === name)!;

// A tiny accessor keeps the tests reading in terms of the model rather than
// reaching through the viewer's internals.
declare module '../src/pptx/viewer/viewer.js' {
  interface Viewer {
    deckShapes(): Shape[];
  }
}
Viewer.prototype.deckShapes = function (this: Viewer): Shape[] {
  return (this as unknown as { deck: Deck }).deck.slides[0]!.shapes;
};

describe('pptx shapes — selectable markers', () => {
  it('marks the slide’s own shapes and not the layout’s', () => {
    mount();
    const marked = container.querySelectorAll(`[${EDIT_ATTR.shape}]`);
    // Two slide shapes; the layout decoration renders but is not addressable.
    expect(marked).toHaveLength(2);
  });

  it('marks nothing when the viewer is read-only', () => {
    mount(false);
    expect(container.querySelectorAll(`[${EDIT_ATTR.shape}]`)).toHaveLength(0);
    expect(container.querySelector('.cbx-sel')).toBeNull();
  });
});

describe('pptx shapes — selection', () => {
  it('selects a shape on press and shows its handles', () => {
    const viewer = mount();
    pointer('pointerdown', shapeEl('Box'), 120, 120);
    expect(viewer.selectedShapes).toHaveLength(1);
    expect(viewer.selectedShapes[0]!.name).toBe('Box');

    const frame = container.querySelector<HTMLElement>('.cbx-sel-frame')!;
    expect(frame.hidden).toBe(false);
    expect(frame.querySelectorAll('.cbx-sel-handle')).toHaveLength(8);
    viewer.destroy();
  });

  it('adds to the selection with shift and clears on empty background', () => {
    const viewer = mount();
    pointer('pointerdown', shapeEl('Box'), 120, 120);
    pointer('pointerdown', shapeEl('Other'), 520, 120, { shiftKey: true });
    expect(viewer.selectedShapes).toHaveLength(2);
    // Handles are hidden for a multiple selection — we do not resize as a block.
    expect(container.querySelector<HTMLElement>('.cbx-sel-frame')!.hidden).toBe(true);

    const slide = container.querySelector<HTMLElement>('.pptx-slide')!;
    pointer('pointerdown', slide, 5, 600);
    expect(viewer.selectedShapes).toHaveLength(0);
    viewer.destroy();
  });

  it('does not select a shape the slide inherited', () => {
    const viewer = mount();
    // The layout decoration renders without a key, so a press on it reads as a
    // press on the background.
    const slide = container.querySelector<HTMLElement>('.pptx-slide')!;
    pointer('pointerdown', slide.firstElementChild!, 10, 10);
    expect(viewer.selectedShapes).toHaveLength(0);
    viewer.destroy();
  });
});

describe('pptx shapes — manipulation', () => {
  it('commits a drag to the deck', () => {
    const viewer = mount();
    expect(named(viewer, 'Box').transform!.x).toBe(100);

    drag(viewer, 'Box', [120, 120], [180, 150]);

    expect(named(viewer, 'Box').transform!.x).toBe(160);
    expect(named(viewer, 'Box').transform!.y).toBe(130);
    expect(viewer.hasEdits).toBe(true);
    // The shape is still selected afterwards, addressed through its XML node.
    expect(viewer.selectedShapes).toHaveLength(1);
    viewer.destroy();
  });

  it('treats a press that does not travel as a click, not a drag', () => {
    const viewer = mount();
    drag(viewer, 'Box', [120, 120], [121, 121]);
    expect(named(viewer, 'Box').transform!.x).toBe(100);
    expect(viewer.hasEdits).toBe(false);
    viewer.destroy();
  });

  it('nudges with the arrow keys, ten times as far with shift', () => {
    const viewer = mount();
    pointer('pointerdown', shapeEl('Box'), 120, 120);

    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
    expect(named(viewer, 'Box').transform!.x).toBe(101);

    container.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'ArrowDown', shiftKey: true, bubbles: true }),
    );
    expect(named(viewer, 'Box').transform!.y).toBe(110);
    viewer.destroy();
  });

  it('deletes the selection with Delete and puts it back on undo', () => {
    const viewer = mount();
    pointer('pointerdown', shapeEl('Box'), 120, 120);
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true }));

    expect(viewer.deckShapes().map((s) => s.name)).not.toContain('Box');
    expect(viewer.selectedShapes).toHaveLength(0);

    viewer.undo();
    expect(viewer.deckShapes().map((s) => s.name)).toContain('Box');
    viewer.destroy();
  });

  it('clears the selection with Escape instead of changing slides', () => {
    const viewer = mount();
    pointer('pointerdown', shapeEl('Box'), 120, 120);
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(viewer.selectedShapes).toHaveLength(0);
    viewer.destroy();
  });

  it('restacks with Ctrl+]', () => {
    const viewer = mount();
    pointer('pointerdown', shapeEl('Box'), 120, 120);
    container.dispatchEvent(
      new KeyboardEvent('keydown', { key: ']', ctrlKey: true, bubbles: true }),
    );
    // Box was first of the slide's own shapes; forward puts it after Other.
    expect(viewer.deckShapes().map((s) => s.name)).toEqual([
      'LayoutDecoration',
      'Other',
      'Box',
    ]);
    viewer.destroy();
  });
});

describe('pptx shapes — text mode', () => {
  it('opens the text on a second click, and leaves it on Escape', () => {
    const viewer = mount();
    const body = () => container.querySelector<HTMLElement>(`[${EDIT_ATTR.body}]`)!;
    // Selection first: the box is addressable but not typeable.
    expect(body().getAttribute('contenteditable')).not.toBe('true');

    // The first click selects the shape…
    click('Box', 120, 120);
    expect(viewer.isEditingText).toBe(false);
    // …and a second click on the shape already selected opens its text, which
    // is both PowerPoint's gesture and what a double-click amounts to.
    click('Box', 120, 120);
    expect(viewer.isEditingText).toBe(true);
    expect(body().getAttribute('contenteditable')).toBe('true');
    // The handles step out of the way while typing.
    expect(container.querySelector<HTMLElement>('.cbx-sel')!.style.display).toBe('none');

    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(viewer.isEditingText).toBe(false);
    expect(body().getAttribute('contenteditable')).not.toBe('true');
    viewer.destroy();
  });

  it('opens text editing from the keyboard with F2', () => {
    const viewer = mount();
    pointer('pointerdown', shapeEl('Box'), 120, 120);
    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true }));
    expect(viewer.isEditingText).toBe(true);
    viewer.destroy();
  });
});
