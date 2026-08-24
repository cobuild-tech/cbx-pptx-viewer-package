/**
 * @vitest-environment jsdom
 *
 * The DOM half of PPTX editing: the markers the renderer emits, reading an
 * edited contentEditable subtree back into paragraphs, and the viewer's
 * commit/undo/keyboard behaviour.
 *
 * Only this file needs a DOM, so it opts in per-file rather than changing the
 * package-wide `environment: 'node'`.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { Deck } from '../src/pptx/deck/deck.js';
import { RelType } from '../src/pptx/relTypes.js';
import { Viewer } from '../src/pptx/viewer/viewer.js';
import { EditContext } from '../src/pptx/edit/context.js';
import { reconcileTextBody } from '../src/pptx/edit/reconcile.js';
import { renderTextBody, EDIT_ATTR } from '../src/pptx/text/render.js';
import type { PresetShape, TextBody } from '../src/pptx/model.js';

// jsdom has no ResizeObserver; the viewer only uses it to re-fit on resize.
vi.stubGlobal(
  'ResizeObserver',
  class {
    observe() {}
    unobserve() {}
    disconnect() {}
  },
);

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
        <p:sldSz cx="12192000" cy="6858000"/>
      </p:presentation>`,
    ),
    'ppt/_rels/presentation.xml.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.Slide}" Target="slides/slide1.xml"/>`),
    ),
    'ppt/slides/slide1.xml': strToU8(
      `<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="1000000"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/>
            <a:p><a:pPr marL="457200" indent="-457200"><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="2400"/><a:t>Hello </a:t></a:r><a:r><a:rPr sz="2400" b="1"/><a:t>World</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="3" name="Dated"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="2000000"/><a:ext cx="3000000" cy="500000"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/>
            <a:p><a:fld id="{1}" type="datetime"><a:rPr sz="1200"/><a:t>1/1/2024</a:t></a:fld></a:p>
          </p:txBody>
        </p:sp>
      </p:spTree></p:cSld></p:sld>`,
    ),
  });
}

function bodyAt(deck: Deck, i: number): TextBody {
  return (deck.slides[0]!.shapes[i] as PresetShape).text!;
}

function textOf(body: TextBody): string {
  return body.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
}

/** Render a body for editing and give back the element plus its context. */
function renderEditable(deck: Deck, shapeIndex: number) {
  const ctx = new EditContext(deck, 0);
  const body = bodyAt(deck, shapeIndex);
  const el = renderTextBody(body, { imageUrl: () => undefined, edit: ctx });
  return { ctx, body, el, resolve: (k: string | null | undefined) => ctx.resolve(k) };
}

describe('pptx edit — render markers', () => {
  it('marks the body, paragraphs and runs, and makes the body editable', () => {
    const deck = Deck.load(buildDeck());
    const { el } = renderEditable(deck, 0);

    expect(el.getAttribute(EDIT_ATTR.body)).toBeTruthy();
    expect(el.contentEditable).toBe('true');
    expect(el.querySelectorAll(`[${EDIT_ATTR.para}]`)).toHaveLength(1);
    expect(el.querySelectorAll(`[${EDIT_ATTR.run}]`)).toHaveLength(2);
  });

  it('locks bullets and field text against editing', () => {
    const deck = Deck.load(buildDeck());

    const bullet = renderEditable(deck, 0).el.querySelector('span:not([data-cbx-run])');
    expect(bullet?.textContent).toBe('•');
    expect((bullet as HTMLElement).contentEditable).toBe('false');

    const fieldRun = renderEditable(deck, 1).el.querySelector(`[${EDIT_ATTR.run}]`);
    expect(fieldRun?.textContent).toBe('1/1/2024');
    expect((fieldRun as HTMLElement).contentEditable).toBe('false');
  });
});

describe('pptx edit — reconciliation', () => {
  it('reads back an unedited body unchanged', () => {
    const deck = Deck.load(buildDeck());
    const { el, resolve, body } = renderEditable(deck, 0);

    const paras = reconcileTextBody(el, resolve);
    expect(paras).toHaveLength(1);
    expect(paras[0]!.src).toBe(body.paragraphs[0]);
    expect(paras[0]!.segments.map((s) => s.text)).toEqual(['Hello ', 'World']);
    // The bullet is decoration and must not have become a run.
    expect(paras[0]!.segments.every((s) => s.text !== '•')).toBe(true);
  });

  it('keeps typed text attached to the run it was typed into', () => {
    const deck = Deck.load(buildDeck());
    const { el, resolve, body } = renderEditable(deck, 0);

    const firstRun = el.querySelector(`[${EDIT_ATTR.run}]`)!;
    firstRun.textContent = 'Goodbye ';

    const segs = reconcileTextBody(el, resolve)[0]!.segments;
    expect(segs.map((s) => s.text)).toEqual(['Goodbye ', 'World']);
    expect(segs[0]!.src).toBe(body.paragraphs[0]!.runs[0]);
  });

  it('gives text typed outside any run the neighbouring run as its source', () => {
    const deck = Deck.load(buildDeck());
    const { el, resolve, body } = renderEditable(deck, 0);

    // What the browser leaves when you type at the very start of a paragraph.
    const para = el.querySelector(`[${EDIT_ATTR.para}]`)!;
    para.insertBefore(el.ownerDocument.createTextNode('Oh '), para.firstChild);

    // It inherits the following run's source, then coalesces into it — so the
    // typed text ends up in one run carrying that run's formatting.
    const segs = reconcileTextBody(el, resolve)[0]!.segments;
    expect(segs.map((s) => s.text)).toEqual(['Oh Hello ', 'World']);
    expect(segs[0]!.src).toBe(body.paragraphs[0]!.runs[0]);
  });

  it('treats a duplicated paragraph block as a split sharing one source', () => {
    const deck = Deck.load(buildDeck());
    const { el, resolve, body } = renderEditable(deck, 0);

    // Chrome clones the paragraph div on Enter, marker attribute included.
    const para = el.querySelector(`[${EDIT_ATTR.para}]`)!;
    const clone = para.cloneNode(true) as Element;
    para.after(clone);

    const paras = reconcileTextBody(el, resolve);
    expect(paras).toHaveLength(2);
    expect(paras[0]!.src).toBe(body.paragraphs[0]);
    expect(paras[1]!.src).toBe(body.paragraphs[0]);
  });

  it('drops the filler <br> browsers put in an emptied paragraph', () => {
    const deck = Deck.load(buildDeck());
    const { el, resolve } = renderEditable(deck, 0);

    const para = el.querySelector(`[${EDIT_ATTR.para}]`)!;
    para.innerHTML = '<br>';

    expect(reconcileTextBody(el, resolve)[0]!.segments).toEqual([]);
  });

  it('turns an interior <br> into a line break', () => {
    const deck = Deck.load(buildDeck());
    const { el, resolve } = renderEditable(deck, 0);

    const runs = el.querySelectorAll(`[${EDIT_ATTR.run}]`);
    runs[0]!.after(el.ownerDocument.createElement('br'));

    const segs = reconcileTextBody(el, resolve)[0]!.segments;
    expect(segs.map((s) => ({ text: s.text, br: !!s.isBreak }))).toEqual([
      { text: 'Hello ', br: false },
      { text: '\n', br: true },
      { text: 'World', br: false },
    ]);
  });

  it('coalesces adjacent text that shares a run and format', () => {
    const deck = Deck.load(buildDeck());
    const { el, resolve } = renderEditable(deck, 0);

    // Browsers routinely split a run's text node while typing.
    const firstRun = el.querySelector(`[${EDIT_ATTR.run}]`)!;
    firstRun.textContent = '';
    firstRun.append(
      el.ownerDocument.createTextNode('Hel'),
      el.ownerDocument.createTextNode('lo '),
    );

    const segs = reconcileTextBody(el, resolve)[0]!.segments;
    expect(segs.map((s) => s.text)).toEqual(['Hello ', 'World']);
  });

  it('reads a format marker as an override on the wrapped run', () => {
    const deck = Deck.load(buildDeck());
    const { el, resolve, body } = renderEditable(deck, 0);

    const firstRun = el.querySelector(`[${EDIT_ATTR.run}]`)!;
    const wrapper = el.ownerDocument.createElement('span');
    wrapper.setAttribute(EDIT_ATTR.fmt, JSON.stringify({ bold: true, colorHex: '00FF00' }));
    firstRun.replaceWith(wrapper);
    wrapper.appendChild(firstRun);

    const segs = reconcileTextBody(el, resolve)[0]!.segments;
    expect(segs[0]!.format).toEqual({ bold: true, colorHex: '00FF00' });
    expect(segs[0]!.src).toBe(body.paragraphs[0]!.runs[0]);
    expect(segs[1]!.format).toBeUndefined();
  });
});

describe('pptx edit — text box affordances', () => {
  beforeEach(() => {
    document.head.innerHTML = '';
    document.body.innerHTML = '';
  });

  it('leaves no inline outline that would beat the stylesheet', () => {
    const deck = Deck.load(buildDeck());
    const { el } = renderEditable(deck, 0);
    expect(el.style.outline).toBe('');
  });

  it('installs a stylesheet outlining editable boxes on hover and focus', () => {
    const deck = Deck.load(buildDeck());
    const container = document.createElement('div');
    document.body.appendChild(container);
    const viewer = new Viewer(deck, container, { editable: true, webFonts: false });

    const style = document.getElementById('cbx-edit-styles');
    expect(style).not.toBeNull();
    expect(style!.textContent).toContain('[data-cbx-body]:hover');
    expect(style!.textContent).toContain('cursor:text');
    // Only the focused box is outlined solid.
    expect(style!.textContent).toContain('[data-cbx-body]:focus');

    viewer.destroy();
    expect(document.getElementById('cbx-edit-styles')).toBeNull();
  });

  it('outlines every box in "always" mode and none in "none"', () => {
    const deck = Deck.load(buildDeck());
    const container = document.createElement('div');
    document.body.appendChild(container);
    const viewer = new Viewer(deck, container, {
      editable: true,
      webFonts: false,
      textBoxOutline: 'always',
    });

    const style = () => document.getElementById('cbx-edit-styles')!.textContent!;
    expect(style()).toContain('[data-cbx-body]:not(:focus)');

    viewer.setTextBoxOutline('none');
    expect(style()).not.toContain('[data-cbx-body]:not(:focus)');
    expect(style()).not.toContain(':hover:not(:focus)');
    // Focus feedback survives in every mode.
    expect(style()).toContain('[data-cbx-body]:focus');
    viewer.destroy();
  });

  it('adds no stylesheet when the viewer is read-only', () => {
    const deck = Deck.load(buildDeck());
    const container = document.createElement('div');
    document.body.appendChild(container);
    const viewer = new Viewer(deck, container, { editable: false, webFonts: false });
    expect(document.getElementById('cbx-edit-styles')).toBeNull();
    viewer.destroy();
  });

  it('keeps the stylesheet until the last editable viewer is destroyed', () => {
    const a = document.createElement('div');
    const b = document.createElement('div');
    document.body.append(a, b);
    const v1 = new Viewer(Deck.load(buildDeck()), a, { editable: true, webFonts: false });
    const v2 = new Viewer(Deck.load(buildDeck()), b, { editable: true, webFonts: false });

    v1.destroy();
    expect(document.getElementById('cbx-edit-styles')).not.toBeNull();
    v2.destroy();
    expect(document.getElementById('cbx-edit-styles')).toBeNull();
  });
});

describe('pptx edit — viewer', () => {
  let container: HTMLElement;

  beforeEach(() => {
    document.body.innerHTML = '';
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  function mount(editable: boolean) {
    const deck = Deck.load(buildDeck());
    const viewer = new Viewer(deck, container, { editable, webFonts: false });
    return { deck, viewer };
  }

  it('renders read-only text bodies by default', () => {
    mount(false);
    expect(container.querySelector(`[${EDIT_ATTR.body}]`)).toBeNull();
  });

  it('commits an edit when focus leaves the text body', () => {
    const { deck, viewer } = mount(true);
    const bodyEl = container.querySelector(`[${EDIT_ATTR.body}]`) as HTMLElement;
    const run = bodyEl.querySelector(`[${EDIT_ATTR.run}]`)!;
    run.textContent = 'Farewell ';

    bodyEl.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));

    expect(textOf(bodyAt(deck, 0))).toBe('Farewell World');
    expect(viewer.hasEdits).toBe(true);
    // The slide was re-rendered from the committed XML.
    expect(container.textContent).toContain('Farewell');
  });

  it('undoes and redoes a committed edit', () => {
    const { deck, viewer } = mount(true);
    const bodyEl = container.querySelector(`[${EDIT_ATTR.body}]`) as HTMLElement;
    bodyEl.querySelector(`[${EDIT_ATTR.run}]`)!.textContent = 'Farewell ';
    bodyEl.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));
    expect(viewer.canUndo).toBe(true);

    viewer.undo();
    expect(textOf(bodyAt(deck, 0))).toBe('Hello World');

    viewer.redo();
    expect(textOf(bodyAt(deck, 0))).toBe('Farewell World');
  });

  it('does not navigate slides while the caret is in text', () => {
    const { viewer } = mount(true);
    const bodyEl = container.querySelector(`[${EDIT_ATTR.body}]`) as HTMLElement;

    const key = new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true, cancelable: true });
    bodyEl.dispatchEvent(key);
    expect(key.defaultPrevented).toBe(false);
    expect(viewer.currentIndex).toBe(0);

    // The same key outside the text body still navigates.
    const outside = new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true, cancelable: true });
    container.dispatchEvent(outside);
    expect(outside.defaultPrevented).toBe(true);
  });

  it('exports a deck that reloads with the edit applied', () => {
    const { viewer } = mount(true);
    const bodyEl = container.querySelector(`[${EDIT_ATTR.body}]`) as HTMLElement;
    bodyEl.querySelector(`[${EDIT_ATTR.run}]`)!.textContent = 'Farewell ';
    bodyEl.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));

    const reloaded = Deck.load(viewer['deck'].toBytes());
    expect(textOf(bodyAt(reloaded, 0))).toBe('Farewell World');
  });
});
