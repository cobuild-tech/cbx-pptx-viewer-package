/**
 * @vitest-environment jsdom
 *
 * The DOM half of DOCX editing: render markers, reading an edited paragraph
 * back, and the viewer's continuous-flow / commit / undo / keyboard behaviour.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { DocxDocument } from '../src/docx/document/document.js';
import { DocxRelType } from '../src/docx/relTypes.js';
import { DocxViewer } from '../src/docx/viewer/viewer.js';
import { DocxEditContext } from '../src/docx/edit/context.js';
import { reconcileParagraph } from '../src/docx/edit/reconcile.js';
import { renderBlock } from '../src/docx/render/dom.js';
import { FLOW_CLASS } from '../src/docx/edit/flow.js';
import { EDIT_ATTR } from '../src/oxml/edit/attrs.js';
import type { DocxParagraph } from '../src/docx/model.js';

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

function buildDocx(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
      </Types>`,
    ),
    '_rels/.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${DocxRelType.OfficeDocument}" Target="word/document.xml"/>`),
    ),
    'word/_rels/document.xml.rels': strToU8(
      rels(`<Relationship Id="rId10" Type="${DocxRelType.Numbering}" Target="numbering.xml"/>`),
    ),
    'word/numbering.xml': strToU8(
      `<w:numbering xmlns:w="w">
        <w:abstractNum w:abstractNumId="0">
          <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
        </w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
      </w:numbering>`,
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="w"><w:body>
        <w:p>
          <w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">Hello </w:t></w:r>
          <w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>World</w:t></w:r>
        </w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Listed</w:t></w:r></w:p>
        <w:p>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText>PAGE</w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>1</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>
      </w:body></w:document>`,
    ),
  });
}

function paras(doc: DocxDocument): DocxParagraph[] {
  return doc.sections[0]!.blocks.filter((b): b is DocxParagraph => b.kind === 'paragraph');
}

const textOf = (p: DocxParagraph) => p.runs.map((r) => r.text).join('');

/** Render one paragraph for editing, with its context. */
function renderEditable(doc: DocxDocument, index: number) {
  const ctx = new DocxEditContext(doc);
  const para = paras(doc)[index]!;
  const el = renderBlock(para, { imageUrl: () => undefined, edit: ctx });
  return { ctx, para, el, resolve: (k: string | null | undefined) => ctx.resolve(k) };
}

describe('docx edit — render markers', () => {
  it('marks the paragraph editable and keys each run', () => {
    const doc = DocxDocument.load(buildDocx());
    const { el } = renderEditable(doc, 0);

    expect(el.getAttribute(EDIT_ATTR.para)).toBeTruthy();
    expect(el.contentEditable).toBe('true');
    expect(el.querySelectorAll(`[${EDIT_ATTR.run}]`)).toHaveLength(2);
  });

  it('locks the generated list marker', () => {
    const doc = DocxDocument.load(buildDocx());
    const { el } = renderEditable(doc, 1);
    const marker = el.querySelector(`span:not([${EDIT_ATTR.run}])`);
    expect(marker?.textContent).toBe('1.');
    expect((marker as HTMLElement).contentEditable).toBe('false');
  });

  it('locks field text', () => {
    const doc = DocxDocument.load(buildDocx());
    const { el } = renderEditable(doc, 2);
    const field = el.querySelector(`[${EDIT_ATTR.run}]`);
    expect((field as HTMLElement).contentEditable).toBe('false');
  });
});

describe('docx edit — reconciliation', () => {
  it('reads back an unedited paragraph unchanged', () => {
    const doc = DocxDocument.load(buildDocx());
    const { el, resolve, para } = renderEditable(doc, 0);

    const edits = reconcileParagraph(el, resolve);
    expect(edits).toHaveLength(1);
    expect(edits[0]!.src).toBe(para);
    expect(edits[0]!.segments.map((s) => s.text)).toEqual(['Hello ', 'World']);
  });

  it('excludes the list marker from the segments', () => {
    const doc = DocxDocument.load(buildDocx());
    const { el, resolve } = renderEditable(doc, 1);
    const segs = reconcileParagraph(el, resolve)[0]!.segments;
    expect(segs.map((s) => s.text)).toEqual(['Listed']);
  });

  it('keeps typed text attached to the run it was typed into', () => {
    const doc = DocxDocument.load(buildDocx());
    const { el, resolve, para } = renderEditable(doc, 0);

    el.querySelector(`[${EDIT_ATTR.run}]`)!.textContent = 'Goodbye ';
    const segs = reconcileParagraph(el, resolve)[0]!.segments;
    expect(segs.map((s) => s.text)).toEqual(['Goodbye ', 'World']);
    expect(segs[0]!.src).toBe(para.runs[0]);
  });

  it('treats a block child as a paragraph split sharing one source', () => {
    const doc = DocxDocument.load(buildDocx());
    const { el, resolve, para } = renderEditable(doc, 0);

    // What the browser leaves after pressing Enter inside the paragraph.
    const second = el.ownerDocument.createElement('div');
    second.textContent = 'Second';
    el.appendChild(second);

    const edits = reconcileParagraph(el, resolve);
    expect(edits).toHaveLength(2);
    expect(edits.every((e) => e.src === para)).toBe(true);
    expect(edits[0]!.segments.map((s) => s.text)).toEqual(['Hello ', 'World']);
    expect(edits[1]!.segments.map((s) => s.text)).toEqual(['Second']);
  });

  it('drops the filler <br> in an emptied paragraph', () => {
    const doc = DocxDocument.load(buildDocx());
    const { el, resolve } = renderEditable(doc, 0);
    el.innerHTML = '<br>';
    expect(reconcileParagraph(el, resolve)[0]!.segments).toEqual([]);
  });

  it('turns an interior <br> into a break on the following segment', () => {
    const doc = DocxDocument.load(buildDocx());
    const { el, resolve } = renderEditable(doc, 0);

    const runs = el.querySelectorAll(`[${EDIT_ATTR.run}]`);
    runs[0]!.after(el.ownerDocument.createElement('br'));

    const segs = reconcileParagraph(el, resolve)[0]!.segments;
    expect(segs.map((s) => ({ t: s.text, br: !!s.breakBefore }))).toEqual([
      { t: 'Hello ', br: false },
      { t: 'World', br: true },
    ]);
  });

  it('reads a format marker as an override on the wrapped run', () => {
    const doc = DocxDocument.load(buildDocx());
    const { el, resolve, para } = renderEditable(doc, 0);

    const run = el.querySelector(`[${EDIT_ATTR.run}]`)!;
    const wrapper = el.ownerDocument.createElement('span');
    wrapper.setAttribute(EDIT_ATTR.fmt, JSON.stringify({ bold: true, sizePt: 18 }));
    run.replaceWith(wrapper);
    wrapper.appendChild(run);

    const segs = reconcileParagraph(el, resolve)[0]!.segments;
    expect(segs[0]!.format).toEqual({ bold: true, sizePt: 18 });
    expect(segs[0]!.src).toBe(para.runs[0]);
    expect(segs[1]!.format).toBeUndefined();
  });
});

describe('docx edit — viewer', () => {
  let container: HTMLElement;

  beforeEach(() => {
    document.head.innerHTML = '';
    document.body.innerHTML = '';
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  function mount(editable: boolean) {
    const doc = DocxDocument.load(buildDocx());
    const viewer = new DocxViewer(doc, container, { editable });
    return { doc, viewer };
  }

  it('paginates into page sheets when read-only', () => {
    mount(false);
    expect(container.querySelector(`.${FLOW_CLASS}`)).toBeNull();
    expect(container.querySelector(`[${EDIT_ATTR.para}]`)).toBeNull();
  });

  it('renders one continuous flow when editable', () => {
    mount(true);
    const flow = container.querySelector(`.${FLOW_CLASS}`);
    expect(flow).not.toBeNull();
    expect(flow!.querySelectorAll(`[${EDIT_ATTR.para}]`).length).toBeGreaterThan(0);
  });

  it('commits an edit when focus leaves the paragraph', () => {
    const { doc, viewer } = mount(true);
    const el = container.querySelector(`[${EDIT_ATTR.para}]`) as HTMLElement;
    el.querySelector(`[${EDIT_ATTR.run}]`)!.textContent = 'Farewell ';

    el.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));

    expect(textOf(paras(doc)[0]!)).toBe('Farewell World');
    expect(viewer.hasEdits).toBe(true);
    expect(container.textContent).toContain('Farewell');
  });

  it('formats the selected text', () => {
    const { doc, viewer } = mount(true);
    const el = container.querySelector(`[${EDIT_ATTR.para}]`) as HTMLElement;
    const text = el.querySelector(`[${EDIT_ATTR.run}]`)!.firstChild as Text;
    const range = document.createRange();
    range.setStart(text, 0);
    range.setEnd(text, 5);
    const sel = document.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);

    // 'Hello ' is plain and 'World' is already bold in the fixture, so bolding
    // the first five characters must leave exactly 'Hello' newly bold and the
    // text untouched. A DOCX paragraph is its own editable unit and carries no
    // body marker, so requiring one made every formatting command a no-op.
    expect(viewer.applyFormat({ bold: true })).toBe(true);
    const runs = paras(doc)[0]!.runs;
    expect(runs.map((r) => r.text).join('')).toBe('Hello World');
    expect(runs.filter((r) => r.bold).map((r) => r.text)).toEqual(['Hello', 'World']);
    expect(viewer.hasEdits).toBe(true);
  });

  it('undoes and redoes a committed edit', () => {
    const { doc, viewer } = mount(true);
    const el = container.querySelector(`[${EDIT_ATTR.para}]`) as HTMLElement;
    el.querySelector(`[${EDIT_ATTR.run}]`)!.textContent = 'Farewell ';
    el.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));
    expect(viewer.canUndo).toBe(true);

    viewer.undo();
    expect(textOf(paras(doc)[0]!)).toBe('Hello World');
    viewer.redo();
    expect(textOf(paras(doc)[0]!)).toBe('Farewell World');
  });

  it('keeps list numbering correct across edit-driven re-renders', () => {
    const { doc } = mount(true);
    const el = container.querySelector(`[${EDIT_ATTR.para}]`) as HTMLElement;
    el.querySelector(`[${EDIT_ATTR.run}]`)!.textContent = 'Changed ';
    el.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));

    expect(paras(doc).find((p) => p.listMarker)?.listMarker).toBe('1.');
  });

  it('does not navigate pages while the caret is in text', () => {
    const { viewer } = mount(true);
    const el = container.querySelector(`[${EDIT_ATTR.para}]`) as HTMLElement;

    const key = new KeyboardEvent('keydown', { key: 'PageDown', bubbles: true, cancelable: true });
    el.dispatchEvent(key);
    expect(key.defaultPrevented).toBe(false);

    const outside = new KeyboardEvent('keydown', { key: 'PageDown', bubbles: true, cancelable: true });
    container.dispatchEvent(outside);
    expect(outside.defaultPrevented).toBe(true);
    expect(viewer).toBeTruthy();
  });

  it('re-paginates when editing is switched back off', () => {
    const { viewer } = mount(false);
    const pagesBefore = viewer.count;

    viewer.setEditable(true);
    expect(container.querySelector(`.${FLOW_CLASS}`)).not.toBeNull();

    viewer.setEditable(false);
    expect(container.querySelector(`.${FLOW_CLASS}`)).toBeNull();
    expect(viewer.count).toBe(pagesBefore);
  });

  it('exports a document that reloads with the edit applied', () => {
    const { doc } = mount(true);
    const el = container.querySelector(`[${EDIT_ATTR.para}]`) as HTMLElement;
    el.querySelector(`[${EDIT_ATTR.run}]`)!.textContent = 'Farewell ';
    el.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));

    const reloaded = DocxDocument.load(doc.toBytes());
    expect(textOf(paras(reloaded)[0]!)).toBe('Farewell World');
  });
});
