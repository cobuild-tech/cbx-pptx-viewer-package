/**
 * @vitest-environment jsdom
 *
 * Every tool the toolbar offers, applied to real text, checked three ways:
 *
 *   1. the **deck** — the edit reached the XML and survived the re-parse;
 *   2. the **DOM** — the re-render shows it, which is what the user judges;
 *   3. the **toolbar** — the state reported back matches what was applied, so
 *      the button reflects the text the caret is in.
 *
 * A tool can pass one and fail another (a bullet wrote correctly but drew in
 * the wrong colour; a size applied but the field kept the old number), so the
 * matrix below deliberately asserts all three for each one.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { Deck } from '../src/pptx/deck/deck.js';
import { RelType } from '../src/pptx/relTypes.js';
import { Viewer } from '../src/pptx/viewer/viewer.js';
import { EDIT_ATTR } from '../src/oxml/edit/attrs.js';
import type { ParaFormat, RunFormat } from '../src/oxml/edit/format.js';
import type { PresetShape, TextBody } from '../src/pptx/model.js';

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

/**
 * One shape, two paragraphs, no paragraph properties at all — so every tool has
 * to write what it needs rather than leaning on what the deck happened to have.
 * The runs state a size and colour, which is what a bullet must adopt.
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
    'ppt/slides/slide1.xml': strToU8(
      `<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="2000000"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/>
            <a:p><a:r><a:rPr sz="3600"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>First line</a:t></a:r></a:p>
            <a:p><a:r><a:rPr sz="3600"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Second line</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
      </p:spTree></p:cSld></p:sld>`,
    ),
  });
}

let container: HTMLElement;
let lastRun: RunFormat;
let lastPara: ParaFormat;

beforeEach(() => {
  document.body.innerHTML = '';
  container = document.createElement('div');
  document.body.appendChild(container);
  lastRun = {};
  lastPara = {};
});

function mount(): { deck: Deck; viewer: Viewer } {
  const deck = Deck.load(buildDeck());
  const viewer = new Viewer(deck, container, {
    editable: true,
    webFonts: false,
    filmstrip: false,
    onSelectionChange: (f) => {
      lastRun = f;
    },
    onParaSelectionChange: (f) => {
      lastPara = f;
    },
  });
  return { deck, viewer };
}

const body = (deck: Deck): TextBody => (deck.slides[0]!.shapes[0] as PresetShape).text!;
const openBody = () =>
  container.querySelector<HTMLElement>(`[${EDIT_ATTR.body}][contenteditable="true"]`)!;
const paraEls = () => openBody().querySelectorAll<HTMLElement>(`[${EDIT_ATTR.para}]`);

/** Open the text and select the whole of paragraph `index`. */
function selectParagraph(viewer: Viewer, index = 0): void {
  const shape = (viewer as unknown as { deck: Deck }).deck.slides[0]!.shapes[0]!;
  if (!viewer.isEditingText) expect(viewer.editText(shape)).toBe(true);
  const run = paraEls()[index]!.querySelector(`[${EDIT_ATTR.run}]`)!;
  const range = document.createRange();
  range.selectNodeContents(run);
  const sel = document.getSelection()!;
  sel.removeAllRanges();
  sel.addRange(range);
}

/** Put a bare caret in paragraph `index` — what a paragraph tool works from. */
function caretIn(viewer: Viewer, index = 0): void {
  selectParagraph(viewer, index);
  document.getSelection()!.collapseToStart();
}

describe('pptx tools — character formatting', () => {
  const cases: Array<{
    name: string;
    format: RunFormat;
    /** What the rendered run must show. */
    css: (el: HTMLElement) => unknown;
    want: unknown;
    /** What the toolbar must report back. */
    reported: (f: RunFormat) => unknown;
  }> = [
    {
      name: 'bold',
      format: { bold: true },
      css: (el) => el.style.fontWeight,
      want: '700',
      reported: (f) => f.bold,
    },
    {
      name: 'italic',
      format: { italic: true },
      css: (el) => el.style.fontStyle,
      want: 'italic',
      reported: (f) => f.italic,
    },
    {
      name: 'underline',
      format: { underline: true },
      css: (el) => el.style.textDecoration,
      want: 'underline',
      reported: (f) => f.underline,
    },
    {
      name: 'strikethrough',
      format: { strike: true },
      css: (el) => el.style.textDecoration,
      want: 'line-through',
      reported: (f) => f.strike,
    },
    {
      name: 'size',
      format: { sizePt: 12 },
      css: (el) => el.style.fontSize,
      want: '16px',
      reported: (f) => f.sizePt,
    },
    {
      name: 'colour',
      format: { colorHex: '00FF00' },
      css: (el) => el.style.color,
      want: 'rgb(0, 255, 0)',
      reported: (f) => f.colorHex,
    },
    {
      name: 'typeface',
      format: { font: 'Georgia' },
      css: (el) => el.style.fontFamily,
      want: '"Georgia", Arial, Helvetica, sans-serif',
      reported: (f) => f.font,
    },
  ];

  for (const c of cases) {
    it(`applies ${c.name} to the selected text, and reports it back`, () => {
      const { deck, viewer } = mount();
      selectParagraph(viewer);

      expect(viewer.applyFormat(c.format)).toBe(true);

      // 1. the deck
      const runs = body(deck).paragraphs[0]!.runs;
      const key = Object.keys(c.format)[0] as keyof RunFormat;
      const expected = key === 'colorHex' ? '00FF00' : c.format[key];
      const got = key === 'colorHex' ? runs[0]!.color?.hex : (runs[0] as never)[key];
      expect(got).toEqual(expected);
      // 2. the DOM
      expect(c.css(openBody().querySelector<HTMLElement>(`[${EDIT_ATTR.run}]`)!)).toBe(c.want);
      // 3. the toolbar
      expect(c.reported(lastRun)).toEqual(expected);
      viewer.destroy();
    });
  }
});

describe('pptx tools — paragraph formatting', () => {
  it('bullets the paragraph, drawn as generated decoration', () => {
    const { deck, viewer } = mount();
    caretIn(viewer);

    expect(viewer.toggleList('bullet')).toBe(true);

    expect(body(deck).paragraphs[0]!.bullet).toMatchObject({ type: 'char', char: '•' });
    const glyph = paraEls()[0]!.querySelector<HTMLElement>('span:not([data-cbx-run])')!;
    expect(glyph.textContent).toContain('•');
    expect(glyph.getAttribute('contenteditable')).toBe('false');
    expect(lastPara.list).toBe('bullet');
    // The paragraph beside it is untouched.
    expect(body(deck).paragraphs[1]!.bullet).toBeUndefined();
    viewer.destroy();
  });

  it('numbers the paragraph', () => {
    const { deck, viewer } = mount();
    caretIn(viewer);
    expect(viewer.toggleList('number')).toBe(true);
    expect(body(deck).paragraphs[0]!.bullet).toMatchObject({ type: 'number' });
    expect(paraEls()[0]!.textContent).toContain('1.');
    expect(lastPara.list).toBe('number');
    viewer.destroy();
  });

  it('takes the bullet glyph’s size, colour and typeface from the text it marks', () => {
    const { viewer } = mount();
    caretIn(viewer);
    viewer.toggleList('bullet');

    // PowerPoint's default is buClrTx/buSzTx: the marker follows the first run
    // rather than the paragraph's ambient style. The text here is 36pt red, so
    // a bullet drawn at the box default in black is the visible bug.
    const glyph = paraEls()[0]!.querySelector<HTMLElement>('span:not([data-cbx-run])')!;
    expect(glyph.style.color).toBe('rgb(255, 0, 0)');
    expect(glyph.style.fontSize).toBe('48px');
    viewer.destroy();
  });

  it('numbers follow the text’s size and colour too', () => {
    const { viewer } = mount();
    caretIn(viewer);
    viewer.toggleList('number');
    const glyph = paraEls()[0]!.querySelector<HTMLElement>('span:not([data-cbx-run])')!;
    expect(glyph.style.color).toBe('rgb(255, 0, 0)');
    expect(glyph.style.fontSize).toBe('48px');
    viewer.destroy();
  });

  it('indents and outdents through the levels', () => {
    const { deck, viewer } = mount();
    caretIn(viewer);

    expect(viewer.indentSelection(1)).toBe(true);
    expect(body(deck).paragraphs[0]!.level).toBe(1);
    expect(lastPara.level).toBe(1);

    caretIn(viewer);
    expect(viewer.indentSelection(-1)).toBe(true);
    expect(body(deck).paragraphs[0]!.level).toBe(0);
    viewer.destroy();
  });

  const ALIGNMENTS: Array<[NonNullable<ParaFormat['align']>, string, string]> = [
    ['left', 'l', 'left'],
    ['center', 'ctr', 'center'],
    ['right', 'r', 'right'],
    ['justify', 'just', 'justify'],
  ];

  for (const [value, model, css] of ALIGNMENTS) {
    it(`aligns ${value}`, () => {
      const { deck, viewer } = mount();
      caretIn(viewer);

      expect(viewer.applyParaFormat({ align: value })).toBe(true);

      expect(body(deck).paragraphs[0]!.align).toBe(model);
      expect(paraEls()[0]!.style.textAlign).toBe(css);
      expect(lastPara.align).toBe(value);
      viewer.destroy();
    });
  }

  it('sets line spacing', () => {
    const { deck, viewer } = mount();
    caretIn(viewer);

    expect(viewer.applyParaFormat({ lineSpacingPct: 1.5 })).toBe(true);

    expect(body(deck).paragraphs[0]!.lineSpacingPct).toBe(1.5);
    // 1.5 x PowerPoint's single-spacing baseline.
    expect(paraEls()[0]!.style.lineHeight).toBe(String(1.5 * 1.2));
    expect(lastPara.lineSpacingPct).toBe(1.5);
    viewer.destroy();
  });

  it('sets space before the paragraph', () => {
    const { deck, viewer } = mount();
    caretIn(viewer);

    expect(viewer.applyParaFormat({ spaceBeforePt: 12 })).toBe(true);

    expect(body(deck).paragraphs[0]!.spaceBeforePt).toBe(12);
    expect(paraEls()[0]!.style.marginTop).toBe('16px');
    expect(lastPara.spaceBeforePt).toBe(12);
    viewer.destroy();
  });

  it('sets space after the paragraph', () => {
    const { deck, viewer } = mount();
    caretIn(viewer);

    expect(viewer.applyParaFormat({ spaceAfterPt: 6 })).toBe(true);

    expect(body(deck).paragraphs[0]!.spaceAfterPt).toBe(6);
    expect(paraEls()[0]!.style.marginBottom).toBe('8px');
    expect(lastPara.spaceAfterPt).toBe(6);
    viewer.destroy();
  });
});

describe('pptx tools — one after another', () => {
  it('keeps every earlier tool when the next one is applied', () => {
    const { deck, viewer } = mount();

    selectParagraph(viewer);
    viewer.applyFormat({ bold: true });
    selectParagraph(viewer);
    viewer.applyFormat({ italic: true });
    selectParagraph(viewer);
    viewer.applyFormat({ sizePt: 20 });
    caretIn(viewer);
    viewer.toggleList('bullet');
    caretIn(viewer);
    viewer.applyParaFormat({ align: 'center' });

    const para = body(deck).paragraphs[0]!;
    expect(para.runs.map((r) => ({ b: r.bold, i: r.italic, sz: r.sizePt }))).toEqual([
      { b: true, i: true, sz: 20 },
    ]);
    expect(para.bullet).toMatchObject({ type: 'char' });
    expect(para.align).toBe('ctr');
    // Five commands, five undo steps, back to where it started.
    for (let i = 0; i < 5; i++) viewer.undo();
    const back = body(deck).paragraphs[0]!;
    expect(back.runs[0]!.bold).toBeUndefined();
    expect(back.bullet).toBeUndefined();
    expect(back.align).toBeUndefined();
    viewer.destroy();
  });
});
