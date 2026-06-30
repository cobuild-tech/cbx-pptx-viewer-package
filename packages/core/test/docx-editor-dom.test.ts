// @vitest-environment jsdom
import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { loadDocx, renderDocxPage } from '../src/index.js';
import { DocxEditController } from '../src/docx/viewer/editing.js';
import { domRunText } from '../src/docx/edit/selection.js';

function buildDocx(bodyInner: string): Uint8Array {
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${bodyInner}</w:body></w:document>`;
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>',
    ),
    '_rels/.rels': strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>',
    ),
    'word/document.xml': strToU8(documentXml),
  });
}

const deps = { imageUrl: () => undefined };
const paraText = (doc: ReturnType<typeof loadDocx>) =>
  doc.pages[0]!.elements
    .filter((b): b is Extract<typeof b, { kind: 'paragraph' }> => b.kind === 'paragraph')
    .map((p) => p.runs.map((r) => r.text).join(''))
    .join('\n');

describe('docx editor DOM pipeline (jsdom)', () => {
  it('makes paragraphs editable and stamps run ids + original text', () => {
    const doc = loadDocx(buildDocx('<w:p><w:r><w:t>Hello</w:t></w:r></w:p>'));
    const container = document.createElement('div');
    document.body.appendChild(container);
    container.appendChild(renderDocxPage(doc.pages[0]!, deps));

    new DocxEditController(doc).attach(container);

    const para = container.querySelector('.docx-para') as HTMLElement;
    const run = container.querySelector('.docx-run[data-docx-id]') as HTMLElement;
    expect(para.contentEditable).toBe('true');
    expect(run.dataset.docxId).toMatch(/^word\/document\.xml#/);
    expect(run.dataset.origText).toBe('Hello');
  });

  it('commits edited paragraph text back into the model on focusout', () => {
    const doc = loadDocx(buildDocx('<w:p><w:r><w:t>Hello</w:t></w:r></w:p>'));
    const container = document.createElement('div');
    document.body.appendChild(container);
    container.appendChild(renderDocxPage(doc.pages[0]!, deps));
    new DocxEditController(doc).attach(container);

    // Simulate the user editing the run's text, then blurring the paragraph.
    (container.querySelector('.docx-run[data-docx-id]') as HTMLElement).textContent = 'Hello, world';
    (container.querySelector('.docx-para') as HTMLElement).dispatchEvent(
      new FocusEvent('focusout', { bubbles: true }),
    );

    expect(paraText(doc)).toBe('Hello, world');
    expect(doc.canUndo).toBe(true);
  });

  it('does not emit an edit when nothing changed', () => {
    const doc = loadDocx(buildDocx('<w:p><w:r><w:t>Unchanged</w:t></w:r></w:p>'));
    const container = document.createElement('div');
    document.body.appendChild(container);
    container.appendChild(renderDocxPage(doc.pages[0]!, deps));
    new DocxEditController(doc).attach(container);

    (container.querySelector('.docx-para') as HTMLElement).dispatchEvent(
      new FocusEvent('focusout', { bubbles: true }),
    );
    expect(doc.canUndo).toBe(false);
  });

  it('domRunText reconstructs <br> as \\n and tab spans as \\t', () => {
    const span = document.createElement('span');
    span.appendChild(document.createTextNode('a'));
    span.appendChild(document.createElement('br'));
    const tab = document.createElement('span');
    tab.className = 'docx-tab';
    span.appendChild(tab);
    span.appendChild(document.createTextNode('b'));
    expect(domRunText(span)).toBe('a\n\tb');
  });
});
