import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { loadDocx, InMemoryVersionStore } from '../src/index.js';
import type { DocxDocument } from '../src/docx/document/document.js';
import type { DocxParagraph, DocxTable } from '../src/docx/model.js';
import { DocxRelType } from '../src/docx/relTypes.js';

/** A small but representative body: plain para, styled+bold para, and a 1×1 table. */
const BODY = [
  '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Title</w:t></w:r></w:p>',
  '<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
].join('');

function buildDocx(bodyInner: string): Uint8Array {
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${bodyInner}</w:body></w:document>`;
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>',
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="${DocxRelType.OfficeDocument}" Target="word/document.xml"/>` +
        '</Relationships>',
    ),
    'word/document.xml': strToU8(documentXml),
  };
  return zipSync(files);
}

const paragraphs = (doc: DocxDocument): DocxParagraph[] => {
  const out: DocxParagraph[] = [];
  for (const page of doc.pages) for (const b of page.elements) if (b.kind === 'paragraph') out.push(b);
  return out;
};
const textOf = (p: DocxParagraph) => p.runs.map((r) => r.text).join('');
const tableOf = (doc: DocxDocument): DocxTable | undefined => {
  for (const page of doc.pages) for (const b of page.elements) if (b.kind === 'table') return b;
  return undefined;
};

describe('docx editing — identity threading', () => {
  it('stamps nodeIds on paragraphs, runs, and cells', () => {
    const doc = loadDocx(buildDocx(BODY));
    const ps = paragraphs(doc);
    expect(ps[0]!.nodeId).toMatch(/^word\/document\.xml#/);
    expect(ps[0]!.runs[0]!.nodeId).toMatch(/^word\/document\.xml#/);
    const cell = tableOf(doc)!.rows[0]![0]!;
    expect(cell.nodeId).toBeDefined();
    expect(cell.content[0]!.runs[0]!.nodeId).toBeDefined();
  });
});

describe('docx editing — text', () => {
  it('replaces run text and round-trips through export', () => {
    const doc = loadDocx(buildDocx(BODY));
    doc.editRunText(paragraphs(doc)[0]!.runs[0]!.nodeId!, 'Goodbye world');
    expect(textOf(paragraphs(doc)[0]!)).toBe('Goodbye world');
    expect(textOf(paragraphs(loadDocx(doc.export()))[0]!)).toBe('Goodbye world');
  });

  it('preserves significant whitespace with xml:space', () => {
    const doc = loadDocx(buildDocx(BODY));
    doc.editRunText(paragraphs(doc)[0]!.runs[0]!.nodeId!, '  spaced  ');
    expect(textOf(paragraphs(loadDocx(doc.export()))[0]!)).toBe('  spaced  ');
  });
});

describe('docx editing — formatting', () => {
  it('toggles bold on, undoes, and redoes', () => {
    const doc = loadDocx(buildDocx(BODY));
    const id = paragraphs(doc)[0]!.runs[0]!.nodeId!;
    expect(paragraphs(doc)[0]!.runs[0]!.bold).toBeUndefined();
    doc.setRunProps(id, { bold: true });
    expect(paragraphs(doc)[0]!.runs[0]!.bold).toBe(true);
    doc.undo();
    expect(paragraphs(doc)[0]!.runs[0]!.bold).toBeUndefined();
    doc.redo();
    expect(paragraphs(doc)[0]!.runs[0]!.bold).toBe(true);
  });

  it('explicitly turns off inherited bold (survives export)', () => {
    const doc = loadDocx(buildDocx(BODY));
    expect(paragraphs(doc)[1]!.runs[0]!.bold).toBe(true);
    doc.setRunProps(paragraphs(doc)[1]!.runs[0]!.nodeId!, { bold: false });
    expect(paragraphs(loadDocx(doc.export()))[1]!.runs[0]!.bold).toBeUndefined();
  });

  it('sets color and size', () => {
    const doc = loadDocx(buildDocx(BODY));
    doc.setRunProps(paragraphs(doc)[0]!.runs[0]!.nodeId!, { color: 'FF0000', sizePt: 18 });
    const r = paragraphs(loadDocx(doc.export()))[0]!.runs[0]!;
    expect(r.color?.hex).toBe('FF0000');
    expect(r.sizePt).toBe(18);
  });
});

describe('docx editing — paragraph props', () => {
  it('sets alignment and style as valid OOXML', () => {
    const doc = loadDocx(buildDocx(BODY));
    doc.setParagraphProps(paragraphs(doc)[0]!.nodeId!, { align: 'center', styleName: 'Heading2' });
    const xml = doc.serializePart('word/document.xml')!;
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml).toContain('<w:pStyle w:val="Heading2"/>');
  });
});

describe('docx editing — structural', () => {
  it('inserts a paragraph after another and undoes', () => {
    const doc = loadDocx(buildDocx(BODY));
    const before = paragraphs(doc).length;
    doc.insertParagraphAfter(paragraphs(doc)[0]!.nodeId!);
    expect(paragraphs(doc).length).toBe(before + 1);
    expect(paragraphs(loadDocx(doc.export())).length).toBe(before + 1);
    doc.undo();
    expect(paragraphs(doc).length).toBe(before);
  });

  it('deletes a paragraph and restores it on undo', () => {
    const doc = loadDocx(buildDocx(BODY));
    const before = paragraphs(doc).length;
    doc.deleteNode(paragraphs(doc)[1]!.nodeId!);
    expect(paragraphs(doc).length).toBe(before - 1);
    expect(textOf(paragraphs(doc)[0]!)).toBe('Hello');
    doc.undo();
    expect(paragraphs(doc).length).toBe(before);
    expect(textOf(paragraphs(doc)[1]!)).toBe('Title');
  });

  it('inserts a blank table row and deletes it', () => {
    const doc = loadDocx(buildDocx(BODY));
    const cellId = tableOf(doc)!.rows[0]![0]!.nodeId!;
    doc.insertRowAfter(cellId);
    expect(tableOf(doc)!.rows.length).toBe(2);
    expect(textOf(tableOf(doc)!.rows[1]![0]!.content[0]!)).toBe('');
    doc.undo();
    expect(tableOf(doc)!.rows.length).toBe(1);
  });
});

describe('docx editing — export integrity', () => {
  it('an unedited export reopens with identical text', () => {
    const doc = loadDocx(buildDocx(BODY));
    const reloaded = loadDocx(doc.export());
    expect(paragraphs(reloaded).map(textOf)).toEqual(paragraphs(doc).map(textOf));
  });

  it('tracks edit/undo state', () => {
    const doc = loadDocx(buildDocx(BODY));
    expect(doc.canUndo).toBe(false);
    expect(doc.isEdited).toBe(false);
    doc.editRunText(paragraphs(doc)[0]!.runs[0]!.nodeId!, 'x');
    expect(doc.canUndo).toBe(true);
    expect(doc.isEdited).toBe(true);
    doc.undo();
    expect(doc.canUndo).toBe(false);
    expect(doc.canRedo).toBe(true);
  });

  it('fires onChange on edit', () => {
    const doc = loadDocx(buildDocx(BODY));
    let n = 0;
    const off = doc.onChange(() => { n++; });
    doc.editRunText(paragraphs(doc)[0]!.runs[0]!.nodeId!, 'x');
    doc.undo();
    off();
    doc.editRunText(paragraphs(doc)[0]!.runs[0]!.nodeId!, 'y');
    expect(n).toBe(2); // edit + undo (not the post-unsubscribe edit)
  });
});

describe('docx editing — character-level formatting (run split)', () => {
  it('splits a run and bolds only the selected middle, as one undo step', () => {
    const doc = loadDocx(buildDocx(BODY));
    const runId = paragraphs(doc)[0]!.runs[0]!.nodeId!; // "Hello"
    // Bold "ell" (offsets 1..4).
    doc.formatRunRange(runId, 1, 4, { bold: true });

    let runs = paragraphs(doc)[0]!.runs;
    expect(runs.map((r) => r.text)).toEqual(['H', 'ell', 'o']);
    expect(runs.map((r) => r.bold)).toEqual([undefined, true, undefined]);
    expect(textOf(paragraphs(doc)[0]!)).toBe('Hello'); // text unchanged

    // Survives export.
    expect(paragraphs(loadDocx(doc.export()))[0]!.runs.map((r) => r.text)).toEqual(['H', 'ell', 'o']);

    // One undo restores the single run.
    doc.undo();
    expect(paragraphs(doc)[0]!.runs.map((r) => r.text)).toEqual(['Hello']);
    expect(paragraphs(doc)[0]!.runs[0]!.bold).toBeUndefined();
  });

  it('whole-run range just sets props without splitting', () => {
    const doc = loadDocx(buildDocx(BODY));
    const runId = paragraphs(doc)[0]!.runs[0]!.nodeId!;
    doc.formatRunRange(runId, 0, 5, { italic: true });
    const runs = paragraphs(doc)[0]!.runs;
    expect(runs.length).toBe(1);
    expect(runs[0]!.italic).toBe(true);
  });
});

describe('docx versioning — save & restore', () => {
  it('saves a version, edits further, and restores the earlier state', async () => {
    const doc = loadDocx(buildDocx(BODY));
    doc.configureVersioning(new InMemoryVersionStore(), 'doc-1');

    doc.editRunText(paragraphs(doc)[0]!.runs[0]!.nodeId!, 'First');
    const v1 = await doc.saveVersion('first edit', 1000);
    expect((await doc.listVersions()).length).toBe(1);

    doc.editRunText(paragraphs(doc)[0]!.runs[0]!.nodeId!, 'Second');
    expect(textOf(paragraphs(doc)[0]!)).toBe('Second');

    await doc.restore(v1.id);
    expect(textOf(paragraphs(doc)[0]!)).toBe('First');
    // The restored state exports to a valid .docx that reopens with the restored text.
    expect(textOf(paragraphs(loadDocx(doc.export()))[0]!)).toBe('First');
  });

  it('restore reverts structural edits back to the baseline', async () => {
    const doc = loadDocx(buildDocx(BODY));
    doc.configureVersioning(new InMemoryVersionStore(), 'doc-2');
    const base = paragraphs(doc).length;
    const v0 = await doc.saveVersion('baseline', 1);

    doc.insertParagraphAfter(paragraphs(doc)[0]!.nodeId!);
    doc.insertParagraphAfter(paragraphs(doc)[0]!.nodeId!);
    expect(paragraphs(doc).length).toBe(base + 2);

    await doc.restore(v0.id);
    expect(paragraphs(doc).length).toBe(base);
  });

  it('tracks unsaved changes and de-duplicates identical saves', async () => {
    const doc = loadDocx(buildDocx(BODY));
    doc.configureVersioning(new InMemoryVersionStore(), 'doc-unsaved');

    expect(doc.hasUnsavedChanges).toBe(false);
    doc.editRunText(paragraphs(doc)[0]!.runs[0]!.nodeId!, 'Changed');
    expect(doc.hasUnsavedChanges).toBe(true);

    const v1 = await doc.saveVersion('v1', 1);
    expect(doc.hasUnsavedChanges).toBe(false); // saved → nothing pending

    // Saving again with no changes must not create a duplicate version.
    const v1again = await doc.saveVersion('v1-dup', 2);
    expect(v1again.id).toBe(v1.id);
    expect((await doc.listVersions()).length).toBe(1);

    // A new edit makes it dirty again.
    doc.editRunText(paragraphs(doc)[0]!.runs[0]!.nodeId!, 'Changed twice');
    expect(doc.hasUnsavedChanges).toBe(true);
    const v2 = await doc.saveVersion('v2', 3);
    expect(v2.id).not.toBe(v1.id);
    expect((await doc.listVersions()).length).toBe(2);
  });

  it('parents versions and records the op-log since the parent', async () => {
    const doc = loadDocx(buildDocx(BODY));
    const store = new InMemoryVersionStore();
    doc.configureVersioning(store, 'doc-3');

    doc.editRunText(paragraphs(doc)[0]!.runs[0]!.nodeId!, 'A');
    const v1 = await doc.saveVersion('v1', 1);
    doc.editRunText(paragraphs(doc)[0]!.runs[0]!.nodeId!, 'B');
    const v2 = await doc.saveVersion('v2', 2);

    expect(v2.parentId).toBe(v1.id);
    const payload = await store.load('doc-3', v2.id);
    expect(payload!.ops.length).toBe(1); // exactly one edit since v1
    expect(payload!.changedParts['word/document.xml']).toContain('B');
  });
});
