# @cobuildx.ai/office-viewer

Renders — and edits — `.pptx`, `.docx` and `.xlsx` files directly in the browser. No server-side conversion, no PDF, no LibreOffice: the OOXML package is parsed and painted straight to the DOM, and edits are written back into the original package so export returns a real Office file.

- **Framework-agnostic core** at the package root — usable from plain JS, Vue, Svelte, etc.
- **React bindings** at `@cobuildx.ai/office-viewer/react` — components and hooks. React is an optional peer dependency: if you never import the `/react` entry, your bundle never sees React.

## 🚀 Updates

We ship updates **every week** — rendering fidelity fixes, new features, and broader format coverage. Don't miss out: check back regularly and keep your dependency up to date.

```
npm update @cobuildx.ai/office-viewer
```

See the [latest version on npm](https://www.npmjs.com/package/@cobuildx.ai/office-viewer) for what's new.

## Install

```
npm install @cobuildx.ai/office-viewer
```

React bindings additionally require `react` and `react-dom` (>=18) in your project.

## React usage

```tsx
import { PptxViewer } from '@cobuildx.ai/office-viewer/react';

function Slide({ file }: { file: File }) {
  return <PptxViewer src={file} style={{ height: '100vh' }} />;
}
```

```tsx
import { DocxViewer } from '@cobuildx.ai/office-viewer/react';

function Doc({ file }: { file: File }) {
  return <DocxViewer src={file} style={{ height: '100vh' }} />;
}
```

```tsx
import { XlsxViewer } from '@cobuildx.ai/office-viewer/react';

function Sheet({ file }: { file: File }) {
  return <XlsxViewer src={file} style={{ height: '100vh' }} />;
}
```

`src` accepts a `File`, `ArrayBuffer`, or `Uint8Array`.

### Editing

Every viewer takes an `editable` prop. It turns on in-place WYSIWYG editing with
a formatting toolbar, undo/redo, and a download button that re-zips the original
file with the edits applied:

```tsx
<PptxViewer src={file} editable />
<DocxViewer src={file} editable />
<XlsxViewer src={file} editable />
```

Export is non-destructive: only the parts you actually changed are re-serialized,
every other part is emitted from its original bytes.

### Components

**`<PptxViewer src toolbar? filmstrip? filmstripWidth? editable? textBoxOutline? className? style? onLoad? onError? onSlideChange? onEdit? />`**
Renders a `.pptx` deck with built-in prev/next navigation. Pass a ref to get `{ next(), prev(), goTo(index), applyFormat(), undo(), redo(), exportBlob(), deck }`.

**`<DocxViewer src toolbar? editable? editorToolbar? className? style? onLoad? onError? onPageChange? onEdit? />`**
Renders a `.docx` document, paginated, with a thumbnail strip. Pass a ref to get `{ next(), prev(), goTo(index), applyFormat(), undo(), redo(), exportBlob(), doc }`.

**`<XlsxViewer src initialSheet? editable? editorToolbar? className? style? onLoad? onError? onSheetChange? onEdit? />`**
Renders a `.xlsx` workbook as a grid with a formula bar and sheet tabs. Pass a ref to get `{ goToSheet(), applyFormat(), commit(), undo(), redo(), exportBlob(), workbook }`.

### Hooks

If you want to drive your own UI instead of the built-in components:

```tsx
import { useDeck } from '@cobuildx.ai/office-viewer/react';

const { deck, loading, error } = useDeck(file); // re-loads on `file` change, disposes on unmount
```

`useDocument(src)` is the `.docx` equivalent, returning `{ doc, loading, error }`, and
`useWorkbook(src)` the `.xlsx` one, returning `{ workbook, loading, error }`.

## Framework-agnostic usage

```ts
import { loadPptx, createViewer } from '@cobuildx.ai/office-viewer';

const deck = loadPptx(arrayBuffer);
// Thumbnail rail on the left, current slide on the right (filmstrip: false to drop it).
const viewer = createViewer(deck, containerEl, { fit: 'contain' });
viewer.next();
viewer.goTo(3);
// when done:
viewer.destroy();
deck.dispose();
```

```ts
import { loadDocx, createDocxViewer } from '@cobuildx.ai/office-viewer';

const doc = loadDocx(arrayBuffer);
const viewer = createDocxViewer(doc, containerEl, { fit: 'width' });
viewer.next();
viewer.goTo(2);
// when done:
viewer.destroy();
doc.dispose();
```

```ts
import { loadXlsx, createXlsxViewer } from '@cobuildx.ai/office-viewer';

const workbook = loadXlsx(arrayBuffer);
const viewer = createXlsxViewer(workbook, containerEl, { editable: true });
viewer.goToSheet('Summary');
viewer.applyFormat({ bold: true, fillHex: 'FFE600' }); // formats the selected cells
const blob = viewer.exportBlob();                      // a valid .xlsx with the edits
// when done:
viewer.destroy();
workbook.dispose();
```

Always call `dispose()` on the deck/document/workbook when you're done with it — it revokes object URLs created for embedded media.

## What's supported

**Rendering**

- PPTX: ~106 preset shapes, charts, tables, pictures, text with autofit, gradients (including radial focus/path gradients); SmartArt is laid out via a cached fast path with a data-model fallback. Effects, transitions and animations are not rendered.
- DOCX: paragraphs, runs, tables (styles + conditional formatting), images, numbering/lists, styles, headers/footers, section-aware pagination.
- XLSX: cell values and formulas, number formats, fonts/fills/borders/alignment, merged cells, column widths and row heights, sheet tabs, formula bar.

**Editing** (`editable`)

- PPTX: the slide's own text bodies, including table cells — paragraph split/merge, character formatting, undo/redo, export. Text inherited from a layout or master renders but stays read-only, since editing it would change every slide that shares it. Whole slides can be deleted with `viewer.deleteSlide(index?)` (the deck must keep at least one). Shape geometry, images and charts are not editable.
- DOCX: body text including table cells. Header and footer text, generated list markers and field results stay read-only. Edit mode renders the document as one continuous column instead of fixed pages (Word reflows on open, so pagination is a display concern only) and re-paginates on exit.
- XLSX: cell values, formulas and cell formatting (bold/italic/underline/strike, size, font, text colour, fill, alignment, wrap, number format), with Excel-like keyboard behaviour — arrows, Enter, Tab, F2, Delete, shift-select. Formulas are stored but not evaluated: the workbook is flagged so Excel recalculates on open. Cells on a protected sheet, cells covered by a merge, and array/shared-formula hosts stay read-only. Inserting or deleting rows and columns is not supported.

## Package layout

| Export | Contents |
|---|---|
| `@cobuildx.ai/office-viewer` | Deck/document loading, viewers, renderers, and low-level OOXML building blocks (`OpcPackage`, XML helpers) — no React dependency. |
| `@cobuildx.ai/office-viewer/react` | `PptxViewer`, `DocxViewer`, `XlsxViewer`, `EditorToolbar`, `useDeck`, `useDocument`, `useWorkbook`. |

