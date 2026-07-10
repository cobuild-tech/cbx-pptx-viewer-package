# @cobuildx.ai/office-viewer

Renders `.pptx` and `.docx` files directly in the browser — no server-side conversion, no PDF, no LibreOffice. Parses the OOXML package and paints slides/pages straight to the DOM.

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

`src` accepts a `File`, `ArrayBuffer`, or `Uint8Array`.

### Components

**`<PptxViewer src toolbar? className? style? onLoad? onError? onSlideChange? />`**
Renders a `.pptx` deck with built-in prev/next navigation. Pass a ref to get `{ next(), prev(), goTo(index), deck }`.

**`<DocxViewer src toolbar? className? style? onLoad? onError? onPageChange? />`**
Renders a `.docx` document, paginated, with a thumbnail strip. Pass a ref to get `{ next(), prev(), goTo(index), doc }`.

### Hooks

If you want to drive your own UI instead of the built-in components:

```tsx
import { useDeck } from '@cobuildx.ai/office-viewer/react';

const { deck, loading, error } = useDeck(file); // re-loads on `file` change, disposes on unmount
```

`useDocument(src)` is the `.docx` equivalent, returning `{ doc, loading, error }`.

## Framework-agnostic usage

```ts
import { loadPptx, createViewer } from '@cobuildx.ai/office-viewer';

const deck = loadPptx(arrayBuffer);
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

Always call `dispose()` on the deck/document when you're done with it — it revokes object URLs created for embedded media.

## What's supported

Both formats are **read-only viewers**.

- PPTX: ~106 preset shapes, charts, tables, pictures, text with autofit, gradients (including radial focus/path gradients); SmartArt is not yet laid out (data model only).
- DOCX: paragraphs, runs, tables (styles + conditional formatting), images, numbering/lists, styles, headers/footers, section-aware pagination.

## Package layout

| Export | Contents |
|---|---|
| `@cobuildx.ai/office-viewer` | Deck/document loading, viewers, renderers, and low-level OOXML building blocks (`OpcPackage`, XML helpers) — no React dependency. |
| `@cobuildx.ai/office-viewer/react` | `PptxViewer`, `DocxViewer`, `useDeck`, `useDocument`. |

