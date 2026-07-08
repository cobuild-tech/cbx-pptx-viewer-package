# pptx-viewer-monorepo

A frontend Office-document viewer — parse and render **PowerPoint (`.pptx`)**
and **Word (`.docx`)** files directly in the browser, instead of converting them
to PDF or images first. Documents render to real HTML/CSS, so text is
selectable, hyperlinks work, and the output is accessible.

Published to npm as [`@cobuildx.ai/office-viewer`](https://www.npmjs.com/package/@cobuildx.ai/office-viewer).

> Status: the **PPTX viewer is read-only**. The **DOCX viewer supports inline
> editing** (undo/redo, run/paragraph formatting, table row insert/delete,
> pluggable version snapshots, export back to `.docx`). See
> [Roadmap](#roadmap) and [Extending the package](#extending-the-package).

## Why

Converting Office files → PDF loses interactivity, needs a server-side tool
(LibreOffice/Office), and degrades fidelity. This project parses the Office Open
XML directly in the browser and renders each slide/page, resolving the same
layout/master/theme and style-inheritance chains the desktop apps use.

## How it works

Both `.pptx` and `.docx` are ZIP packages of XML parts (the OPC package). Each
format runs the same shape of pipeline over a shared low-level layer:

```
PPTX:  read (OPC)  ->  parse (XML -> model)  ->  resolve (inheritance)  ->  render (DOM)
DOCX:  read (OPC)  ->  parse (XML -> model)  ->  paginate (sections)    ->  render (DOM)
```

- **read** — unzip the package, resolve content types and relationships (shared
  `oxml/` layer)
- **parse** — turn the format's XML parts into a render-agnostic model
- **resolve / paginate** — the accuracy core.
  - *PPTX*: every slide inherits geometry, colors, fonts, and text styles from
    its layout → master → theme. Scheme colors are mapped through the master's
    color map and transformed (lumMod/lumOff/tint/shade/alpha).
  - *DOCX*: styles cascade (run → paragraph → linked → docDefaults), then blocks
    are measured off-screen and flowed into pages by section page-size/margins.
- **render** — emit HTML/CSS/SVG. PPTX positions shapes absolutely and scales
  the whole slide to fit; DOCX flows pages in a scrollable stack.

All geometry is converted from EMU (914,400 per inch) to CSS pixels at 96 DPI
via the shared `oxml/units` helpers.

## Repository layout

This is an npm-workspaces monorepo:

```
packages/
  core/     @cobuildx.ai/office-viewer — the published package (source lives here)
app/        React app: upload a .pptx/.docx and view it (installs the published
            package from npm — the main demo of "does this work for a real consumer")
demo/       Vanilla-TS inspector: dumps a package's parts/content-types (runs
            directly against packages/core's TypeScript source, no build step)
```

`@cobuildx.ai/office-viewer` ships as a single npm package with two entry
points, so non-React consumers never pull in React:

| Entry | Contents |
| --- | --- |
| `@cobuildx.ai/office-viewer` | framework-agnostic: `loadPptx`/`loadDocx`, viewers, renderers, low-level OOXML building blocks |
| `@cobuildx.ai/office-viewer/react` | `<PptxViewer />`, `<DocxViewer />`, `<DocxEditorToolbar />`, `useDeck`, `useDocument` |

The engine itself is organized as a **shared low-level layer + per-format
feature slices**. Each format is self-contained and must not import another
format's code; only `oxml/` is shared. (See
[Extending the package](#extending-the-package): *docx is a sibling of pptx and
must never import pptx*.)

```
packages/core/src/
  index.ts          public API: loadPptx / loadDocx, viewers, low-level exports
  react/            React entry point (@cobuildx.ai/office-viewer/react)
    PptxViewer.tsx  DocxViewer.tsx  DocxEditorToolbar.tsx  useDeck.ts  useDocument.ts

  oxml/             SHARED, format-agnostic
    package.ts      OPC: unzip, content types, relationship resolution
    xml.ts          order-preserving, namespace-aware XML tree + helpers
    units.ts        EMU / point / twip -> CSS pixel conversions

  pptx/             PowerPoint feature slices
    deck/           top-level loader; presentation + slide/layout/master/theme
    slides/         slide composition (master -> layout -> slide)
    shapes/         fills, geometry, placeholders, shape props/render
    text/           text bodies, runs, text-style inheritance
    tables/  charts/  diagrams/ (SmartArt)  pictures/  effects/
    color.ts        scheme-color map + modifier transforms
    model.ts        render-agnostic intermediate representation
    render/         model -> HTML/CSS/SVG (dom.ts), font install, primitives
    viewer/         navigation, fit-to-viewport scaling, keyboard

  docx/             Word feature slices (same shape as pptx)
    document/       top-level loader; body parsing
    paragraphs/     paragraphs + runs
    styles/  numbering/  tables/  images/
    edit/           edit ops, node addressing, version snapshots
    model.ts        DOCX page/block/paragraph/table model
    render/         model -> HTML/CSS (dom.ts), pagination-aware
    viewer/         paginated scroll view + thumbnail strip + inline editing
```

## Usage

## Install

```bash
npm install @cobuildx.ai/office-viewer
```

React bindings additionally require `react` and `react-dom` (>=18) — they're
an optional peer dependency, so plain-JS consumers never pull in React.

### React

```tsx
import { useState } from 'react';
import { PptxViewer, DocxViewer } from '@cobuildx.ai/office-viewer/react';

function App() {
  const [file, setFile] = useState<File | null>(null);
  const isDocx = file?.name.toLowerCase().endsWith('.docx');
  return (
    <>
      <input type="file" accept=".pptx,.docx"
        onChange={(e) => setFile(e.target.files?.[0] ?? null)} />
      {file && (isDocx ? <DocxViewer src={file} /> : <PptxViewer src={file} />)}
    </>
  );
}
```

`src` accepts a `File`, `ArrayBuffer`, or `Uint8Array`.

- **`PptxViewer`** by default **fits the whole slide to its container and
  centres it** — like PowerPoint's slideshow. It shows a built-in prev/next
  toolbar (`toolbar={false}` to hide) and exposes `next()` / `prev()` /
  `goTo()` via a ref.
- **`DocxViewer`** flows pages in a vertical scroll stack with a thumbnail
  strip, and exposes the same ref controls for page navigation.

### Framework-agnostic core

```ts
import { loadPptx, createViewer, loadDocx, createDocxViewer } from '@cobuildx.ai/office-viewer';

// PPTX
const deck = loadPptx(arrayBuffer);
const viewer = createViewer(deck, document.getElementById('stage')!);
viewer.next();
viewer.goTo(3);
viewer.destroy();   // frees image object URLs
deck.dispose();

// DOCX
const doc = loadDocx(arrayBuffer);
const docViewer = createDocxViewer(doc, document.getElementById('stage')!, { fit: 'width' });
docViewer.goTo(2);
docViewer.destroy();
doc.dispose();
```

PPTX sizing defaults to `fit: 'contain'` (fit the whole slide and centre it).
Pass `fit: 'width'` for the embedded-document style that fills the width and
lets the page scroll vertically.

## Development

Requires Node 18+ (developed on Node 26).

```bash
npm install

npm run dev            # run the React app (app/) at http://localhost:5173
npm run dev:inspector  # run the vanilla part-inspector (demo/)

npm test               # run @cobuildx.ai/office-viewer unit tests (vitest)
npm run typecheck      # typecheck the package + app
npm run build          # build the published package (packages/core)
npm run build:app      # production build of the React app
```

`demo/` is aliased directly to `packages/core`'s TypeScript source, so there's
no build step while developing it — edits hot-reload. `app/` is intentionally
**not** aliased: it installs `@cobuildx.ai/office-viewer` from the real npm
registry as an ordinary dependency, so it exercises the actual published
package rather than local source. After changing `packages/core`, run
`npm run build` there, bump its version, and republish before `app/` will pick
up the change (or `npm link` locally if you want to iterate faster).

### Comparing renderers

The React app lets you upload a `.pptx`/`.docx` and (for `.pptx`) view the same
file through this package or three third-party libraries, to compare fidelity:

| Option | Package | Approach |
| --- | --- | --- |
| cbx-ppt-viewer (default) | `@cobuildx.ai/office-viewer/react` | HTML/CSS DOM, this repo's parser |
| pptx-react-viewer | [`pptx-react-viewer`](https://www.npmjs.com/package/pptx-react-viewer) | HTML/CSS, full editor (heavy deps) |
| pptxviewjs | [`pptxviewjs`](https://www.npmjs.com/package/pptxviewjs) | `<canvas>` renderer |
| @cyntler/react-doc-viewer | [`@cyntler/react-doc-viewer`](https://www.npmjs.com/package/@cyntler/react-doc-viewer) | embeds the MS Office Online viewer |

Each third-party renderer is lazy-loaded, so its bundle only downloads when
selected. Note `@cyntler/react-doc-viewer` renders via the Office Online viewer,
which needs a **publicly reachable URL** — a locally-uploaded file (blob URL)
won't load there.

## What's supported

### PPTX

- Slides composited from **master → layout → slide** (logos, decorations,
  backgrounds inherited from layouts/masters)
- **Text**: paragraphs and runs; bold/italic/underline/strike; font size, family
  (incl. theme major/minor), color; alignment; bullets (char + auto-number);
  indents, line/paragraph spacing; vertical anchor; line breaks; fields;
  autofit shrink; hyperlinks
- **Color**: theme scheme colors via the master color map, with
  lumMod/lumOff/tint/shade/alpha modifiers
- **Fills**: solid, linear/radial gradient, image; **outlines** with dashes
- **Shapes**: ~106 preset geometries (rect, ellipse, rounded rect, triangle,
  diamond, arrows, polygons, callouts…) + **custom geometry** paths
- **Pictures**: with cropping and clipping to a non-rectangular shape
- **Groups** with nested coordinate transforms; rotation / flip
- **Tables**: grid, row/column spans, per-cell fill/borders/text
- **Charts**: bar/line/area/scatter/pie/doughnut drawn as static SVG (the cached
  snapshot PowerPoint shows when the workbook is detached)
- **SmartArt / diagrams**: cached `dsp:drawing` fast path, plus a data-model
  fallback that lays out nodes by layout family (cycle/process/list/hierarchy/
  pyramid) so the real content shows instead of a placeholder
- Navigation: prev/next/goto, keyboard, fit-to-viewport scaling

### DOCX

- Style cascade: run → paragraph → linked → docDefaults
- Paragraphs and runs with character formatting; list numbering
- Tables (grid, cell borders/fills/text)
- Inline images
- Section-aware **pagination** (page size + margins), scroll view + thumbnails
- **Inline WYSIWYG editing** (opt-in via `editable`): undo/redo, run/paragraph
  formatting, table row insert/delete, pluggable version snapshots
  (`DocxVersionStore`), export the edited document back to a `.docx` `Blob`

## Known limitations

- **PPTX is read-only** — no editing yet (the model is intentionally
  render-agnostic to make round-trip editing tractable later, as it already is
  for DOCX)
- PPTX: only the implemented preset geometries are exact; the rest fall back to
  a rectangle
- Image *fills* inside shapes don't yet apply `srcRect` cropping (standalone
  pictures do)
- Effects (shadow / glow / reflection), slide transitions, and animations are
  not rendered
- Charts render as a static snapshot, not an interactive chart
- `.ppt` / `.doc` (legacy binary formats) are out of scope — the OOXML zip
  formats only
- `clip-path: path()` requires a modern evergreen browser

## Extending the package

The codebase is structured so new formats and, eventually, an editor slot in
without rewrites. Conventions to follow:

- **Shared vs. format code.** Anything format-agnostic (OPC packaging, XML
  helpers, unit conversion) lives in `oxml/` and is the *only* shared layer.
  Each format (`pptx/`, `docx/`) is a self-contained slice.
- **Formats never import each other.** `docx/` must not import from `pptx/` and
  vice-versa. If two formats need the same thing, lift it into `oxml/`.
- **Adding a format** (e.g. `.xlsx`): add a sibling `src/<format>/` with its own
  `model.ts`, `relTypes.ts`, parse slices, `render/dom.ts`, and `viewer/`,
  mirroring the `pptx`/`docx` shape; reuse `oxml/` for read/units/xml; then
  export `load<Format>` + `create<Format>Viewer` from `src/index.ts` and add a
  React wrapper + hook in `packages/react/src/`.
- **Render-agnostic model first.** Parsing produces a plain model; rendering is a
  separate pass. Keep new features on that boundary — this is what makes an
  editor (model edit → DOM re-render → XML round-trip) feasible.
- **Fidelity is generic.** Fixes must work for any conformant file, never
  special-cased to a particular deck.
- **No PDF/LibreOffice step**, ever — render natively in the browser, including
  for any reference/diff comparison.

### Toward an editor

DOCX already proves the pattern out: edits mutate the render-agnostic model
via `applyOp`, the DOM re-renders, and `exportBlob()` serializes back to a real
`.docx`. Extending the same model → render → serialize loop to PPTX (text
first, then shape move/resize/restyle) is the main remaining piece.

## Roadmap

- PPTX editing: text-content editing with round-trip export, then shape
  move / resize / restyle
- More PPTX preset geometries and effect mapping (CSS shadow / filter)
- `srcRect` cropping for image fills inside shapes
- Speaker-notes panel, fullscreen
- Richer DOCX coverage (headers/footers, footnotes, fields)
- Optional visual-diff testing against reference renders
- Further formats (e.g. `.xlsx`) as new sibling slices

## License

TBD.
