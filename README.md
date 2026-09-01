# pptx-viewer-monorepo

A frontend Office-document viewer and editor — parse, render and edit
**PowerPoint (`.pptx`)**, **Word (`.docx`)** and **Excel (`.xlsx`)** files
directly in the browser, instead of converting them to PDF or images first.
Documents render to real HTML/CSS, so text is selectable, hyperlinks work, and
the output is accessible.

Published to npm as [`@cobuildx.ai/office-viewer`](https://www.npmjs.com/package/@cobuildx.ai/office-viewer).

> Status: all three formats render to the DOM, and all three support opt-in
> in-place text/cell editing with undo/redo and export back to a valid Office
> file. See [What's supported](#whats-supported) and
> [Extending the package](#extending-the-package).

## Why

Converting Office files → PDF loses interactivity, needs a server-side tool
(LibreOffice/Office), and degrades fidelity. This project parses the Office Open
XML directly in the browser and renders each slide/page, resolving the same
layout/master/theme and style-inheritance chains the desktop apps use.

## How it works

`.pptx`, `.docx` and `.xlsx` are all ZIP packages of XML parts (the OPC
package). Each format runs the same shape of pipeline over a shared low-level
layer:

```
PPTX:  read (OPC)  ->  parse (XML -> model)  ->  resolve (inheritance)  ->  render (DOM)
DOCX:  read (OPC)  ->  parse (XML -> model)  ->  paginate (sections)    ->  render (DOM)
XLSX:  read (OPC)  ->  parse (XML -> model, lazily per sheet)           ->  render (DOM)
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
  - *XLSX*: a cell carries no inline formatting — `<c s="3">` resolves through
    `cellXfs` to a font, fill, border and number format.
- **render** — emit HTML/CSS/SVG. PPTX positions shapes absolutely and scales
  the whole slide to fit; DOCX flows pages in a scrollable stack; XLSX paints a
  grid with sticky headers, a formula bar and sheet tabs.
- **edit** (opt-in) — the DOM subtree is read back into the model on commit,
  written into the original XML, and only the parts that changed are
  re-serialized on export.

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
| `@cobuildx.ai/office-viewer` | framework-agnostic: `loadPptx`/`loadDocx`/`loadXlsx`, viewers, edit sessions, renderers, low-level OOXML building blocks |
| `@cobuildx.ai/office-viewer/react` | `<PptxViewer />`, `<DocxViewer />`, `<XlsxViewer />`, `<EditorToolbar />`, `useDeck`, `useDocument`, `useWorkbook` |

The engine itself is organized as a **shared low-level layer + per-format
feature slices**. Each format is self-contained and must not import another
format's code; only `oxml/` is shared. (See
[Extending the package](#extending-the-package): *docx and xlsx are siblings of
pptx and must never import pptx*.)

```
packages/core/src/
  index.ts          public API: loadPptx / loadDocx / loadXlsx, viewers, low-level exports
  react/            React entry point (@cobuildx.ai/office-viewer/react)
    PptxViewer.tsx  DocxViewer.tsx  XlsxViewer.tsx  EditorToolbar.tsx
    useDeck.ts      useDocument.ts  useWorkbook.ts

  oxml/             SHARED, format-agnostic
    package.ts      OPC: unzip, content types, relationship resolution
    xml.ts          order-preserving, namespace-aware XML tree + helpers
    units.ts        EMU / point / twip -> CSS pixel conversions
    edit/           editing primitives every format shares: DOM markers, the
                    RunFormat value, snapshot undo, selection, styles

  pptx/             PowerPoint feature slices
    deck/           top-level loader; presentation + slide/layout/master/theme
    slides/         slide composition (master -> layout -> slide)
    shapes/         fills, geometry, placeholders, shape props/render
    text/           text bodies, runs, text-style inheritance
    tables/  charts/  diagrams/ (SmartArt)  pictures/  effects/
    color.ts        scheme-color map + modifier transforms
    model.ts        render-agnostic intermediate representation
    render/         model -> HTML/CSS/SVG (dom.ts), font install, primitives
    viewer/         thumbnail rail, navigation, fit-to-viewport scaling, keyboard

  docx/             Word feature slices (same shape as pptx)
    document/       top-level loader; body parsing
    edit/           text editing + continuous-flow renderer
    paragraphs/     paragraphs + runs
    styles/  numbering/  tables/  images/
    model.ts        DOCX page/block/paragraph/table model
    render/         model -> HTML/CSS (dom.ts), pagination-aware
    viewer/         paginated scroll view + thumbnail strip

  xlsx/             Excel feature slices (same shape as pptx/docx)
    workbook/       top-level loader; sheet list, lazy per-sheet parse
    sheets/         worksheet parse (rows, cells, merges, cols) + ref utilities
    styles/         numFmts / fonts / fills / borders / cellXfs
    edit/           cell editing: values, worksheet write-back, style interning
    model.ts        XLSX sheet/row/cell model
    render/         model -> HTML grid + formula bar, selection, keyboard
    viewer/         sheet tabs, commit cycle, export
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
      {/* add `editable` to any viewer for in-place editing + export */}
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
import {
  loadPptx, createViewer,
  loadDocx, createDocxViewer,
  loadXlsx, createXlsxViewer,
} from '@cobuildx.ai/office-viewer';

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

// XLSX
const workbook = loadXlsx(arrayBuffer);
const sheetViewer = createXlsxViewer(workbook, document.getElementById('stage')!, {
  editable: true,
});
sheetViewer.applyFormat({ bold: true, fillHex: 'FFE600' }); // formats the selected cells
const blob = sheetViewer.exportBlob();                      // a valid .xlsx with the edits
sheetViewer.destroy();
workbook.dispose();
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

The React app lets you upload a `.pptx`/`.docx`/`.xlsx` and (for `.pptx`) view the same
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
- PowerPoint's layout: a scrolling thumbnail rail on the left, the current
  slide on the stage to its right (`filmstrip: false` for a bare stage)
- Navigation: prev/next/goto, click a thumbnail, keyboard, fit-to-viewport
  scaling

### DOCX

- Style cascade: run → paragraph → linked → docDefaults
- Paragraphs and runs with character formatting; list numbering
- Tables (grid, cell borders/fills/text)
- Inline images
- Header/footer bands (incl. page-number/STYLEREF fields and banner images)
- Section-aware **pagination** (page size + margins), scroll view + thumbnails

### XLSX

- Cell values and formulas, shared and inline strings, number formats
- Fonts, fills, borders, alignment and wrapping resolved through `cellXfs`
- Merged cells, column widths, row heights, hidden rows
- Grid with sticky row/column headers, formula bar and sheet tabs

### Editing (opt in with `editable`)

- **PPTX**: the slide's own text bodies, incl. table cells — split/merge
  paragraphs, character formatting, undo/redo, export. Layout/master text is
  read-only, since editing it would change every slide that shares it.
  Whole slides can be deleted (`viewer.deleteSlide()`), which also drops the
  slide's notes and every reference to it; a deck must keep one slide.
- **DOCX**: body text incl. table cells. Headers/footers, list markers and field
  results are read-only. Edit mode renders one continuous column and
  re-paginates on exit (pagination isn't stored in a `.docx` anyway).
- **XLSX**: cell values, formulas and cell formatting (font, fill, alignment,
  wrap, number format) with Excel-like keys — arrows, Enter, Tab, F2, Delete,
  shift-select, formula bar. Protected sheets, merged-over cells and
  array/shared-formula hosts are read-only.
- Export is non-destructive: only edited parts are re-serialized, everything
  else is re-zipped from its original bytes.

## Known limitations

- Editing covers text and cell content, plus deleting whole PPTX slides —
  shape geometry, images, charts, and inserting/deleting spreadsheet rows and
  columns are not editable
- Formulas are never evaluated; the workbook is flagged so Excel recalculates
  when it opens the exported file
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

The codebase is structured so new formats slot in without rewrites.
Conventions to follow:

- **Shared vs. format code.** Anything format-agnostic (OPC packaging, XML
  helpers, unit conversion) lives in `oxml/` and is the *only* shared layer.
  Each format (`pptx/`, `docx/`, `xlsx/`) is a self-contained slice.
- **Formats never import each other.** `docx/` must not import from `pptx/` and
  vice-versa. If two formats need the same thing, lift it into `oxml/`.
- **Adding a format**: add a sibling `src/<format>/` with its own
  `model.ts`, `relTypes.ts`, parse slices, `render/dom.ts`, and `viewer/`,
  mirroring the `pptx`/`docx`/`xlsx` shape; reuse `oxml/` for read/units/xml; then
  export `load<Format>` + `create<Format>Viewer` from `src/index.ts` and add a
  React wrapper + hook in `packages/react/src/`.
- **Render-agnostic model first.** Parsing produces a plain model; rendering is a
  separate pass. Keep new features on that boundary.
- **Fidelity is generic.** Fixes must work for any conformant file, never
  special-cased to a particular deck.
- **No PDF/LibreOffice step**, ever — render natively in the browser, including
  for any reference/diff comparison.

## Roadmap

- More PPTX preset geometries and effect mapping (CSS shadow / filter)
- `srcRect` cropping for image fills inside shapes
- Speaker-notes panel, fullscreen
- Richer DOCX coverage (footnotes, more field types)
- Optional visual-diff testing against reference renders
- XLSX: row/column insert & delete, charts, conditional formatting

## License

TBD.
