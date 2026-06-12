# cbx-pptx-viewer-package

A frontend PowerPoint (`.pptx`) viewer — parse and render PowerPoint decks
directly in the browser, instead of converting them to PDF or images first.
Slides render to real HTML/CSS, so text is selectable, hyperlinks work, and the
output is accessible.

> Status: **viewer works, read-only.** Editing is not yet implemented — see
> [Roadmap](#roadmap).

## Why

Converting `.pptx` → PDF loses interactivity, needs a server-side tool
(LibreOffice/Office), and degrades fidelity. This project parses the Office Open
XML directly in the browser and renders each slide, resolving the same
layout/master/theme inheritance chain PowerPoint uses.

## How it works

A `.pptx` is a ZIP of XML parts (the OPC package). The pipeline is four stages:

```
read (OPC)  ->  parse (XML -> model)  ->  resolve (inheritance)  ->  render (DOM)
```

- **read** — unzip the package, resolve content types and relationships
- **parse** — turn slide/layout/master/theme XML into a render-agnostic model
- **resolve** — the accuracy core: every slide inherits geometry, colors, fonts,
  and text styles from its layout → master → theme. Scheme colors are mapped
  through the master's color map and transformed (lumMod/tint/shade/alpha).
- **render** — emit absolutely-positioned HTML/CSS; the viewer scales the whole
  slide to fit the viewport

All geometry is converted from EMU (914,400 per inch) to CSS pixels at 96 DPI.

## Repository layout

This is an npm-workspaces monorepo:

```
packages/
  core/     @pptx-viewer/core  — framework-agnostic parser + DOM renderer
  react/    @pptx-viewer/react — <PptxViewer /> component + useDeck() hook
app/        React app: upload a .pptx and view it (the main demo)
demo/       Vanilla-TS inspector: dumps a package's parts/content-types
```

### `@pptx-viewer/core`

The engine. Key modules in [`packages/core/src`](packages/core/src):

| Module | Responsibility |
| --- | --- |
| `opc/` | Unzip, content types, relationship resolution |
| `xml.ts` | Order-preserving, namespace-aware XML tree + helpers |
| `model.ts` | The render-agnostic intermediate representation |
| `resolve/` | Color, fill/stroke, placeholder & text-style inheritance |
| `geometry/` | Preset & custom geometry → SVG paths |
| `parse/` | Presentation, text bodies, tables, and the top-level deck loader |
| `render/dom.ts` | Model → HTML/CSS/SVG elements |
| `viewer/` | Navigation, fit-to-viewport scaling, keyboard |

## Usage

### React

```tsx
import { useState } from 'react';
import { PptxViewer } from '@pptx-viewer/react';

function App() {
  const [file, setFile] = useState<File | null>(null);
  return (
    <>
      <input type="file" accept=".pptx"
        onChange={(e) => setFile(e.target.files?.[0] ?? null)} />
      {file && <PptxViewer src={file} />}
    </>
  );
}
```

`src` accepts a `File`, `ArrayBuffer`, or `Uint8Array`. The viewer **self-sizes
to the slide's true aspect ratio** (fills the available width, height follows) —
you don't pass a fixed size or fit mode. It shows a built-in prev/next toolbar
(`toolbar={false}` to hide) and exposes `next()` / `prev()` / `goTo()` via a ref.

### Framework-agnostic core

```ts
import { loadPptx, createViewer } from '@pptx-viewer/core';

const deck = loadPptx(arrayBuffer);
const viewer = createViewer(deck, document.getElementById('stage')!);
viewer.next();
viewer.goTo(3);

// when done — frees the image object URLs
viewer.destroy();
deck.dispose();
```

## Development

Requires Node 18+ (developed on Node 26).

```bash
npm install

npm run dev            # run the React app (app/) at http://localhost:5173
npm run dev:inspector  # run the vanilla part-inspector (demo/)

npm test               # run @pptx-viewer/core unit tests (vitest)
npm run typecheck      # typecheck all workspaces
npm run build          # build the core + react libraries
npm run build:app      # production build of the React app
```

The app and demo are aliased directly to the packages' TypeScript source, so
there's no build step while developing — edits to `core`/`react` hot-reload.

## What's supported

- Slides composited from **master → layout → slide** (logos, decorations,
  backgrounds inherited from layouts/masters)
- **Text**: paragraphs and runs; bold/italic/underline/strike; font size, family
  (incl. theme major/minor), color; alignment; bullets (char + auto-number);
  indents, line/paragraph spacing; vertical anchor; line breaks; fields;
  autofit shrink; hyperlinks
- **Color**: theme scheme colors via the master color map, with
  lumMod/lumOff/tint/shade/alpha modifiers
- **Fills**: solid, linear/radial gradient, image; **outlines** with dashes
- **Shapes**: common preset geometries (rect, ellipse, rounded rect, triangle,
  diamond, arrows, polygons…) + **custom geometry** paths
- **Pictures**: with cropping and clipping to a non-rectangular shape
- **Groups** with nested coordinate transforms; rotation / flip
- **Tables**: grid, row/column spans, per-cell fill/borders/text
- Navigation: prev/next/goto, keyboard, fit-to-viewport scaling

## Known limitations

- **Read-only** — no editing yet
- **Charts** and **SmartArt** render as a labeled placeholder (not drawn)
- Only the common preset geometries are exact; the rest fall back to a rectangle
- Image *fills* inside shapes don't yet apply `srcRect` cropping (standalone
  pictures do)
- Effects (shadow / glow / reflection), slide transitions, and animations are
  not rendered
- `.ppt` (legacy binary format) is out of scope — `.pptx` only
- `clip-path: path()` requires a modern evergreen browser

## Roadmap

- Editing: text-content editing with round-trip export to `.pptx` (tracked back
  to the source XML), then shape move / resize / restyle
- Charts via the chart part data; SmartArt via its fallback drawing
- More preset geometries and effect mapping (CSS shadow / filter)
- Thumbnail strip, speaker-notes panel, fullscreen
- Optional visual-diff testing against reference renders

## License

TBD.
