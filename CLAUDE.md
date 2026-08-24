# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

An npm-workspaces monorepo containing `@cobuildx.ai/office-viewer`, a published npm package that parses and renders `.pptx`/`.docx` files directly in the browser (real HTML/CSS/SVG output — no PDF conversion, no LibreOffice, ever, including for reference/diff comparisons).

```
packages/core/   the package itself (source, tests, build config)
app/             React app that installs @cobuildx.ai/office-viewer from the real npm
                 registry (NOT a workspace-linked dependency) — the main demo, also lets
                 you compare rendering fidelity against third-party pptx/docx viewers
demo/            Vanilla-TS inspector aliased directly to packages/core's TS source
                 (no build step) — dumps a package's parts/content-types
```

`app/` is deliberately excluded from the root `workspaces` array in `package.json` so it exercises the actual published package like any external consumer, not local source. After changing `packages/core`, you must `npm run build`, bump the version, and republish before `app/` picks up the change — there is no live-reload path from `packages/core` source into `app/`. `demo/` *is* aliased to local source and hot-reloads.

## Commands

Run from the repo root unless noted.

```bash
npm install                    # installs root + demo + packages/* workspaces (NOT app/)
npm --prefix app install       # app/ has its own node_modules/lockfile, install separately

npm run dev                    # app/ dev server (npm --prefix app run dev)
npm run dev:inspector          # demo/ dev server

npm test                       # vitest run, packages/core only
npm run typecheck              # tsc --noEmit for packages/core, then app/
npm run build                  # vite build for packages/core (produces dist/index.js + dist/react.js)
npm run build:app              # production build of app/
```

Inside `packages/core/`:

```bash
npx vitest run test/deck.test.ts   # run a single test file
npm run test:watch                 # vitest watch mode
```

Test files live in `packages/core/test/*.test.ts`; vitest runs in `environment: 'node'` with `globals: true` (see `packages/core/vite.config.ts`).

## Publishing

`packages/core/package.json` is the published unit. It builds two entry points via Vite lib mode (`src/index.ts` → `dist/index.js`, `src/react/index.ts` → `dist/react.js`), exposed via `exports: { ".": ..., "./react": ... }` so plain-JS consumers never pull in React (`react`/`react-dom` are optional peer dependencies, `peerDependenciesMeta.optional: true`).

To publish: bump `version` in `packages/core/package.json`, `npm run build`, then `npm publish --access public` from `packages/core/`. The npm account has 2FA; publishing needs either an OTP or a Classic **Automation**-type access token passed via a temporary `--userconfig` file (granular access tokens do not reliably bypass 2FA for publish even with write permission). Never commit a token — `.env` is gitignored for this reason.

## Architecture

Both `.pptx` and `.docx` are OPC (zip-of-XML) packages. Each format runs the same pipeline shape over one shared low-level layer:

```
PPTX:  read (OPC) -> parse (XML -> model) -> resolve (inheritance)  -> render (DOM)
DOCX:  read (OPC) -> parse (XML -> model) -> paginate (sections)    -> render (DOM)
```

- **read** (`oxml/package.ts`) — unzip, resolve content types + relationships.
- **parse** — turn XML parts into a render-agnostic model (`pptx/model.ts` / `docx/model.ts`).
- **resolve/paginate** — the fidelity core. PPTX slides inherit geometry/color/font/text styles through layout → master → theme, with scheme colors mapped through the master's color map and transformed (lumMod/lumOff/tint/shade/alpha). DOCX styles cascade run → paragraph → linked → docDefaults, then blocks are measured off-screen and flowed into pages by section page-size/margins.
- **render** — model → HTML/CSS/SVG. PPTX positions shapes absolutely and scales the whole slide to fit; DOCX flows pages in a scrollable stack. All geometry converts EMU (914,400/inch) → CSS px at 96 DPI via `oxml/units.ts`.

### Package layout (`packages/core/src/`)

```
index.ts          public API: loadPptx / loadDocx, viewers, low-level exports
react/            React entry point (@cobuildx.ai/office-viewer/react)
                  PptxViewer.tsx, DocxViewer.tsx, DocxEditorToolbar.tsx, useDeck.ts, useDocument.ts

oxml/             SHARED, format-agnostic — the only layer pptx/ and docx/ may both depend on
  package.ts      OPC: unzip, content types, relationship resolution
  xml.ts          order-preserving, namespace-aware XML tree + helpers
  units.ts        EMU / point / twip -> CSS pixel conversions

pptx/             PowerPoint feature slices
  deck/           top-level loader; presentation + slide/layout/master/theme
  edit/           text editing: source addressing, DOM reconciliation,
                  XML write-back, formatting, undo/redo
  slides/         slide composition (master -> layout -> slide)
  shapes/         fills, geometry, placeholders, shape props/render
  text/           text bodies, runs, text-style inheritance
  tables/ charts/ diagrams/ (SmartArt) pictures/ effects/
  color.ts        scheme-color map + modifier transforms
  model.ts        render-agnostic intermediate representation
  render/         model -> HTML/CSS/SVG (dom.ts), font install, primitives
  viewer/         navigation, fit-to-viewport scaling, keyboard

docx/             Word feature slices (same shape as pptx)
  document/       top-level loader; body parsing
  paragraphs/ styles/ numbering/ tables/ images/
  model.ts        DOCX page/block/paragraph/table model
  render/         model -> HTML/CSS (dom.ts), pagination-aware
  viewer/         paginated scroll view + thumbnail strip
```

**Hard rule: `pptx/` and `docx/` must never import from each other.** Anything both need belongs in `oxml/`, not duplicated or cross-imported.

### Status

- PPTX: renders ~106 preset shapes + custom geometry, charts (static SVG snapshot), tables, SmartArt (cached fast path + data-model layout fallback by family), gradients (incl. radial focus/path), autofit text. Effects (shadow/glow/reflection), transitions, and animations are not rendered. `.ppt` (legacy binary) is out of scope.
- PPTX **text editing** (opt in via `<PptxViewer editable>` / `createViewer(deck, el, { editable: true })`): inline WYSIWYG editing of the slide's own text bodies — including table cell text — with paragraph split/merge, a formatting toolbar (bold/italic/underline/strike, size, colour, typeface), undo/redo, and export back to a `.pptx` `Blob` via `exportBlob()`. Text inherited from a layout or master is rendered but read-only, since editing it would change every slide that shares it. Shape geometry, images and charts are not editable.
- DOCX: read-only. (Editing is **not** implemented — the pipeline below is the template for bringing it there.)
- XLSX: read-only.

### How PPTX editing works

`pptx/edit/` is the reference pattern for adding editing to another format:

1. **Address at parse time.** Model indices do not match XML indices (empty runs are dropped, `<a:br>`/`<a:fld>` become runs), so `ParseScope.recordSource` captures the exact `XmlNode` each text body/paragraph/run came from, into a `WeakMap` on the `Deck`. The model itself stays a pure value tree. Because each of a slide's master/layout/slide scopes carries its own part path, this also tells the editor what is inherited (read-only) versus slide-owned.
2. **Reconcile, don't intercept.** contentEditable is left alone to do whatever it likes; on blur the DOM subtree is read back into a `ParaEdit[]` (`reconcile.ts`). Typing, Enter, Backspace, paste and formatting therefore share one code path rather than one per interaction.
3. **Reuse over recreation.** `xmlWrite.ts` splices the *original* `<a:r>` node back in whenever a segment's text and formatting are unchanged, so an edit to one word leaves every other run — and every property the parser never read — byte-identical.
4. **Snapshot for undo.** `OpcPackage.serializePart` / `setPart` give exact undo without inverse ops.
5. **Export is non-destructive.** `toBytes()` re-serializes only dirty parts; every other part is emitted from its original bytes unchanged (there is a test asserting this).

Known limitations: `normAutofit` shrink factors are baked in at parse time and are not recomputed after an edit, so text can overflow its box (PowerPoint's "do not autofit" behaviour); `<a:fld>` runs render but are locked, since PowerPoint regenerates their text.

### Conventions when extending

- **Fidelity fixes must be generic** — correct for any conformant Office file, never special-cased to a particular deck/document.
- **Render-agnostic model first.** Parsing produces a plain model; rendering is a separate pass. New features should respect that boundary.
- **Adding a new format** (e.g. `.xlsx`): add a sibling `src/<format>/` mirroring the `pptx`/`docx` shape (`model.ts`, `relTypes.ts`, parse slices, `render/dom.ts`, `viewer/`), reuse `oxml/` for read/units/xml, then export `load<Format>` + `create<Format>Viewer` from `src/index.ts` and add a React wrapper/hook under `src/react/`.
