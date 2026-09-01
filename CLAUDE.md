# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

An npm-workspaces monorepo containing `@cobuildx.ai/office-viewer`, a published npm package that parses, renders and edits `.pptx`/`.docx`/`.xlsx` files directly in the browser (real HTML/CSS/SVG output — no PDF conversion, no LibreOffice, ever, including for reference/diff comparisons).

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

`.pptx`, `.docx` and `.xlsx` are all OPC (zip-of-XML) packages. Each format runs the same pipeline shape over one shared low-level layer:

```
PPTX:  read (OPC) -> parse (XML -> model) -> resolve (inheritance)  -> render (DOM)
DOCX:  read (OPC) -> parse (XML -> model) -> paginate (sections)    -> render (DOM)
XLSX:  read (OPC) -> parse (XML -> model, lazily per sheet)         -> render (DOM)
```

- **read** (`oxml/package.ts`) — unzip, resolve content types + relationships.
- **parse** — turn XML parts into a render-agnostic model (`pptx/model.ts` / `docx/model.ts`).
- **resolve/paginate** — the fidelity core. PPTX slides inherit geometry/color/font/text styles through layout → master → theme, with scheme colors mapped through the master's color map and transformed (lumMod/lumOff/tint/shade/alpha). DOCX styles cascade run → paragraph → linked → docDefaults, then blocks are measured off-screen and flowed into pages by section page-size/margins.
- **render** — model → HTML/CSS/SVG. PPTX positions shapes absolutely and scales the whole slide to fit, laid out PowerPoint-style with a thumbnail rail beside the stage; DOCX flows pages in a scrollable stack. All geometry converts EMU (914,400/inch) → CSS px at 96 DPI via `oxml/units.ts`.

### Package layout (`packages/core/src/`)

```
index.ts          public API: loadPptx / loadDocx, viewers, low-level exports
react/            React entry point (@cobuildx.ai/office-viewer/react)
                  PptxViewer.tsx, DocxViewer.tsx, DocxEditorToolbar.tsx, useDeck.ts, useDocument.ts

oxml/             SHARED, format-agnostic — the only layer pptx/ and docx/ may both depend on
  package.ts      OPC: unzip, content types, relationship resolution
  xml.ts          order-preserving, namespace-aware XML tree + helpers
  units.ts        EMU / point / twip -> CSS pixel conversions
  stylesheet.ts   ref-counted <style> injection, shared by all viewer chrome
  edit/           editing primitives every format shares: attrs.ts (DOM marker
                  names), format.ts (the RunFormat value), history.ts (snapshot
                  undo), selection.ts, styles.ts (editable-region outlines)

pptx/             PowerPoint feature slices
  deck/           top-level loader; presentation + slide/layout/master/theme
  edit/           text editing: source addressing, DOM reconciliation,
                  XML write-back, formatting, undo/redo; slideOps.ts holds the
                  structural (whole-slide) edits, shapeOps.ts the per-shape ones
                  (geometry write-back, deletion, z-order) and geometry.ts the
                  pure move/resize/rotate math
  slides/         slide composition (master -> layout -> slide)
  shapes/         fills, geometry, placeholders, shape props/render
  text/           text bodies, runs, text-style inheritance
  tables/ charts/ diagrams/ (SmartArt) pictures/ effects/
  color.ts        scheme-color map + modifier transforms
  model.ts        render-agnostic intermediate representation
  render/         model -> HTML/CSS/SVG (dom.ts), font install, primitives
  viewer/         viewer.ts (navigation, fit-to-viewport scaling, keyboard)
                  filmstrip.ts (the left-hand thumbnail rail)
                  selection.ts (shape handles, dragging, marquee)

docx/             Word feature slices (same shape as pptx)
  document/       top-level loader; body parsing
  edit/           text editing: source addressing, DOM reconciliation, XML
                  write-back, continuous-flow renderer
  paragraphs/ styles/ numbering/ tables/ images/
  model.ts        DOCX page/block/paragraph/table model
  render/         model -> HTML/CSS (dom.ts), pagination-aware
  viewer/         paginated scroll view + thumbnail strip
```

```
xlsx/             Excel feature slices (same shape as pptx/docx)
  workbook/       top-level loader; sheet list, lazy per-sheet parse, source map
  sheets/         worksheet parse (rows, cells, merges, cols) + ref utilities
  styles/         numFmts / fonts / fills / borders / cellXfs -> XlsxCellStyle
  edit/           cell editing: value interpretation (values.ts), worksheet
                  write-back (xmlWrite.ts), style interning (styleWrite.ts),
                  read-only rules (context.ts), undo/redo + commits (session.ts)
  model.ts        sheet/row/cell model
  render/         model -> HTML grid + formula bar (dom.ts), selection, keyboard
  viewer/         sheet tabs, commit cycle, export
```

**Hard rule: `pptx/`, `docx/` and `xlsx/` must never import from each other.** Anything both need belongs in `oxml/`, not duplicated or cross-imported.

### Status

- PPTX: renders ~106 preset shapes + custom geometry, charts (static SVG snapshot), tables, SmartArt (cached fast path + data-model layout fallback by family), gradients (incl. radial focus/path), autofit text. Effects (shadow/glow/reflection), transitions, and animations are not rendered. `.ppt` (legacy binary) is out of scope.
- PPTX **text editing** (opt in via `<PptxViewer editable>` / `createViewer(deck, el, { editable: true })`): inline WYSIWYG editing of the slide's own text bodies — including table cell text — with paragraph split/merge, a formatting toolbar (bold/italic/underline/strike, size, colour, typeface), undo/redo, and export back to a `.pptx` `Blob` via `exportBlob()`. Text inherited from a layout or master is rendered but read-only, since editing it would change every slide that shares it. Shape geometry, images and charts are not editable.
- PPTX **shape selection and direct manipulation** (same `editable` opt-in, `shapeEditing: false` to switch off): click a shape to select it, shift-click to add, drag on empty space to rubber-band. Drag to move, the eight grips to resize (shift keeps the ratio, alt resizes about the centre), the knob above to rotate (shift snaps to 15°). Arrows nudge 1px, shift-arrows 10px; Delete removes; Ctrl+`]`/`[` restack, adding shift to send all the way; Tab walks the slide's top-level shapes. Only the slide's own shapes are selectable — a shape composited in from the layout or master is drawn but belongs to every slide that shares it. Shape *content* (fill, geometry, images) is still not editable, and multiple shapes move together but do not resize as a block.
- PPTX **groups open on a second click**, as in PowerPoint: a click selects the group as a unit, clicking it again steps inside and selects the child that was under the pointer (repeat for a nested group), and Escape steps back out one level per press. A child inside a group is moved, resized, restacked, deleted and typed in like any other shape — its box is stated in the group's child space, so `render/primitives.ts` records the `ShapeFrame` each shape was drawn in and the overlay maps through it in both directions. A rotated or mirrored group stays the unit: handles for a child inside one could not line up with what is drawn.
- PPTX **text entry is selection-first**, as in PowerPoint: one click selects the shape, a second click on the already-selected shape (or Enter/F2) puts the caret in its text, and Escape leaves the text and re-selects the shape. A table works the same way — the second click opens the cell it landed in (Enter/F2 opens the first cell with text); moving between cells means leaving one and clicking into the next. Only the box being typed in is contentEditable, which is what lets a plain click mean "select".
- PPTX **slide deletion** (`viewer.deleteSlide(index?)`, same `editable` opt-in): removes the slide from the running order, its relationship, its `[Content_Types].xml` override, its own part and rels, and its notes slide, and purges custom-show / section references to it. A deck must keep at least one slide. Undoable like any other edit.
- DOCX **text editing** (opt in via `<DocxViewer editable>` / `createDocxViewer(doc, el, { editable: true })`): inline WYSIWYG editing of body text including table cell text, paragraph split/merge, the same formatting toolbar, undo/redo, and export to a `.docx` `Blob`. Header and footer text renders but is read-only, as are generated list markers and field results. **Edit mode renders the document as one continuous column instead of fixed pages** — see below.
- XLSX **cell editing** (opt in via `<XlsxViewer editable>` / `createXlsxViewer(wb, el, { editable: true })`): in-place editing of cell values and formulas with Excel-like keys (arrows, Enter, Tab, F2, Delete, shift-select, formula bar), cell formatting (font, fill, alignment, wrap, number format), undo/redo, and export to a `.xlsx` `Blob`. Formulas are stored but never evaluated — the workbook is flagged `fullCalcOnLoad` so Excel recalculates on open. Cells on a protected sheet, cells covered by a merge, and array/shared-formula hosts are read-only. Inserting/deleting rows and columns is out of scope (it would have to rewrite every reference and formula).

### How editing works

`pptx/edit/`, `docx/edit/` and `xlsx/edit/` share one architecture; the format-agnostic parts live in `oxml/edit/`:

1. **Address at parse time.** Model indices do not match XML indices (PPTX drops empty runs and turns `<a:br>`/`<a:fld>` into runs; DOCX's `logicalChildren` flattens `<w:sdt>`/`<w:smartTag>`/`<w:fldSimple>`), so `ParseScope.recordSource` / `ParseContext.recordSource` captures the exact `XmlNode` each object came from, into a `WeakMap` on the `Deck`/`DocxDocument`. The model itself stays a pure value tree. Because each parse scope carries its own part path, this also identifies what is read-only for free — layout/master text in PPTX, header/footer text in DOCX.
2. **Reconcile, don't intercept.** contentEditable is left alone to do whatever it likes; on blur the DOM subtree is read back into a `ParaEdit[]` (`reconcile.ts`). Typing, Enter, Backspace, paste and formatting therefore share one code path rather than one per interaction. The editable unit is a text body in PPTX and a single paragraph in DOCX, because a DOCX body is one long flow with no txBody-sized container.
3. **Reuse over recreation.** `xmlWrite.ts` splices the *original* run node back in whenever a segment's text and formatting are unchanged, so an edit to one word leaves every other run — and every property the parser never read — byte-identical.
4. **Snapshot for undo.** `OpcPackage.serializePart` / `setPart` give exact undo without inverse ops (`oxml/edit/history.ts`).
5. **Export is non-destructive.** `toBytes()` re-serializes only dirty parts; every other part is emitted from its original bytes unchanged (there is a test asserting this).

**Format-specific wrinkles worth knowing before you touch a slice:**

- **DOCX runs are many-to-one.** One `<w:r>` emits one `DocxRun` per `<w:t>` child, so a run is addressed by its `<w:t>`, with the owning `<w:r>` carried alongside for `<w:rPr>`. In PPTX an `<a:r>` is one run.
- **XLSX has no inline formatting.** A cell points at a `cellXfs` entry, which points at a font/fill/border/numFmt. So formatting means *interning* (`edit/styleWrite.ts`): derive the new entry from the one the cell already uses, reuse an identical existing entry, append only when there is none. It also means one formatting change spans two parts (worksheet + `xl/styles.xml`), which is why `oxml/edit/history.ts` stores a *change set* rather than a single snapshot.
- **XLSX text is written inline.** An edited string becomes `t="inlineStr"` rather than a new entry in `xl/sharedStrings.xml`, so editing one cell never rewrites a part every sheet shares.
- **DOCX edits in continuous flow.** The paginator (`docx/viewer/paginate.ts`) splits a paragraph mid-text across a page boundary into *cloned* objects on two page sheets — a contentEditable region cannot span that, and the clones have no source identity. So edit mode renders one continuous column (`docx/edit/flow.ts`) and re-paginates on exit. This costs no fidelity: **pagination is not stored in a `.docx`** (Word reflows on open), so our pages are purely a display concern.

**Traps that have already bitten:**

- **Every shape kind keeps its transform somewhere else.** `<p:sp>`/`<p:pic>`/`<p:cxnSp>` use `<a:xfrm>` inside `<p:spPr>`, a group uses one inside `<p:grpSpPr>`, and a `<p:graphicFrame>` uses a `<p:xfrm>` as its own direct child in the *presentation* namespace. The element also frequently does not exist (a placeholder inherits its position), and the schema is sequence-ordered, so it has to be created in the right slot — `<a:xfrm>` first in spPr — not appended.
- **Resizing a group must not touch `chOff`/`chExt`.** Children map through the ext/chExt ratio, so changing `ext` alone is exactly what scales the group's contents; rewriting chExt would freeze them.
- **`preventDefault()` on `pointerdown` suppresses the compatibility mouse events**, `click` and `dblclick` included. The drag needs that call (otherwise the browser starts its own selection drag), so entering a text box is driven from `pointerup` on a shape that was already selected — which is PowerPoint's gesture anyway — and never from a `dblclick` listener. It also means focus does not move on its own, so selecting a shape focuses the container explicitly or the keys bound there never fire.
- **A group's children are marked selectable too, but a press picks the shallowest shape the user is not already inside.** That one rule is what makes a click select the group and the next click reach into it, and it is why the selection layer walks the *chain* of `data-cbx-shape` ancestors rather than calling `closest()` once.
- **`el.contentEditable = 'true'` does not reflect to the attribute in jsdom**, so the box is opened with `setAttribute` — the viewer finds the box it has just opened with an attribute selector, and a property-only write made that silently fail under test while working in a browser.
- **A resize must pin the corner the user is *not* dragging.** Rotation and mirroring both change which stored corner that is, so `edit/geometry.ts` works in the shape's local space and recomputes the offset from the anchor rather than adding the delta to a corner. There are tests asserting the anchor is fixed at every rotation and flip.

- `Numbering` (`docx/numbering/numbering.ts`) holds live counters that `marker()` advances while walking the body. `DocxDocument.rebuild()` must construct a **fresh** `Numbering`, or a re-parse renders an ordered list as "4. 5. 6.". There is a test for this.
- **A commit invalidates every model reference held across it.** Both `Deck.rebuildSlide` and `DocxDocument.rebuild` build fresh paragraphs/runs and drop the source map, so re-read from `deck.slides` / `doc.sections` after each commit. A *structural* PPTX edit goes further: `Deck.rebuild()` re-reads the running order and rebuilds every slide, so slide indices shift too.
- **Thumbnails render through `renderSlide`, never a second renderer** — a thumbnail that disagrees with its slide is worse than no thumbnail. They are drawn lazily (IntersectionObserver, with an eager fallback where it is missing) and *without* the edit context, so the rail can never become a second place to type.
- **A deleted slide is referenced from four places at once** — `<p:sldIdLst>`, presentation.xml.rels, a `[Content_Types].xml` override, and (optionally) custom shows / the `<p14:section>` list. Leaving any one behind makes PowerPoint call the file corrupt, so `pptx/edit/slideOps.ts` cleans all of them in one pass. Media the slide used is deliberately *left*: other slides usually share it, and an unreferenced media part is harmless.
- **`OpcPackage.deletePart` also drops the part's `.rels`,** so anything snapshotting for undo must list those rels parts explicitly — a deletion it cannot put back in full is not undoable. `Snapshot.absent` is how "this part did not exist" round-trips through `History`.
- `OpcPackage.toBlob` defaults to the **docx** content type — PPTX export must pass its own.

- **Rows and cells must stay in ascending order.** Excel rejects a worksheet where they are not, so `xlsx/edit/xmlWrite.ts` splices new `<row>`/`<c>` elements into position rather than appending.

Known limitations: PPTX `normAutofit` shrink factors are baked in at parse time and are not recomputed after an edit, so text can overflow its box (PowerPoint's "do not autofit" behaviour). Generated text — `<a:fld>` in PPTX, field results and list markers in DOCX — renders but is locked, since Office regenerates it.

### Conventions when extending

- **Fidelity fixes must be generic** — correct for any conformant Office file, never special-cased to a particular deck/document.
- **Render-agnostic model first.** Parsing produces a plain model; rendering is a separate pass. New features should respect that boundary.
- **Adding a new format**: add a sibling `src/<format>/` mirroring the `pptx`/`docx`/`xlsx` shape (`model.ts`, `relTypes.ts`, parse slices, `render/dom.ts`, `viewer/`), reuse `oxml/` for read/units/xml, then export `load<Format>` + `create<Format>Viewer` from `src/index.ts` and add a React wrapper/hook under `src/react/`.
