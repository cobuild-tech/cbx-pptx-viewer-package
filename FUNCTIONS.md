# FUNCTIONS.md — cbx-pptx-viewer-package

Function/class reference for the main source files. Update this file after any edits.

---

## `packages/core/src/pdf/viewer/viewer.ts`

**`class PdfViewer`**  
Core PDF viewer. Mounts a `PdfDocument` into a container, renders all pages to canvas, and presents a vertically scrollable stack with a right-side thumbnail strip and zoom controls.

| Method / Property | Purpose |
|---|---|
| `constructor(doc, container, options)` | Builds the viewer DOM (top toolbar with zoom + undo/redo + formatting slot, scroll area, thumbnail sidebar). Wires resize/intersection observers and optional keyboard shortcuts. |
| `get editMode` | Returns whether edit mode is currently active. |
| `setEditable(enabled)` | Toggles in-place text editing. On first enable, loads the text layer and mounts the edit overlay. On disable, clears the formatting slot and resets edit overlays. |
| `goTo(index)` | Smooth-scrolls to the page at `index`. |
| `next()` | Advances to the next page. |
| `prev()` | Goes back to the previous page. |
| `zoomIn()` | Increases zoom by one step. |
| `zoomOut()` | Decreases zoom by one step. |
| `setZoom(level)` | Sets an explicit zoom level (clamped to `ZOOM_MIN`–`ZOOM_MAX`). |
| `zoomFit()` | Resets zoom to fit-width. |
| `get count` | Total number of pages. |
| `get currentIndex` | Index of the currently most-visible page. |
| `destroy()` | Removes the viewer DOM, disconnects all observers and keyboard listeners. |
| `createPdfViewer(doc, container, options)` | Convenience factory. |

**Top toolbar layout (DOM, left→right):**  
`− zoom% +` · `|` · *formatting slot* (font / colour controls appear here when a text block is focused in edit mode; horizontally scrollable via mouse wheel or trackpad)

---

## `packages/core/src/pdf/edit/editing.ts`

**`class PdfEditController`**  
Manages interactive contenteditable text overlays for in-place PDF editing.

| Method / Property | Purpose |
|---|---|
| `constructor(pageContainers, blocks, liveEdits, annotations, blockStylesMap, onCommit, formattingSlot?)` | Mounts overlay elements on every page container. If `formattingSlot` is supplied, formatting controls render inline inside that element (scrollable top bar) instead of as a floating panel. |
| `setEditable(enabled)` | Shows/hides overlay interactivity. Clears `formattingSlot` content when disabling. |
| `refreshAll()` | Re-syncs all block and annotation overlays from the live data maps (called after undo/redo). |
| `unmount()` | Removes all overlays and clears the formatting slot. |
| `showToolbar(wrapperEl, state, contentDiv, onPickerColor)` *(private)* | Builds the formatting toolbar (font picker, Bold, Italic, alignment, size ±, colour swatches, custom colour input). Renders into `formattingSlot` when provided (no fixed positioning); falls back to a floating `position:fixed` panel otherwise. Returns a cleanup function. |
| `mkBtn(label, cb)` *(private)* | Creates a styled toolbar button. |
| `startNewAnnotation(...)` *(private)* | Places a new free-form text annotation on a page click. Box uses `fit-content` with right/bottom resize handles; `ann.width` and `ann.height` captured on blur. |
| `buildCommittedEl(ann, pageIndex, annotationEls)` *(private)* | Builds a committed (persisted) annotation element with drag, right/bottom resize handles, focus, and blur handling. Width and height updates committed as `updateAnnotation` ops (undo-able). |
| `mountOverlays(containers)` *(private)* | Creates and attaches an overlay div per page. |
| `wireBlockDrag` *(private)* | Mouse-drag handler for repositioning existing text blocks via grab bar or wrapper body; resize handles are excluded. |
| `wireDrag` *(private)* | Mouse-drag handler for annotation elements via grab bar or wrapper body; resize handles are excluded. |
| `wireResize(handle, wrapper, direction, container, onCommit)` *(private)* | Unified resize handler. `direction:'h'` adjusts width (right-edge handle), `direction:'v'` adjusts min-height (bottom-edge handle). Calls `onCommit` once on mouse-up with the new size. |
| `applyBlockWrapperStyle` *(private)* | Applies visual state (edit / overlay / hidden) to a block wrapper. Uses an **opaque sampled background** on the overlay `contentDiv` to cover canvas text — no canvas modification for plain edits or style changes. Canvas is only erased (`eraseBlockOnCanvas`) when the block has been **physically repositioned** (`hasMoved`), preventing drag from corrupting other blocks. In edit mode uses `width:fit-content; min-width:cssWidth` and shows right/bottom resize handles. |
| `applyWrapperStyle` *(private)* | Applies visual state to an annotation wrapper. In edit mode uses `width:fit-content; min-width:ann.width` and shows right/bottom resize handles. Width is committed to `ann.width` on blur for PDF export. |
| `eraseBlockOnCanvas / restoreBlockOnCanvas` *(private)* | Erase canvas pixels at a block's **original** coordinates (save/restore via `ImageData`). Called only when a block has been repositioned (`hasMoved`) so the original canvas text disappears from the old position. Restored on undo when the block returns to hidden state. |
| `sampleBlockBg` *(private)* | Samples the canvas pixel colour above and below a text block to infer background and contrast text colour. Result is cached in `bgCache` and used as the opaque overlay background. |

---

## `packages/core/src/pdf/edit/fontPicker.ts`

**`showFontPicker(anchorEl, currentFont, onSelect)`**  
Shows a floating dark-themed font picker panel anchored near `anchorEl`. Includes a search input, category tabs (All / Sans-serif / Serif / Monospace / Display / Handwriting), and a scrollable font list with live Google Fonts previews. Scroll events that originate inside the panel itself do **not** close it — only external page scrolls do. All interactions use `mousedown + preventDefault` to avoid stealing focus from the calling contenteditable. Returns a cleanup function that removes the picker.

---

## `packages/core/src/pdf/edit/fonts.ts`

| Export | Purpose |
|---|---|
| `FONTS` | Static array of `FontDefinition` objects (name, category, cssStack, googleFamily). |
| `ensureGoogleFontsLoaded(fonts)` | Injects a `<link>` tag for any Google Font not yet loaded. |
| `resolveCssFontStack(name)` | Returns the full CSS font-family stack for a given font display name. |

---

## `packages/core/src/pdf/edit/ops.ts`

Defines the `PdfEditOp` discriminated union and the logic for applying and undoing each operation type (`replaceText`, `styleBlock`, `addAnnotation`, `updateAnnotation`, `removeAnnotation`).

---

## `packages/core/src/pdf/edit/export.ts`

**`exportPdf(doc, ops)`**  
Applies all committed edit operations to a `PDFDocument` (via `pdf-lib`) and returns the modified PDF as a `Uint8Array`.

---

## `packages/core/src/pdf/edit/textLayer.ts`

| Export | Purpose |
|---|---|
| `extractPageTextItems(proxy, pageIndex, pageHeightPx)` | Uses pdf.js `getTextContent()` to extract positioned `PdfTextItem` objects from one page, converting PDF coordinates (origin bottom-left, y-up) to CSS coordinates (origin top-left, y-down). |
| `groupItemsIntoBlocks(items)` | Groups text items into `PdfTextBlock` units: items on the same baseline (within `avgFontSize × 0.5`) are merged, split by horizontal column gaps (`> avgFontSize × 2.5`). |

---

## `packages/core/src/pdf/edit/versions.ts`

Undo/redo stack manager for `PdfEditOp` groups.

| Export | Purpose |
|---|---|
| `class InMemoryPdfVersionStore` | Stores committed op-groups in memory. Exposes `push`, `undo`, `redo`, `canUndo`, `canRedo`. |

---

## `packages/core/src/pdf/document/document.ts`

**`class PdfDocument`**  
Wraps a `PDFDocumentProxy` (pdf.js). Owns the edit state maps (`editsMap`, `annotationsMap`, `blockStylesMap`) and the version store.

| Method | Purpose |
|---|---|
| `renderPage(index, scale)` | Renders a page to a `<canvas>` at the given scale (returns a Promise). |
| `loadAllTextBlocks()` | Lazily extracts and groups text blocks for all pages. |
| `applyEdits(ops)` | Applies a group of `PdfEditOp`s and pushes them onto the version store. |
| `undo() / redo()` | Walks the version store and re-applies/reverts the op group. |
| `exportBlob()` | Exports the document with all edits applied as a `Blob` (calls `exportPdf`). |
| `saveVersion() / configureVersioning(store)` | External version-store integration. |
| `get hasUnsavedChanges / canUndo / canRedo` | State accessors. |

---

## `packages/react/src/PdfViewer.tsx`

**`PdfViewer` (React component, forwardRef)**  
React wrapper around the core `PdfViewer` class.

| Prop | Purpose |
|---|---|
| `src` | `File`, `ArrayBuffer`, or `Uint8Array`. |
| `editable` | Enables text editing mode. Default `false`. |
| `versionStore` | External `PdfVersionStore` for save/load of edit history. |
| `toolbar` | Show the built-in bottom toolbar. Default `true`. |
| `onLoad / onError / onPageChange / onVersionSaved` | Lifecycle callbacks. |

**Top action bar** (shown when `editable` is `true`): `Edit / Done` · `Save Version` · `Download`  
**Bottom toolbar:** `↑ Prev` · `Page X / Y` · `Next ↓`

**Ref handle (`PdfViewerHandle`):** `next()`, `prev()`, `goTo(index)`, `setEditable(enabled)`, `doc`.

---

## `packages/react/src/PptxViewer.tsx`

**`PptxViewer` (React component, forwardRef)**  
React wrapper around the core `Viewer` class for PPTX files.

| Prop | Purpose |
|---|---|
| `src` | `File`, `ArrayBuffer`, or `Uint8Array`. |
| `toolbar` | Show the built-in slide navigation toolbar. Default `true`. |
| `onLoad / onError / onSlideChange` | Lifecycle callbacks. |

**Bottom toolbar:** `‹ Prev` · `Slide X / Y` · `Next ›`

**Ref handle (`PptxViewerHandle`):** `next()`, `prev()`, `goTo(index)`, `deck`.

---

## `packages/react/src/DocxViewer.tsx`

React wrapper around the core `DocxViewer` class for DOCX files.

---

## `packages/react/src/usePdf.ts`

**`usePdf(src)`**  
React hook. Loads a PDF source (`File` | `ArrayBuffer` | `Uint8Array`) and returns `{ doc, loading, error }`.

---

## `packages/react/src/useDeck.ts`

**`useDeck(src)`**  
React hook. Loads a PPTX source and returns `{ deck, loading, error }`.

---

## `packages/react/src/useDocument.ts`

**`useDocument(src)`**  
React hook. Loads a DOCX source and returns `{ document, loading, error }`.
