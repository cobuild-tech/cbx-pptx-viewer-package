export { PptxViewer, type PptxViewerProps, type PptxViewerHandle } from './PptxViewer.js';
export { EditorToolbar, type EditorToolbarProps } from './EditorToolbar.js';
export { ParaControls, type ParaControlsProps } from './ParaControls.js';
// Kept for compatibility with 0.1.6, which shipped the PPTX-specific name.
export {
  EditorToolbar as PptxEditorToolbar,
  type EditorToolbarProps as PptxEditorToolbarProps,
} from './EditorToolbar.js';
export { useDeck, type DeckSource, type DeckState } from './useDeck.js';
export { DocxViewer, type DocxViewerProps, type DocxViewerHandle } from './DocxViewer.js';
export { useDocument, type DocxSource, type DocumentState } from './useDocument.js';
export { XlsxViewer, type XlsxViewerProps, type XlsxViewerHandle } from './XlsxViewer.js';
export { useWorkbook, type WorkbookSource, type WorkbookState } from './useWorkbook.js';

