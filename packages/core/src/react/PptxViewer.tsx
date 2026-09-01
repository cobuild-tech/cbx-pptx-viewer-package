import { useEffect, useImperativeHandle, useRef, useState, forwardRef } from 'react';
import type { CSSProperties } from 'react';
import { Viewer, type Deck, type RunFormat, type TextBoxOutline } from '../index.js';
import { useDeck, type DeckSource } from './useDeck.js';
import { EditorToolbar } from './EditorToolbar.js';

export interface PptxViewerHandle {
  next(): void;
  prev(): void;
  goTo(index: number): void;
  /** Format the current selection (editable mode only). */
  applyFormat(format: RunFormat): void;
  /** Commit whatever text body is focused, without waiting for blur. */
  commit(): void;
  /** Delete a slide (the current one by default). Editable mode only. */
  deleteSlide(index?: number): void;
  undo(): void;
  redo(): void;
  /** Re-zip the deck with all edits applied. */
  exportBlob(): Blob | undefined;
  deck: Deck | null;
}

export interface PptxViewerProps {
  /** A File (e.g. from an <input>), ArrayBuffer, or Uint8Array. */
  src: DeckSource;
  /** Show the built-in navigation toolbar. Default true. */
  toolbar?: boolean;
  /**
   * Show the thumbnail rail down the left, PowerPoint-style. Default true.
   */
  filmstrip?: boolean;
  /** Width of the thumbnail rail in CSS px. Default 200. */
  filmstripWidth?: number;
  className?: string;
  style?: CSSProperties;
  onLoad?: (deck: Deck) => void;
  onError?: (error: Error) => void;
  onSlideChange?: (index: number, count: number) => void;
  /**
   * Edit the slide's own text in place. Text inherited from the layout or
   * master stays read-only. Off by default.
   */
  editable?: boolean;
  /** Show the formatting toolbar when `editable`. Default true. */
  editorToolbar?: boolean;
  /** Called after each committed edit, undo or redo. */
  onEdit?: (slideIndex: number) => void;
  /**
   * How editable text boxes are outlined: `'hover'` (default), `'always'` to
   * reveal every box at once, or `'none'`. Ignored unless `editable`.
   */
  textBoxOutline?: TextBoxOutline;
}

export const PptxViewer = forwardRef<PptxViewerHandle, PptxViewerProps>(function PptxViewer(
  {
    src,
    toolbar = true,
    filmstrip = true,
    filmstripWidth,
    className,
    style,
    onLoad,
    onError,
    onSlideChange,
    editable = false,
    editorToolbar = true,
    onEdit,
    textBoxOutline = 'hover',
  },
  ref,
) {
  const { deck, loading, error } = useDeck(src);
  const stageRef = useRef<HTMLDivElement>(null);
  const viewerRef = useRef<Viewer | null>(null);
  const [index, setIndex] = useState(0);
  const [count, setCount] = useState(0);
  const [format, setFormat] = useState<RunFormat>({});
  // The viewer owns undo/redo state; mirror it so the toolbar re-renders.
  const [editState, setEditState] = useState({
    canUndo: false,
    canRedo: false,
    hasEdits: false,
    canDelete: false,
  });

  useEffect(() => {
    if (error) onError?.(error);
  }, [error, onError]);

  useEffect(() => {
    if (!deck || !stageRef.current) return;
    onLoad?.(deck);
    const viewer = new Viewer(deck, stageRef.current, {
      editable,
      textBoxOutline,
      filmstrip,
      ...(filmstripWidth !== undefined ? { filmstripWidth } : {}),
      onChange: (i, c) => {
        setIndex(i);
        setCount(c);
        // Deleting the second-to-last slide makes the last one undeletable, so
        // this has to track navigation as well as edits.
        setEditState((prev) => ({ ...prev, canDelete: viewer.canDeleteSlide() }));
        onSlideChange?.(i, c);
      },
      onEdit: (i) => {
        setEditState({
          canUndo: viewer.canUndo,
          canRedo: viewer.canRedo,
          hasEdits: viewer.hasEdits,
          canDelete: viewer.canDeleteSlide(),
        });
        onEdit?.(i);
      },
      onSelectionChange: setFormat,
    });
    viewerRef.current = viewer;
    return () => {
      viewer.destroy();
      viewerRef.current = null;
    };
    // The rail is part of the viewer's DOM layout, so toggling it rebuilds.
  }, [deck, editable, filmstrip, filmstripWidth]); // eslint-disable-line react-hooks/exhaustive-deps

  // Changing the outline mode restyles in place — no need to rebuild the viewer.
  useEffect(() => {
    viewerRef.current?.setTextBoxOutline(textBoxOutline);
  }, [textBoxOutline]);

  useImperativeHandle(
    ref,
    () => ({
      next: () => viewerRef.current?.next(),
      prev: () => viewerRef.current?.prev(),
      goTo: (i: number) => viewerRef.current?.goTo(i),
      applyFormat: (f: RunFormat) => viewerRef.current?.applyFormat(f),
      commit: () => viewerRef.current?.commitActive(),
      deleteSlide: (i?: number) => viewerRef.current?.deleteSlide(i),
      undo: () => viewerRef.current?.undo(),
      redo: () => viewerRef.current?.redo(),
      exportBlob: () => viewerRef.current?.exportBlob(),
      deck,
    }),
    [deck],
  );

  const downloadEdits = () => {
    const blob = viewerRef.current?.exportBlob();
    if (!blob) return;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'edited.pptx';
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className={className} style={{ display: 'flex', flexDirection: 'column', ...style }}>
      {editable && editorToolbar && deck && (
        <EditorToolbar
          format={format}
          onFormat={(f) => viewerRef.current?.applyFormat(f)}
          onUndo={() => viewerRef.current?.undo()}
          onRedo={() => viewerRef.current?.redo()}
          canUndo={editState.canUndo}
          canRedo={editState.canRedo}
          hasEdits={editState.hasEdits}
          onExport={downloadEdits}
          exportLabel="Download .pptx"
        />
      )}
      {/* The viewer fits the slide to this area (contain) and centres it. */}
      <div style={{ flex: 1, minHeight: 0, overflow: 'hidden', background: '#525659' }}>
        <div ref={stageRef} style={{ width: '100%', height: '100%' }} />
      </div>
      {loading && <div style={statusStyle}>Loading…</div>}
      {error && <div style={{ ...statusStyle, color: '#e57373' }}>Error: {error.message}</div>}
      {toolbar && deck && (
        <div style={toolbarStyle}>
          <button style={btn} onClick={() => viewerRef.current?.prev()} disabled={index <= 0}>
            ‹ Prev
          </button>
          <span style={{ minWidth: 90, textAlign: 'center' }}>
            Slide {count === 0 ? 0 : index + 1} / {count}
          </span>
          <button
            style={btn}
            onClick={() => viewerRef.current?.next()}
            disabled={index >= count - 1}
          >
            Next ›
          </button>
          {editable && (
            <button
              style={{ ...btn, marginLeft: 12 }}
              onClick={() => viewerRef.current?.deleteSlide()}
              disabled={!editState.canDelete}
              title={
                editState.canDelete
                  ? 'Delete this slide'
                  : 'A presentation must keep at least one slide'
              }
            >
              Delete slide
            </button>
          )}
        </div>
      )}
    </div>
  );
});

const statusStyle: CSSProperties = {
  padding: '8px 12px',
  font: '13px system-ui, sans-serif',
  color: '#ddd',
};

const toolbarStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  gap: 12,
  padding: '8px 12px',
  background: '#2a2a2a',
  color: '#eee',
  font: '13px system-ui, sans-serif',
};

const btn: CSSProperties = {
  padding: '4px 12px',
  borderRadius: 6,
  border: '1px solid #555',
  background: '#3a3a3a',
  color: '#eee',
  cursor: 'pointer',
};
