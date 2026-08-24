import { useEffect, useImperativeHandle, useRef, useState, forwardRef } from 'react';
import type { CSSProperties } from 'react';
import {
  DocxViewer as DocxViewerController,
  type DocxDocument,
  type RunFormat,
} from '../index.js';
import { useDocument, type DocxSource } from './useDocument.js';
import { EditorToolbar } from './EditorToolbar.js';

export interface DocxViewerHandle {
  next(): void;
  prev(): void;
  goTo(index: number): void;
  zoomIn(): void;
  zoomOut(): void;
  fitWidth(): void;
  /** Format the current selection (editable mode only). */
  applyFormat(format: RunFormat): void;
  /** Commit the focused paragraph without waiting for blur. */
  commit(): void;
  undo(): void;
  redo(): void;
  /** Re-zip the document with all edits applied. */
  exportBlob(): Blob | undefined;
  doc: DocxDocument | null;
}

export interface DocxViewerProps {
  /** A File (e.g. from an <input>), ArrayBuffer, or Uint8Array. */
  src: DocxSource;
  /** Show the built-in page-navigation toolbar. Default true. */
  toolbar?: boolean;
  className?: string;
  style?: CSSProperties;
  onLoad?: (doc: DocxDocument) => void;
  onError?: (error: Error) => void;
  onPageChange?: (index: number, count: number) => void;
  /**
   * Edit the document's body text in place. Editing renders the document as a
   * single continuous column rather than fixed pages; header and footer text
   * stays read-only. Off by default.
   */
  editable?: boolean;
  /** Show the formatting toolbar when `editable`. Default true. */
  editorToolbar?: boolean;
  /** Called after each committed edit, undo or redo. */
  onEdit?: () => void;
}

export const DocxViewer = forwardRef<DocxViewerHandle, DocxViewerProps>(function DocxViewer(
  {
    src,
    toolbar = true,
    className,
    style,
    onLoad,
    onError,
    onPageChange,
    editable = false,
    editorToolbar = true,
    onEdit,
  },
  ref,
) {
  const { doc, loading, error } = useDocument(src);
  const stageRef = useRef<HTMLDivElement>(null);
  const viewerRef = useRef<DocxViewerController | null>(null);
  const [index, setIndex] = useState(0);
  const [count, setCount] = useState(0);
  const [scale, setScale] = useState(1);
  const [format, setFormat] = useState<RunFormat>({});
  // The viewer owns undo/redo state; mirror it so the toolbar re-renders.
  const [editState, setEditState] = useState({ canUndo: false, canRedo: false, hasEdits: false });

  useEffect(() => {
    if (error) onError?.(error);
  }, [error, onError]);

  useEffect(() => {
    if (!doc || !stageRef.current) return;
    onLoad?.(doc);
    const viewer = new DocxViewerController(doc, stageRef.current, {
      editable,
      onChange: (i, c) => {
        setIndex(i);
        setCount(c);
        onPageChange?.(i, c);
      },
      onScaleChange: setScale,
      onEdit: () => {
        setEditState({
          canUndo: viewer.canUndo,
          canRedo: viewer.canRedo,
          hasEdits: viewer.hasEdits,
        });
        onEdit?.();
      },
      onSelectionChange: setFormat,
    });
    viewerRef.current = viewer;
    return () => {
      viewer.destroy();
      viewerRef.current = null;
    };
  }, [doc, editable]); // eslint-disable-line react-hooks/exhaustive-deps

  useImperativeHandle(
    ref,
    () => ({
      next: () => viewerRef.current?.next(),
      prev: () => viewerRef.current?.prev(),
      goTo: (i: number) => viewerRef.current?.goTo(i),
      zoomIn: () => viewerRef.current?.zoomIn(),
      zoomOut: () => viewerRef.current?.zoomOut(),
      fitWidth: () => viewerRef.current?.fitWidth(),
      applyFormat: (f: RunFormat) => viewerRef.current?.applyFormat(f),
      commit: () => viewerRef.current?.commitActive(),
      undo: () => viewerRef.current?.undo(),
      redo: () => viewerRef.current?.redo(),
      exportBlob: () => viewerRef.current?.exportBlob(),
      doc,
    }),
    [doc],
  );

  const downloadEdits = () => {
    const blob = viewerRef.current?.exportBlob();
    if (!blob) return;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'edited.docx';
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className={className} style={{ display: 'flex', flexDirection: 'column', ...style }}>
      {editable && editorToolbar && doc && (
        <EditorToolbar
          format={format}
          onFormat={(f) => viewerRef.current?.applyFormat(f)}
          onUndo={() => viewerRef.current?.undo()}
          onRedo={() => viewerRef.current?.redo()}
          canUndo={editState.canUndo}
          canRedo={editState.canRedo}
          hasEdits={editState.hasEdits}
          onExport={downloadEdits}
          exportLabel="Download .docx"
        />
      )}
      <div ref={stageRef} style={{ flex: 1, minHeight: 0 }} />
      {loading && <div style={statusStyle}>Loading…</div>}
      {error && <div style={{ ...statusStyle, color: '#e57373' }}>Error: {error.message}</div>}
      {toolbar && doc && (
        <div style={toolbarStyle}>
          <button style={btn} onClick={() => viewerRef.current?.prev()} disabled={index <= 0}>
            ‹ Prev
          </button>
          <span style={{ minWidth: 90, textAlign: 'center' }}>
            Page {count === 0 ? 0 : index + 1} / {count}
          </span>
          <button
            style={btn}
            onClick={() => viewerRef.current?.next()}
            disabled={index >= count - 1}
          >
            Next ›
          </button>

          <span style={divider} />

          <button style={iconBtn} title="Zoom out (Ctrl -)" onClick={() => viewerRef.current?.zoomOut()}>
            −
          </button>
          <span style={{ minWidth: 46, textAlign: 'center' }}>{Math.round(scale * 100)}%</span>
          <button style={iconBtn} title="Zoom in (Ctrl +)" onClick={() => viewerRef.current?.zoomIn()}>
            +
          </button>
          <button style={btn} title="Fit width (Ctrl 0)" onClick={() => viewerRef.current?.fitWidth()}>
            Fit
          </button>
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

const iconBtn: CSSProperties = {
  ...btn,
  padding: '4px 10px',
  fontSize: 15,
  lineHeight: 1,
  minWidth: 30,
};

const divider: CSSProperties = {
  width: 1,
  height: 20,
  background: '#444',
};
