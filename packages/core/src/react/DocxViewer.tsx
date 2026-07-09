import { useEffect, useImperativeHandle, useRef, useState, forwardRef } from 'react';
import type { CSSProperties } from 'react';
import { DocxViewer as DocxViewerController, type DocxDocument } from '../index.js';
import { useDocument, type DocxSource } from './useDocument.js';

export interface DocxViewerHandle {
  next(): void;
  prev(): void;
  goTo(index: number): void;
  zoomIn(): void;
  zoomOut(): void;
  fitWidth(): void;
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
}

export const DocxViewer = forwardRef<DocxViewerHandle, DocxViewerProps>(function DocxViewer(
  { src, toolbar = true, className, style, onLoad, onError, onPageChange },
  ref,
) {
  const { doc, loading, error } = useDocument(src);
  const stageRef = useRef<HTMLDivElement>(null);
  const viewerRef = useRef<DocxViewerController | null>(null);
  const [index, setIndex] = useState(0);
  const [count, setCount] = useState(0);
  const [scale, setScale] = useState(1);

  useEffect(() => {
    if (error) onError?.(error);
  }, [error, onError]);

  useEffect(() => {
    if (!doc || !stageRef.current) return;
    onLoad?.(doc);
    const viewer = new DocxViewerController(doc, stageRef.current, {
      onChange: (i, c) => {
        setIndex(i);
        setCount(c);
        onPageChange?.(i, c);
      },
      onScaleChange: setScale,
    });
    viewerRef.current = viewer;
    return () => {
      viewer.destroy();
      viewerRef.current = null;
    };
  }, [doc]); // eslint-disable-line react-hooks/exhaustive-deps

  useImperativeHandle(
    ref,
    () => ({
      next: () => viewerRef.current?.next(),
      prev: () => viewerRef.current?.prev(),
      goTo: (i: number) => viewerRef.current?.goTo(i),
      zoomIn: () => viewerRef.current?.zoomIn(),
      zoomOut: () => viewerRef.current?.zoomOut(),
      fitWidth: () => viewerRef.current?.fitWidth(),
      doc,
    }),
    [doc],
  );

  return (
    <div className={className} style={{ display: 'flex', flexDirection: 'column', ...style }}>
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
