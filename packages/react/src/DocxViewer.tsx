import { useEffect, useImperativeHandle, useRef, forwardRef } from 'react';
import type { CSSProperties } from 'react';
import { renderDocument, type DocxDocument } from '@pptx-viewer/core';
import { useDocx, type DocxSource } from './useDocx.js';

export interface DocxViewerHandle {
  document: DocxDocument | null;
}

export interface DocxViewerProps {
  /** A File (e.g. from an <input>), ArrayBuffer, or Uint8Array. */
  src: DocxSource;
  className?: string;
  style?: CSSProperties;
  onLoad?: (doc: DocxDocument) => void;
  onError?: (error: Error) => void;
}

export const DocxViewer = forwardRef<DocxViewerHandle, DocxViewerProps>(function DocxViewer(
  { src, className, style, onLoad, onError },
  ref,
) {
  const { document: doc, loading, error } = useDocx(src);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (error) onError?.(error);
  }, [error, onError]);

  useEffect(() => {
    if (!doc || !containerRef.current) return;
    onLoad?.(doc);
    renderDocument(doc, containerRef.current);
  }, [doc]);

  useImperativeHandle(
    ref,
    () => ({
      document: doc,
    }),
    [doc],
  );

  return (
    <div className={className} style={{ display: 'flex', flexDirection: 'column', overflow: 'auto', ...style }}>
      <div style={{ flex: 1, minHeight: 0, padding: '24px 16px', background: '#323639', display: 'flex', justifyContent: 'center' }}>
        <div ref={containerRef} style={{ width: '100%', maxWidth: '816px' }} />
      </div>
      {loading && <div style={statusStyle}>Loading docx…</div>}
      {error && <div style={{ ...statusStyle, color: '#e57373' }}>Error: {error.message}</div>}
    </div>
  );
});

const statusStyle: CSSProperties = {
  padding: '8px 12px',
  font: '13px system-ui, sans-serif',
  color: '#ddd',
};
