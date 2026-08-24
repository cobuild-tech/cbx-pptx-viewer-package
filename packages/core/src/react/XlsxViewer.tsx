import { useEffect, useImperativeHandle, useRef, forwardRef } from 'react';
import type { CSSProperties } from 'react';
import { XlsxViewer as XlsxViewerController, type Workbook } from '../index.js';
import { useWorkbook, type WorkbookSource } from './useWorkbook.js';

export interface XlsxViewerHandle {
  goToSheet(sheet: number | string): void;
  workbook: Workbook | null;
}

export interface XlsxViewerProps {
  /** A File (e.g. from an <input>), ArrayBuffer, or Uint8Array. */
  src: WorkbookSource;
  /** Initial sheet name or 0-indexed position */
  initialSheet?: number | string;
  className?: string;
  style?: CSSProperties;
  onLoad?: (workbook: Workbook) => void;
  onError?: (error: Error) => void;
  onSheetChange?: (sheetName: string, index: number) => void;
}

export const XlsxViewer = forwardRef<XlsxViewerHandle, XlsxViewerProps>(function XlsxViewer(
  { src, initialSheet, className, style, onLoad, onError, onSheetChange },
  ref,
) {
  const { workbook, loading, error } = useWorkbook(src);
  const stageRef = useRef<HTMLDivElement>(null);
  const viewerRef = useRef<XlsxViewerController | null>(null);

  useEffect(() => {
    if (error) onError?.(error);
  }, [error, onError]);

  useEffect(() => {
    if (!workbook || !stageRef.current) return;
    onLoad?.(workbook);
    const viewer = new XlsxViewerController(workbook, stageRef.current, {
      initialSheet,
      onSheetChange,
    });
    viewerRef.current = viewer;
    return () => {
      viewer.destroy();
      viewerRef.current = null;
    };
  }, [workbook]); // eslint-disable-line react-hooks/exhaustive-deps

  useImperativeHandle(
    ref,
    () => ({
      goToSheet: (sheet: number | string) => viewerRef.current?.goToSheet(sheet),
      workbook,
    }),
    [workbook],
  );

  return (
    <div className={className} style={{ display: 'flex', flexDirection: 'column', width: '100%', height: '100%', ...style }}>
      <div ref={stageRef} style={{ flex: 1, minHeight: 0 }} />
      {loading && <div style={statusStyle}>Loading spreadsheet…</div>}
      {error && <div style={{ ...statusStyle, color: '#e57373' }}>Error: {error.message}</div>}
    </div>
  );
});

const statusStyle: CSSProperties = {
  padding: '8px 12px',
  font: '13px system-ui, sans-serif',
  color: '#ddd',
};
