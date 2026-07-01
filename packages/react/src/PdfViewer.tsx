import { forwardRef, useCallback, useEffect, useImperativeHandle, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import { PdfViewer as PdfViewerCore, type PdfDocument, type PdfVersionStore } from '@pptx-viewer/core';
import { usePdf, type PdfSource } from './usePdf.js';

export interface PdfViewerHandle {
  next(): void;
  prev(): void;
  goTo(index: number): void;
  setEditable(enabled: boolean): void;
  doc: PdfDocument | null;
}

export interface PdfViewerProps {
  /** A File (e.g. from an <input>), ArrayBuffer, or Uint8Array. */
  src: PdfSource;
  /** Enable in-place text editing. Default false. */
  editable?: boolean;
  /** External version store (e.g. InMemoryPdfVersionStore). */
  versionStore?: PdfVersionStore;
  /** Show the built-in page indicator / edit toolbar. Default true. */
  toolbar?: boolean;
  className?: string;
  style?: CSSProperties;
  onLoad?: (doc: PdfDocument) => void;
  onError?: (error: Error) => void;
  onPageChange?: (index: number, count: number) => void;
  /** Fired after a version is saved successfully. */
  onVersionSaved?: () => void;
}

export const PdfViewer = forwardRef<PdfViewerHandle, PdfViewerProps>(function PdfViewer(
  {
    src,
    editable = false,
    versionStore,
    toolbar = true,
    className,
    style,
    onLoad,
    onError,
    onPageChange,
    onVersionSaved,
  },
  ref,
) {
  const { doc, loading, error } = usePdf(src);
  const stageRef  = useRef<HTMLDivElement>(null);
  const viewerRef = useRef<PdfViewerCore | null>(null);

  const [index,        setIndex]       = useState(0);
  const [count,        setCount]       = useState(0);
  const [editMode,     setEditMode]    = useState(false);
  const [hasChanges,   setHasChanges]  = useState(false);
  const [saving,       setSaving]      = useState(false);
  const [exportError,  setExportError] = useState<string | null>(null);

  useEffect(() => {
    if (error) onError?.(error);
  }, [error, onError]);

  useEffect(() => {
    if (!doc || !stageRef.current) return;

    if (versionStore) doc.configureVersioning(versionStore);

    onLoad?.(doc);

    const viewer = new PdfViewerCore(doc, stageRef.current, {
      onChange: (i, c) => {
        setIndex(i);
        setCount(c);
        onPageChange?.(i, c);
      },
      onEditModeChange: (enabled) => {
        setEditMode(enabled);
      },
    });

    doc.onChange = () => {
      setHasChanges(doc.hasUnsavedChanges);
    };

    viewerRef.current = viewer;
    return () => {
      doc.onChange = undefined;
      viewer.destroy();
      viewerRef.current = null;
    };
  }, [doc]); // eslint-disable-line react-hooks/exhaustive-deps

  const handleEditToggle = useCallback(() => {
    viewerRef.current?.setEditable(!editMode);
  }, [editMode]);

  const handleSaveVersion = useCallback(() => {
    if (!doc) return;
    const saved = doc.saveVersion();
    if (saved) onVersionSaved?.();
    setHasChanges(false);
  }, [doc, onVersionSaved]);

  const handleDownload = useCallback(async () => {
    if (!doc) return;
    setSaving(true);
    setExportError(null);
    try {
      const blob = await doc.exportBlob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href     = url;
      a.download = 'edited.pdf';
      a.click();
      // Revoke after a short delay so the browser has time to initiate the download.
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error('[PdfViewer] Export failed:', err);
      setExportError(msg);
    } finally {
      setSaving(false);
    }
  }, [doc]);

  useImperativeHandle(
    ref,
    () => ({
      next:        () => viewerRef.current?.next(),
      prev:        () => viewerRef.current?.prev(),
      goTo:        (i: number) => viewerRef.current?.goTo(i),
      setEditable: (enabled: boolean) => viewerRef.current?.setEditable(enabled),
      doc,
    }),
    [doc],
  );

  return (
    <div className={className} style={{ display: 'flex', flexDirection: 'column', ...style }}>
      {/* Top action bar — Edit / Done / Save Version / Download */}
      {toolbar && doc && editable && (
        <div style={actionBarStyle}>
          <button
            style={{
              ...btn,
              background: editMode ? '#1a5f9e' : '#3a3a3a',
              borderColor: editMode ? '#3a8fd4' : '#555',
            }}
            onClick={handleEditToggle}
            title={editMode ? 'Exit edit mode' : 'Enter edit mode to change text'}
          >
            {editMode ? 'Done' : 'Edit'}
          </button>

          {editMode && (
            <button
              style={{ ...btn, opacity: hasChanges ? 1 : 0.4 }}
              onClick={handleSaveVersion}
              disabled={!hasChanges}
              title="Save current edits as a version"
            >
              Save Version
            </button>
          )}

          <button
            style={btn}
            onClick={handleDownload}
            disabled={saving}
            title="Download PDF with edits applied"
          >
            {saving ? 'Saving…' : 'Download'}
          </button>
        </div>
      )}

      {/* Viewer — the core viewer mounts its own toolbar (zoom + format slot) here */}
      <div style={{ flex: 1, minHeight: 0, position: 'relative', background: '#525659' }}>
        <div ref={stageRef} style={{ width: '100%', height: '100%' }} />
      </div>

      {loading      && <div style={statusStyle}>Loading…</div>}
      {error        && <div style={{ ...statusStyle, color: '#e57373' }}>Load error: {error.message}</div>}
      {exportError  && (
        <div style={{ ...statusStyle, color: '#e57373', cursor: 'pointer' }} onClick={() => setExportError(null)}>
          Export failed: {exportError} &nbsp;✕
        </div>
      )}

      {/* Bottom navigation bar */}
      {toolbar && doc && (
        <div style={toolbarStyle}>
          <button
            style={btn}
            onClick={() => viewerRef.current?.prev()}
            disabled={index <= 0}
            title="Scroll to previous page"
          >
            ↑ Prev
          </button>
          <span style={{ minWidth: 90, textAlign: 'center' }}>
            Page {count === 0 ? 0 : index + 1} / {count}
          </span>
          <button
            style={btn}
            onClick={() => viewerRef.current?.next()}
            disabled={index >= count - 1}
            title="Scroll to next page"
          >
            Next ↓
          </button>
        </div>
      )}
    </div>
  );
});

// ── Styles ────────────────────────────────────────────────────────────────────

const statusStyle: CSSProperties = {
  padding: '8px 12px',
  font: '13px system-ui, sans-serif',
  color: '#ddd',
};

const actionBarStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '6px 12px',
  background: '#2a2a2a',
  color: '#eee',
  font: '13px system-ui, sans-serif',
  flexShrink: 0,
  borderBottom: '1px solid #444',
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
  flexShrink: 0,
};

const btn: CSSProperties = {
  padding: '4px 12px',
  borderRadius: 6,
  border: '1px solid #555',
  background: '#3a3a3a',
  color: '#eee',
  cursor: 'pointer',
};
