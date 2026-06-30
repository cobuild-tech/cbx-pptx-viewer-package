import { useCallback, useEffect, useImperativeHandle, useRef, useState, forwardRef } from 'react';
import type { CSSProperties } from 'react';
import {
  DocxViewer as DocxViewerCore,
  type DocxDocument,
  type DocxVersionStore,
  type VersionMeta,
} from '@pptx-viewer/core';
import { useDocument, type DocxSource } from './useDocument.js';
import { DocxEditorToolbar } from './DocxEditorToolbar.js';

export interface DocxViewerHandle {
  next(): void;
  prev(): void;
  goTo(index: number): void;
  undo(): void;
  redo(): void;
  saveVersion(label?: string): Promise<VersionMeta | undefined>;
  restore(versionId: string): Promise<void>;
  /** Edited .docx as a Blob (for custom download/upload flows). */
  exportBlob(): Blob | undefined;
  doc: DocxDocument | null;
}

export interface DocxViewerProps {
  /** A File (e.g. from an <input>), ArrayBuffer, or Uint8Array. */
  src: DocxSource;
  /** Enable inline WYSIWYG editing + the editing toolbar. Default false. */
  editable?: boolean;
  /** Pluggable version store; enables Save/Restore when provided with editable. */
  versionStore?: DocxVersionStore;
  /** Stable id for this document within the version store. Default: file name. */
  docId?: string;
  /** Show the built-in page indicator toolbar. Default true. */
  toolbar?: boolean;
  className?: string;
  style?: CSSProperties;
  onLoad?: (doc: DocxDocument) => void;
  onError?: (error: Error) => void;
  onPageChange?: (index: number, count: number) => void;
  /** Fired after any edit is applied. */
  onEdit?: () => void;
  /** Fired after a version is saved. */
  onVersionSaved?: (meta: VersionMeta) => void;
}

export const DocxViewer = forwardRef<DocxViewerHandle, DocxViewerProps>(function DocxViewer(
  {
    src,
    editable = false,
    versionStore,
    docId,
    toolbar = true,
    className,
    style,
    onLoad,
    onError,
    onPageChange,
    onEdit,
    onVersionSaved,
  },
  ref,
) {
  const { doc, loading, error } = useDocument(src);
  const stageRef = useRef<HTMLDivElement>(null);
  const viewerRef = useRef<DocxViewerCore | null>(null);
  const [index, setIndex] = useState(0);
  const [count, setCount] = useState(0);
  const [editState, setEditState] = useState({ canUndo: false, canRedo: false, canSave: false });
  const [versions, setVersions] = useState<VersionMeta[]>([]);

  useEffect(() => {
    if (error) onError?.(error);
  }, [error, onError]);

  const refreshVersions = useCallback(() => {
    if (doc && versionStore) doc.listVersions().then(setVersions).catch(() => {});
  }, [doc, versionStore]);

  useEffect(() => {
    if (!doc || !stageRef.current) return;
    onLoad?.(doc);
    const viewer = new DocxViewerCore(doc, stageRef.current, {
      editable,
      onChange: (i, c) => {
        setIndex(i);
        setCount(c);
        onPageChange?.(i, c);
      },
    });
    viewerRef.current = viewer;

    // Track edit state on every change (in any mode) for toolbar enablement.
    const unsub = doc.onChange(() => {
      setEditState({ canUndo: doc.canUndo, canRedo: doc.canRedo, canSave: doc.hasUnsavedChanges });
      onEdit?.();
    });
    refreshVersions();

    return () => {
      unsub?.();
      viewer.destroy();
      viewerRef.current = null;
    };
  }, [doc]); // eslint-disable-line react-hooks/exhaustive-deps

  // Re-configure versioning whenever the store or docId prop changes so callers
  // that provide these props asynchronously (or swap stores) are handled correctly.
  useEffect(() => {
    if (!doc || !versionStore) return;
    doc.configureVersioning(versionStore, docId ?? fileName(src) ?? 'document');
    refreshVersions();
  }, [doc, versionStore, docId, src, refreshVersions]);

  // Toggle editing on the live viewer without recreating it (keeps edits in place).
  useEffect(() => {
    viewerRef.current?.setEditable(editable);
  }, [editable]);

  const saveVersion = useCallback(
    async (label = 'Untitled'): Promise<VersionMeta | undefined> => {
      if (!doc || !versionStore) return undefined;
      const meta = await doc.saveVersion(label, Date.now());
      setEditState({ canUndo: doc.canUndo, canRedo: doc.canRedo, canSave: doc.hasUnsavedChanges });
      refreshVersions();
      onVersionSaved?.(meta);
      return meta;
    },
    [doc, versionStore, refreshVersions, onVersionSaved],
  );

  const restore = useCallback(
    async (versionId: string) => {
      if (!doc) return;
      await doc.restore(versionId);
      setEditState({ canUndo: doc.canUndo, canRedo: doc.canRedo, canSave: doc.hasUnsavedChanges });
      refreshVersions();
    },
    [doc, refreshVersions],
  );

  const download = useCallback(() => {
    if (!doc) return;
    const blob = doc.exportBlob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = downloadName(src);
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }, [doc, src]);

  useImperativeHandle(
    ref,
    () => ({
      next: () => viewerRef.current?.next(),
      prev: () => viewerRef.current?.prev(),
      goTo: (i: number) => viewerRef.current?.goTo(i),
      undo: () => viewerRef.current?.undo(),
      redo: () => viewerRef.current?.redo(),
      saveVersion: (label?: string) => saveVersion(label),
      restore,
      exportBlob: () => doc?.exportBlob(),
      doc,
    }),
    [doc, saveVersion, restore],
  );

  return (
    <div className={className} style={{ display: 'flex', flexDirection: 'column', ...style }}>
      {editable && doc && (
        <DocxEditorToolbar
          canUndo={editState.canUndo}
          canRedo={editState.canRedo}
          canSave={editState.canSave}
          versions={versions}
          onUndo={() => viewerRef.current?.undo()}
          onRedo={() => viewerRef.current?.redo()}
          onToggle={(p) => viewerRef.current?.toggleFormat(p)}
          onColor={(hex) => viewerRef.current?.format({ color: hex })}
          onSize={(pt) => viewerRef.current?.format({ sizePt: pt })}
          onInsertParagraph={() => viewerRef.current?.insertParagraph()}
          onDeleteParagraph={() => viewerRef.current?.deleteParagraph()}
          onInsertRow={() => viewerRef.current?.insertTableRow()}
          onDeleteRow={() => viewerRef.current?.deleteTableRow()}
          onSaveVersion={(label) => void saveVersion(label)}
          onRestore={(id) => void restore(id)}
          onDownload={download}
        />
      )}

      {/*
        The viewer creates its own scroll container absolutely inside stageRef,
        so stageRef must fill the remaining height (not auto-grow).
      */}
      <div style={{ flex: 1, minHeight: 0, position: 'relative', background: '#525659' }}>
        <div ref={stageRef} style={{ width: '100%', height: '100%' }} />
      </div>

      {loading && <div style={statusStyle}>Loading…</div>}
      {error && <div style={{ ...statusStyle, color: '#e57373' }}>Error: {error.message}</div>}

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

function fileName(src: DocxSource): string | undefined {
  return src instanceof File ? src.name : undefined;
}

function downloadName(src: DocxSource): string {
  const name = fileName(src);
  if (!name) return 'document.docx';
  return name.toLowerCase().endsWith('.docx') ? name.replace(/\.docx$/i, ' (edited).docx') : `${name}.docx`;
}

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
