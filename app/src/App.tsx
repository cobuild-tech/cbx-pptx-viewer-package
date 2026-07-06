import { Suspense, lazy, useCallback, useMemo, useRef, useState } from 'react';
import type { CSSProperties, DragEvent } from 'react';
import { PptxViewer } from '@cobuild-tech/pptx-viewer-react';
import { DocxViewer } from '@cobuild-tech/pptx-viewer-react';
import { InMemoryVersionStore } from '@cobuild-tech/pptx-viewer-core';

const PptxReactViewerWrap = lazy(() => import('./renderers/PptxReactViewerWrap'));
const PptxViewJsWrap = lazy(() => import('./renderers/PptxViewJsWrap'));
const CyntlerViewerWrap = lazy(() => import('./renderers/CyntlerViewerWrap'));

type FileType = 'pptx' | 'docx';
type PptxRendererId = 'mine' | 'pptx-react-viewer' | 'pptxviewjs' | 'cyntler';
type DocxRendererId = 'mine-docx';
type RendererId = PptxRendererId | DocxRendererId;

const PPTX_RENDERERS: { id: PptxRendererId; label: string }[] = [
  { id: 'mine', label: 'cbx-ppt-viewer' },
  { id: 'pptx-react-viewer', label: 'pptx-react-viewer' },
  { id: 'pptxviewjs', label: 'pptxviewjs (canvas)' },
  { id: 'cyntler', label: '@cyntler/react-doc-viewer' },
];

const DOCX_RENDERERS: { id: DocxRendererId; label: string }[] = [
  { id: 'mine-docx', label: 'cbx-doc-viewer' },
];

const FORMAT_META: Record<FileType, { ext: string; accept: string; icon: string; color: string; desc: string }> = {
  pptx: {
    ext: '.pptx',
    accept: '.pptx',
    icon: '📊',
    color: '#c43b1c',
    desc: 'PowerPoint Presentation',
  },
  docx: {
    ext: '.docx',
    accept: '.docx',
    icon: '📄',
    color: '#2b579a',
    desc: 'Word Document',
  },
};

function getFileType(file: File): FileType | null {
  const name = file.name.toLowerCase();
  if (name.endsWith('.pptx')) return 'pptx';
  if (name.endsWith('.docx')) return 'docx';
  return null;
}

export function App() {
  const [file, setFile] = useState<File | null>(null);
  const [mode, setMode] = useState<FileType>('pptx');
  const [renderer, setRenderer] = useState<RendererId>('mine');
  const [dragging, setDragging] = useState(false);
  const [editMode, setEditMode] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  // One version store for the session. Swap for a backend-folder adapter
  // (implements DocxVersionStore) to persist versions across reloads.
  const versionStore = useMemo(() => new InMemoryVersionStore(), []);

  const accept = useCallback((f: File | undefined) => {
    if (!f) return;
    const type = getFileType(f);
    if (!type) return;
    setFile(f);
    setMode(type);
    setRenderer(type === 'docx' ? 'mine-docx' : 'mine');
    setEditMode(false);
  }, []);

  const switchMode = useCallback((next: FileType) => {
    setMode(next);
    setFile(null);
    setRenderer(next === 'docx' ? 'mine-docx' : 'mine');
    setEditMode(false);
  }, []);

  const onDrop = useCallback(
    (e: DragEvent) => {
      e.preventDefault();
      setDragging(false);
      accept(e.dataTransfer.files?.[0]);
    },
    [accept],
  );

  const meta = FORMAT_META[mode];
  const renderers = mode === 'docx' ? DOCX_RENDERERS : PPTX_RENDERERS;

  return (
    <div style={page}>
      {/* ── Header ─────────────────────────────────────────────────────────── */}
      <header style={header}>
        {/* Logo / brand */}
        <div style={brand}>
          <span style={{ fontSize: 18, lineHeight: 1 }}>📁</span>
          <strong style={{ fontSize: 14, letterSpacing: '-0.3px' }}>cbx viewer</strong>
        </div>

        <div style={divider} />

        {/* Format tabs */}
        <div style={tabGroup}>
          {(['pptx', 'docx'] as FileType[]).map((fmt) => {
            const m = FORMAT_META[fmt];
            const active = mode === fmt;
            return (
              <button
                key={fmt}
                style={{
                  ...tabBtn,
                  background: active ? m.color : 'transparent',
                  color: active ? '#fff' : '#aaa',
                  borderColor: active ? m.color : '#444',
                }}
                onClick={() => switchMode(fmt)}
              >
                <span style={{ fontSize: 13 }}>{m.icon}</span>
                <span style={{ fontWeight: active ? 700 : 400 }}>{fmt.toUpperCase()}</span>
              </button>
            );
          })}
        </div>

        <div style={divider} />

        {/* Upload button */}
        <button
          style={{ ...uploadBtn, background: meta.color }}
          onClick={() => inputRef.current?.click()}
        >
          Upload {meta.ext}
        </button>

        {/* Renderer picker (only when file loaded) */}
        {file && renderers.length > 1 && (
          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13 }}>
            <span style={{ color: '#aaa' }}>Renderer:</span>
            <select
              style={select}
              value={renderer}
              onChange={(e) => setRenderer(e.target.value as RendererId)}
            >
              {renderers.map((r) => (
                <option key={r.id} value={r.id}>
                  {r.label}
                </option>
              ))}
            </select>
          </label>
        )}

        {/* Edit toggle (DOCX only) */}
        {file && mode === 'docx' && (
          <button
            style={{ ...uploadBtn, background: editMode ? '#1f7a3d' : '#3a3a3a' }}
            onClick={() => setEditMode((v) => !v)}
            title="Toggle inline editing"
          >
            {editMode ? '✓ Editing' : '✎ Edit'}
          </button>
        )}

        {/* File info chip */}
        {file && (
          <div style={fileChip}>
            <span style={{ ...modeBadge, background: meta.color }}>
              {meta.icon} {mode.toUpperCase()}
            </span>
            <span style={{ color: '#ddd', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {file.name}
            </span>
            <span style={{ color: '#666', flexShrink: 0 }}>
              {(file.size / 1024).toFixed(0)} KB
            </span>
            <button
              style={closeBtn}
              title="Close file"
              onClick={() => { setFile(null); }}
            >
              ✕
            </button>
          </div>
        )}

        <input
          ref={inputRef}
          type="file"
          accept={meta.accept}
          style={{ display: 'none' }}
          onChange={(e) => accept(e.target.files?.[0])}
        />
      </header>

      {/* ── Main ───────────────────────────────────────────────────────────── */}
      <main style={main}>
        {file ? (
          <Suspense fallback={<div style={loadingFallback}>Loading renderer…</div>}>
            {/* PPTX renderers */}
            {mode === 'pptx' && renderer === 'mine' && (
              <PptxViewer key={`pptx:${file.name}`} src={file} style={{ flex: 1, minHeight: 0 }} />
            )}
            {mode === 'pptx' && renderer === 'pptx-react-viewer' && (
              <PptxReactViewerWrap key={`prv:${file.name}`} file={file} />
            )}
            {mode === 'pptx' && renderer === 'pptxviewjs' && (
              <PptxViewJsWrap key={`pvjs:${file.name}`} file={file} />
            )}
            {mode === 'pptx' && renderer === 'cyntler' && (
              <CyntlerViewerWrap key={`cyn:${file.name}`} file={file} />
            )}

            {/* DOCX renderer — note: editMode is NOT in the key, so toggling edit
                mode does not remount/reload the doc (edits persist across modes). */}
            {mode === 'docx' && renderer === 'mine-docx' && (
              <DocxViewer
                key={`docx:${file.name}`}
                src={file}
                editable={editMode}
                versionStore={versionStore}
                onVersionSaved={() => setEditMode(false)}
                style={{ flex: 1, minHeight: 0 }}
              />
            )}
          </Suspense>
        ) : (
          /* ── Drop zone ─────────────────────────────────────────────────── */
          <div style={dropWrapper}>
            {/* Inline format switcher */}
            <div style={inlineTabGroup}>
              {(['pptx', 'docx'] as FileType[]).map((fmt) => {
                const m = FORMAT_META[fmt];
                const active = mode === fmt;
                return (
                  <button
                    key={fmt}
                    style={{
                      ...inlineTab,
                      borderColor: active ? m.color : '#444',
                      background: active ? `${m.color}22` : 'transparent',
                      color: active ? '#eee' : '#888',
                    }}
                    onClick={() => switchMode(fmt)}
                  >
                    <span style={{ fontSize: 24 }}>{m.icon}</span>
                    <span style={{ fontSize: 14, fontWeight: 600 }}>{fmt.toUpperCase()}</span>
                    <span style={{ fontSize: 12, color: active ? '#bbb' : '#666' }}>{m.desc}</span>
                  </button>
                );
              })}
            </div>

            {/* Drop area */}
            <div
              style={{
                ...dropZone,
                borderColor: dragging ? meta.color : '#444',
                background: dragging ? `${meta.color}11` : 'transparent',
              }}
              onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={onDrop}
              onClick={() => inputRef.current?.click()}
            >
              <div style={{ fontSize: 40, marginBottom: 12 }}>{meta.icon}</div>
              <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 6 }}>
                Drop a {meta.ext} file here
              </div>
              <div style={{ color: '#888', fontSize: 13, marginBottom: 16 }}>
                or click to browse
              </div>
              <div style={{ ...uploadBtn, background: meta.color, display: 'inline-block', cursor: 'pointer' }}>
                Choose {meta.ext} file
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

// ── Styles ────────────────────────────────────────────────────────────────────

const page: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  height: '100%',
  background: '#1a1a1a',
  color: '#eee',
};

const header: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 12,
  padding: '8px 16px',
  borderBottom: '1px solid #2e2e2e',
  background: '#202020',
  flexShrink: 0,
  flexWrap: 'wrap',
};

const brand: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  color: '#eee',
};

const divider: CSSProperties = {
  width: 1,
  height: 22,
  background: '#333',
  flexShrink: 0,
};

const tabGroup: CSSProperties = {
  display: 'flex',
  gap: 4,
};

const tabBtn: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 5,
  padding: '4px 11px',
  borderRadius: 6,
  border: '1px solid',
  fontSize: 12,
  cursor: 'pointer',
  transition: 'all 0.15s',
  letterSpacing: '0.3px',
};

const uploadBtn: CSSProperties = {
  padding: '6px 14px',
  borderRadius: 6,
  border: 'none',
  color: '#fff',
  fontWeight: 600,
  fontSize: 13,
  cursor: 'pointer',
  flexShrink: 0,
};

const select: CSSProperties = {
  padding: '4px 8px',
  borderRadius: 6,
  border: '1px solid #444',
  background: '#2e2e2e',
  color: '#eee',
  cursor: 'pointer',
  fontSize: 13,
};

const fileChip: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '3px 10px',
  borderRadius: 20,
  background: '#2a2a2a',
  border: '1px solid #383838',
  fontSize: 12,
  maxWidth: 380,
};

const modeBadge: CSSProperties = {
  padding: '1px 7px',
  borderRadius: 10,
  color: '#fff',
  fontSize: 10,
  fontWeight: 700,
  letterSpacing: '0.4px',
  flexShrink: 0,
};

const closeBtn: CSSProperties = {
  background: 'none',
  border: 'none',
  color: '#666',
  cursor: 'pointer',
  fontSize: 11,
  padding: '0 2px',
  lineHeight: 1,
  flexShrink: 0,
};

const main: CSSProperties = {
  flex: 1,
  minHeight: 0,
  display: 'flex',
};

const loadingFallback: CSSProperties = {
  margin: 'auto',
  color: '#888',
  fontSize: 14,
};

const dropWrapper: CSSProperties = {
  margin: 'auto',
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  gap: 24,
  padding: '40px 20px',
  width: '100%',
  maxWidth: 580,
};

const inlineTabGroup: CSSProperties = {
  display: 'flex',
  gap: 12,
  width: '100%',
};

const inlineTab: CSSProperties = {
  flex: 1,
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  gap: 4,
  padding: '16px 12px',
  borderRadius: 10,
  border: '1.5px solid',
  cursor: 'pointer',
  transition: 'all 0.15s',
};

const dropZone: CSSProperties = {
  width: '100%',
  padding: '36px 40px',
  border: '2px dashed',
  borderRadius: 12,
  textAlign: 'center',
  cursor: 'pointer',
  transition: 'all 0.15s',
};
