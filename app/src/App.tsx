import { Suspense, lazy, useCallback, useRef, useState } from 'react';
import type { CSSProperties, DragEvent } from 'react';
import { PptxViewer } from '@cobuildx.ai/office-viewer/react';

const PptxReactViewerWrap = lazy(() => import('./renderers/PptxReactViewerWrap'));
const PptxViewJsWrap = lazy(() => import('./renderers/PptxViewJsWrap'));
const CyntlerViewerWrap = lazy(() => import('./renderers/CyntlerViewerWrap'));

type RendererId = 'mine' | 'pptx-react-viewer' | 'pptxviewjs' | 'cyntler';

const PPTX_RENDERERS: { id: RendererId; label: string }[] = [
  { id: 'mine', label: 'cbx-ppt-viewer' },
  { id: 'pptx-react-viewer', label: 'pptx-react-viewer' },
  { id: 'pptxviewjs', label: 'pptxviewjs (canvas)' },
  { id: 'cyntler', label: '@cyntler/react-doc-viewer' },
];

const PPTX_COLOR = '#c43b1c';

export function App() {
  const [file, setFile] = useState<File | null>(null);
  const [renderer, setRenderer] = useState<RendererId>('mine');
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const accept = useCallback((f: File | undefined) => {
    if (!f || !f.name.toLowerCase().endsWith('.pptx')) return;
    setFile(f);
    setRenderer('mine');
  }, []);

  const onDrop = useCallback(
    (e: DragEvent) => {
      e.preventDefault();
      setDragging(false);
      accept(e.dataTransfer.files?.[0]);
    },
    [accept],
  );

  return (
    <div style={page}>
      {/* ── Header ─────────────────────────────────────────────────────────── */}
      <header style={header}>
        <div style={brand}>
          <span style={{ fontSize: 18, lineHeight: 1 }}>📊</span>
          <strong style={{ fontSize: 14, letterSpacing: '-0.3px' }}>cbx viewer</strong>
        </div>

        <div style={divider} />

        <button
          style={{ ...uploadBtn, background: PPTX_COLOR }}
          onClick={() => inputRef.current?.click()}
        >
          Upload .pptx
        </button>

        {file && (
          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13 }}>
            <span style={{ color: '#aaa' }}>Renderer:</span>
            <select
              style={select}
              value={renderer}
              onChange={(e) => setRenderer(e.target.value as RendererId)}
            >
              {PPTX_RENDERERS.map((r) => (
                <option key={r.id} value={r.id}>
                  {r.label}
                </option>
              ))}
            </select>
          </label>
        )}

        {file && (
          <div style={fileChip}>
            <span style={{ color: '#ddd', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {file.name}
            </span>
            <span style={{ color: '#666', flexShrink: 0 }}>{(file.size / 1024).toFixed(0)} KB</span>
            <button style={closeBtn} title="Close file" onClick={() => setFile(null)}>
              ✕
            </button>
          </div>
        )}

        <input
          ref={inputRef}
          type="file"
          accept=".pptx"
          style={{ display: 'none' }}
          onChange={(e) => accept(e.target.files?.[0])}
        />
      </header>

      {/* ── Main ───────────────────────────────────────────────────────────── */}
      <main style={main}>
        {file ? (
          <Suspense fallback={<div style={loadingFallback}>Loading renderer…</div>}>
            {renderer === 'mine' && (
              <PptxViewer key={`pptx:${file.name}`} src={file} style={{ flex: 1, minHeight: 0 }} />
            )}
            {renderer === 'pptx-react-viewer' && <PptxReactViewerWrap key={`prv:${file.name}`} file={file} />}
            {renderer === 'pptxviewjs' && <PptxViewJsWrap key={`pvjs:${file.name}`} file={file} />}
            {renderer === 'cyntler' && <CyntlerViewerWrap key={`cyn:${file.name}`} file={file} />}
          </Suspense>
        ) : (
          <div style={dropWrapper}>
            <div
              style={{
                ...dropZone,
                borderColor: dragging ? PPTX_COLOR : '#444',
                background: dragging ? `${PPTX_COLOR}11` : 'transparent',
              }}
              onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={onDrop}
              onClick={() => inputRef.current?.click()}
            >
              <div style={{ fontSize: 40, marginBottom: 12 }}>📊</div>
              <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 6 }}>Drop a .pptx file here</div>
              <div style={{ color: '#888', fontSize: 13, marginBottom: 16 }}>or click to browse</div>
              <div style={{ ...uploadBtn, background: PPTX_COLOR, display: 'inline-block', cursor: 'pointer' }}>
                Choose .pptx file
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

const dropZone: CSSProperties = {
  width: '100%',
  padding: '36px 40px',
  border: '2px dashed',
  borderRadius: 12,
  textAlign: 'center',
  cursor: 'pointer',
  transition: 'all 0.15s',
};
