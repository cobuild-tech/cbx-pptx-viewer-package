import { Suspense, lazy, useCallback, useRef, useState } from 'react';
import type { CSSProperties, DragEvent } from 'react';
import { PptxViewer } from '@pptx-viewer/react';

const PptxReactViewerWrap = lazy(() => import('./renderers/PptxReactViewerWrap'));
const PptxViewJsWrap = lazy(() => import('./renderers/PptxViewJsWrap'));
const CyntlerViewerWrap = lazy(() => import('./renderers/CyntlerViewerWrap'));

type RendererId = 'mine' | 'pptx-react-viewer' | 'pptxviewjs' | 'cyntler';

const RENDERERS: { id: RendererId; label: string }[] = [
  { id: 'mine', label: 'cbx-ppt-viewer' },
  { id: 'pptx-react-viewer', label: 'pptx-react-viewer' },
  { id: 'pptxviewjs', label: 'pptxviewjs (canvas)' },
  { id: 'cyntler', label: '@cyntler/react-doc-viewer' },
];

export function App() {
  const [file, setFile] = useState<File | null>(null);
  const [renderer, setRenderer] = useState<RendererId>('mine');
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const accept = useCallback((f: File | undefined) => {
    if (f && f.name.toLowerCase().endsWith('.pptx')) setFile(f);
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
      <header style={header}>
        <strong style={{ fontSize: 15 }}>PPTX Viewer</strong>
        <button style={uploadBtn} onClick={() => inputRef.current?.click()}>
          Upload .pptx
        </button>
        <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13 }}>
          <span style={{ color: '#aaa' }}>Renderer:</span>
          <select
            style={select}
            value={renderer}
            onChange={(e) => setRenderer(e.target.value as RendererId)}
          >
            {RENDERERS.map((r) => (
              <option key={r.id} value={r.id}>
                {r.label}
              </option>
            ))}
          </select>
        </label>
        {file && (
          <span style={{ color: '#aaa', fontSize: 13 }}>
            {file.name} · {(file.size / 1024).toFixed(0)} KB
          </span>
        )}
        <input
          ref={inputRef}
          type="file"
          accept=".pptx"
          style={{ display: 'none' }}
          onChange={(e) => accept(e.target.files?.[0])}
        />
      </header>

      <main style={main}>
        {file ? (
          // Remount each renderer when the file or selection changes (key) so
          // imperative/canvas viewers fully reset.
          <Suspense fallback={<div style={{ margin: 'auto', color: '#ddd' }}>Loading renderer…</div>}>
            {renderer === 'mine' && (
              <PptxViewer key={`mine:${file.name}`} src={file} style={{ flex: 1, minHeight: 0 }} />
            )}
            {renderer === 'pptx-react-viewer' && (
              <PptxReactViewerWrap key={`prv:${file.name}`} file={file} />
            )}
            {renderer === 'pptxviewjs' && <PptxViewJsWrap key={`pvjs:${file.name}`} file={file} />}
            {renderer === 'cyntler' && <CyntlerViewerWrap key={`cyn:${file.name}`} file={file} />}
          </Suspense>
        ) : (
          <div
            style={{ ...dropZone, borderColor: dragging ? '#7aa2f7' : '#555' }}
            onDragOver={(e) => {
              e.preventDefault();
              setDragging(true);
            }}
            onDragLeave={() => setDragging(false)}
            onDrop={onDrop}
            onClick={() => inputRef.current?.click()}
          >
            <div style={{ fontSize: 17, marginBottom: 8 }}>Drop a .pptx file here</div>
            <div style={{ color: '#888' }}>or click to browse</div>
          </div>
        )}
      </main>
    </div>
  );
}

const page: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  height: '100%',
  background: '#1e1e1e',
  color: '#eee',
};

const header: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 16,
  padding: '10px 16px',
  borderBottom: '1px solid #333',
  background: '#252525',
};

const uploadBtn: CSSProperties = {
  padding: '6px 14px',
  borderRadius: 6,
  border: 'none',
  background: '#7aa2f7',
  color: '#10131c',
  fontWeight: 600,
  cursor: 'pointer',
};

const select: CSSProperties = {
  padding: '5px 8px',
  borderRadius: 6,
  border: '1px solid #555',
  background: '#3a3a3a',
  color: '#eee',
  cursor: 'pointer',
  fontSize: 13,
};

const main: CSSProperties = {
  flex: 1,
  minHeight: 0,
  display: 'flex',
};

const dropZone: CSSProperties = {
  margin: 'auto',
  width: 'min(640px, 80%)',
  padding: '60px 40px',
  border: '2px dashed #555',
  borderRadius: 12,
  textAlign: 'center',
  cursor: 'pointer',
  transition: 'border-color 0.15s',
};
