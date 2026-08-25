import { Suspense, lazy, useCallback, useRef, useState } from 'react';
import type { CSSProperties, DragEvent } from 'react';
import { PptxViewer, DocxViewer, XlsxViewer } from '@cobuildx.ai/office-viewer/react';
import type { TextBoxOutline } from '@cobuildx.ai/office-viewer';

const PptxReactViewerWrap = lazy(() => import('./renderers/PptxReactViewerWrap'));
const PptxViewJsWrap = lazy(() => import('./renderers/PptxViewJsWrap'));
const CyntlerViewerWrap = lazy(() => import('./renderers/CyntlerViewerWrap'));

type RendererId = 'mine' | 'pptx-react-viewer' | 'pptxviewjs' | 'cyntler';
type Kind = 'pptx' | 'docx' | 'xlsx';

const PPTX_RENDERERS: { id: RendererId; label: string }[] = [
  { id: 'mine', label: 'cbx-ppt-viewer' },
  { id: 'pptx-react-viewer', label: 'pptx-react-viewer' },
  { id: 'pptxviewjs', label: 'pptxviewjs (canvas)' },
  { id: 'cyntler', label: '@cyntler/react-doc-viewer' },
];

const PPTX_COLOR = '#c43b1c';
const DOCX_COLOR = '#2b579a';
const XLSX_COLOR = '#107c41';

function kindOf(name: string): Kind | null {
  const n = name.toLowerCase();
  if (n.endsWith('.pptx')) return 'pptx';
  if (n.endsWith('.docx')) return 'docx';
  if (n.endsWith('.xlsx')) return 'xlsx';
  return null;
}

export function App() {
  const [file, setFile] = useState<File | null>(null);
  const [kind, setKind] = useState<Kind>('pptx');
  const [renderer, setRenderer] = useState<RendererId>('mine');
  const [editable, setEditable] = useState(false);
  const [outline, setOutline] = useState<TextBoxOutline>('hover');
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const accept = useCallback((f: File | undefined) => {
    if (!f) return;
    const k = kindOf(f.name);
    if (!k) return;
    setFile(f);
    setKind(k);
    setRenderer('mine');
    // A new file starts read-only, so nobody edits by accident on open.
    setEditable(false);
  }, []);

  const onDrop = useCallback(
    (e: DragEvent) => {
      e.preventDefault();
      setDragging(false);
      accept(e.dataTransfer.files?.[0]);
    },
    [accept],
  );

  const accent = kind === 'docx' ? DOCX_COLOR : kind === 'xlsx' ? XLSX_COLOR : PPTX_COLOR;

  return (
    <div style={page}>
      {/* ── Header ─────────────────────────────────────────────────────────── */}
      <header style={header}>
        <div style={brand}>
          <span style={{ fontSize: 18, lineHeight: 1 }}>📄</span>
          <strong style={{ fontSize: 14, letterSpacing: '-0.3px' }}>cbx viewer</strong>
        </div>

        <div style={divider} />

        <button style={{ ...uploadBtn, background: accent }} onClick={() => inputRef.current?.click()}>
          Upload .pptx / .docx / .xlsx
        </button>

        {file && kind === 'pptx' && (
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

        {file && ((kind === 'pptx' && renderer === 'mine') || kind === 'docx' || kind === 'xlsx') && (
          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13 }}>
            <input
              type="checkbox"
              checked={editable}
              onChange={(e) => setEditable(e.target.checked)}
              style={{ accentColor: accent }}
            />
            <span style={{ color: editable ? '#eee' : '#aaa' }}>
              {kind === 'xlsx' ? 'Edit cells' : 'Edit text'}
            </span>
          </label>
        )}

        {file && kind === 'pptx' && renderer === 'mine' && editable && (
          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13 }}>
            <span style={{ color: '#aaa' }}>Boxes:</span>
            <select
              style={select}
              value={outline}
              onChange={(e) => setOutline(e.target.value as TextBoxOutline)}
              title="Outline editable text boxes"
            >
              <option value="hover">On hover</option>
              <option value="always">Always</option>
              <option value="none">Only when focused</option>
            </select>
          </label>
        )}

        {file && (
          <div style={fileChip}>
            <span style={{ ...chipTag, background: accent }}>{kind.toUpperCase()}</span>
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
          accept=".pptx,.docx,.xlsx"
          style={{ display: 'none' }}
          onChange={(e) => accept(e.target.files?.[0])}
        />
      </header>

      {/* ── Main ───────────────────────────────────────────────────────────── */}
      <main style={main}>
        {file ? (
          <Suspense fallback={<div style={loadingFallback}>Loading renderer…</div>}>
            {kind === 'docx' && (
              <DocxViewer
                key={`docx:${file.name}:${editable ? 'edit' : 'read'}`}
                src={file}
                editable={editable}
                style={{ flex: 1, minHeight: 0 }}
              />
            )}
            {kind === 'xlsx' && (
              <XlsxViewer
                key={`xlsx:${file.name}:${editable ? 'edit' : 'read'}`}
                src={file}
                editable={editable}
                style={{ flex: 1, minHeight: 0 }}
              />
            )}
            {kind === 'pptx' && renderer === 'mine' && (
              <PptxViewer
                key={`pptx:${file.name}:${editable ? 'edit' : 'read'}`}
                src={file}
                editable={editable}
                textBoxOutline={outline}
                style={{ flex: 1, minHeight: 0 }}
              />
            )}
            {kind === 'pptx' && renderer === 'pptx-react-viewer' && <PptxReactViewerWrap key={`prv:${file.name}`} file={file} />}
            {kind === 'pptx' && renderer === 'pptxviewjs' && <PptxViewJsWrap key={`pvjs:${file.name}`} file={file} />}
            {kind === 'pptx' && renderer === 'cyntler' && <CyntlerViewerWrap key={`cyn:${file.name}`} file={file} />}
          </Suspense>
        ) : (
          <div style={dropWrapper}>
            <div
              style={{
                ...dropZone,
                borderColor: dragging ? accent : '#444',
                background: dragging ? `${accent}11` : 'transparent',
              }}
              onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={onDrop}
              onClick={() => inputRef.current?.click()}
            >
              <div style={{ fontSize: 40, marginBottom: 12 }}>📊</div>
              <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 6 }}>Drop a .pptx, .docx, or .xlsx file here</div>
              <div style={{ color: '#888', fontSize: 13, marginBottom: 16 }}>or click to browse</div>
              <div style={{ ...uploadBtn, background: accent, display: 'inline-block', cursor: 'pointer' }}>
                Choose a file
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
  maxWidth: 420,
};

const chipTag: CSSProperties = {
  fontSize: 10,
  fontWeight: 700,
  letterSpacing: '0.4px',
  color: '#fff',
  padding: '1px 6px',
  borderRadius: 4,
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

const dropZone: CSSProperties = {
  width: '100%',
  padding: '36px 40px',
  border: '2px dashed',
  borderRadius: 12,
  textAlign: 'center',
  cursor: 'pointer',
  transition: 'all 0.15s',
};
