import { useCallback, useRef, useState } from 'react';
import type { CSSProperties, DragEvent } from 'react';
import { PptxViewer } from '@pptx-viewer/react';

export function App() {
  const [file, setFile] = useState<File | null>(null);
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
          <PptxViewer src={file} fit="contain" style={{ height: '100%' }} />
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
