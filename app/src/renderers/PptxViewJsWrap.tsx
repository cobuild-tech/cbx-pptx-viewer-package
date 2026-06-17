import { useEffect, useRef, useState } from 'react';
import { PPTXViewer } from 'pptxviewjs';

/** Wrapper for `pptxviewjs` — a canvas renderer with an imperative API. */
export default function PptxViewJsWrap({ file }: { file: File }) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const viewerRef = useRef<PPTXViewer | null>(null);
  const [index, setIndex] = useState(0);
  const [count, setCount] = useState(0);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let alive = true;
    const canvas = canvasRef.current;
    if (!canvas) return;
    setError(null);

    const viewer = new PPTXViewer({ canvas, slideSizeMode: 'fit' });
    viewerRef.current = viewer;
    viewer
      .loadFile(file)
      .then(() => viewer.render(canvas))
      .then(() => {
        if (!alive) return;
        setCount(viewer.getSlideCount());
        setIndex(viewer.getCurrentSlideIndex());
      })
      .catch((e) => alive && setError(String(e?.message ?? e)));

    return () => {
      alive = false;
      viewer.destroy();
      viewerRef.current = null;
    };
  }, [file]);

  const go = (i: number) => {
    const viewer = viewerRef.current;
    if (!viewer) return;
    viewer.goToSlide(i, canvasRef.current).then(() => setIndex(viewer.getCurrentSlideIndex()));
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', width: '100%' }}>
      <div
        style={{
          flex: 1,
          minHeight: 0,
          overflow: 'auto',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          background: '#525659',
        }}
      >
        {error ? (
          <div style={{ color: '#e57373', padding: 16 }}>Error: {error}</div>
        ) : (
          <canvas ref={canvasRef} style={{ maxWidth: '100%', maxHeight: '100%' }} />
        )}
      </div>
      <div style={navBar}>
        <button style={navBtn} onClick={() => go(index - 1)} disabled={index <= 0}>
          ‹ Prev
        </button>
        <span style={{ minWidth: 90, textAlign: 'center' }}>
          Slide {count === 0 ? 0 : index + 1} / {count}
        </span>
        <button style={navBtn} onClick={() => go(index + 1)} disabled={index >= count - 1}>
          Next ›
        </button>
      </div>
    </div>
  );
}

const navBar: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  gap: 12,
  padding: '8px 12px',
  background: '#2a2a2a',
  color: '#eee',
  font: '13px system-ui, sans-serif',
};

const navBtn: React.CSSProperties = {
  padding: '4px 12px',
  borderRadius: 6,
  border: '1px solid #555',
  background: '#3a3a3a',
  color: '#eee',
  cursor: 'pointer',
};
