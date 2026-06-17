import { useEffect, useState } from 'react';
import { PowerPointViewer } from 'pptx-react-viewer';
import 'pptx-react-viewer/styles';

/** Wrapper for `pptx-react-viewer` (full-featured editor; consumes Uint8Array). */
export default function PptxReactViewerWrap({ file }: { file: File }) {
  const [bytes, setBytes] = useState<Uint8Array | null>(null);

  useEffect(() => {
    let alive = true;
    setBytes(null);
    file.arrayBuffer().then((ab) => {
      if (alive) setBytes(new Uint8Array(ab));
    });
    return () => {
      alive = false;
    };
  }, [file]);

  if (!bytes) return <div style={{ padding: 16, color: '#ddd' }}>Loading…</div>;

  return (
    <div style={{ height: '100%', width: '100%' }}>
      <PowerPointViewer content={bytes} canEdit={false} />
    </div>
  );
}
