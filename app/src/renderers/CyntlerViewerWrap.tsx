import { useEffect, useState } from 'react';
import DocViewer, { DocViewerRenderers } from '@cyntler/react-doc-viewer';
import '@cyntler/react-doc-viewer/dist/index.css';

/**
 * Wrapper for `@cyntler/react-doc-viewer`. Its .pptx support embeds the Microsoft
 * Office Online viewer, which can only fetch a PUBLICLY reachable URL — a local
 * `blob:` object URL won't load. So this renders correctly for hosted files but
 * shows the Office viewer's error for a locally-uploaded one. Included for
 * completeness of the comparison.
 */
export default function CyntlerViewerWrap({ file }: { file: File }) {
  const [uri, setUri] = useState<string>('');

  useEffect(() => {
    const url = URL.createObjectURL(file);
    setUri(url);
    return () => URL.revokeObjectURL(url);
  }, [file]);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', width: '100%' }}>
      <div style={banner}>
        Note: this package renders .pptx via the Microsoft Office Online viewer, which needs a
        public URL — a locally-uploaded file (blob URL) typically won't load.
      </div>
      <div style={{ flex: 1, minHeight: 0 }}>
        {uri && (
          <DocViewer
            documents={[{ uri, fileName: file.name, fileType: 'pptx' }]}
            pluginRenderers={DocViewerRenderers}
            style={{ height: '100%' }}
          />
        )}
      </div>
    </div>
  );
}

const banner: React.CSSProperties = {
  padding: '8px 12px',
  background: '#3a3320',
  color: '#e8d8a0',
  font: '12px system-ui, sans-serif',
  borderBottom: '1px solid #5a4f2a',
};
