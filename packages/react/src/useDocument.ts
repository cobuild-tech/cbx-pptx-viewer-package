import { useEffect, useState } from 'react';
import { loadDocx, type DocxDocument } from '@cobuild-tech/pptx-viewer-core';

export type DocxSource = File | ArrayBuffer | Uint8Array | null | undefined;

export interface DocumentState {
  doc: DocxDocument | null;
  loading: boolean;
  error: Error | null;
}

async function toBytes(
  src: Exclude<DocxSource, null | undefined>,
): Promise<ArrayBuffer | Uint8Array> {
  if (src instanceof File) return src.arrayBuffer();
  return src;
}

/**
 * Load a .docx into a {@link DocxDocument}, re-loading when `src` changes and
 * disposing the previous document (and its media object URLs) on change/unmount.
 */
export function useDocument(src: DocxSource): DocumentState {
  const [state, setState] = useState<DocumentState>({ doc: null, loading: false, error: null });

  useEffect(() => {
    if (!src) {
      setState({ doc: null, loading: false, error: null });
      return;
    }
    let cancelled = false;
    let loaded: DocxDocument | null = null;
    setState({ doc: null, loading: true, error: null });

    toBytes(src)
      .then((bytes) => {
        if (cancelled) return;
        loaded = loadDocx(bytes);
        setState({ doc: loaded, loading: false, error: null });
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        setState({ doc: null, loading: false, error: err as Error });
      });

    return () => {
      cancelled = true;
      loaded?.dispose();
    };
  }, [src]);

  return state;
}
