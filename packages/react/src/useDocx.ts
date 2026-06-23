import { useEffect, useState } from 'react';
import { loadDocx, type DocxDocument } from '@pptx-viewer/core';

export type DocxSource = File | ArrayBuffer | Uint8Array | null | undefined;

export interface DocxState {
  document: DocxDocument | null;
  loading: boolean;
  error: Error | null;
}

async function toBytes(src: Exclude<DocxSource, null | undefined>): Promise<ArrayBuffer | Uint8Array> {
  if (src instanceof File) return src.arrayBuffer();
  return src;
}

/**
 * Load a .docx into a {@link DocxDocument}, re-loading when `src` changes and disposing
 * the previous document on change/unmount.
 */
export function useDocx(src: DocxSource): DocxState {
  const [state, setState] = useState<DocxState>({ document: null, loading: false, error: null });

  useEffect(() => {
    if (!src) {
      setState({ document: null, loading: false, error: null });
      return;
    }
    let cancelled = false;
    let loaded: DocxDocument | null = null;
    setState({ document: null, loading: true, error: null });

    toBytes(src)
      .then((bytes) => {
        if (cancelled) return;
        loaded = loadDocx(bytes);
        setState({ document: loaded, loading: false, error: null });
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        setState({ document: null, loading: false, error: err as Error });
      });

    return () => {
      cancelled = true;
      loaded?.dispose();
    };
  }, [src]);

  return state;
}
