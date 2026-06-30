import { useEffect, useState } from 'react';
import { PdfDocument } from '@pptx-viewer/core';

export type PdfSource = File | ArrayBuffer | Uint8Array | null | undefined;

export interface PdfState {
  doc: PdfDocument | null;
  loading: boolean;
  error: Error | null;
}

async function toBytes(
  src: Exclude<PdfSource, null | undefined>,
): Promise<ArrayBuffer | Uint8Array> {
  if (src instanceof File) return src.arrayBuffer();
  return src;
}

/**
 * Load a .pdf into a {@link PdfDocument}, re-loading when `src` changes and
 * disposing the previous document on change/unmount.
 */
export function usePdf(src: PdfSource): PdfState {
  const [state, setState] = useState<PdfState>({ doc: null, loading: false, error: null });

  useEffect(() => {
    if (!src) {
      setState({ doc: null, loading: false, error: null });
      return;
    }
    let cancelled = false;
    let loaded: PdfDocument | null = null;
    setState({ doc: null, loading: true, error: null });

    toBytes(src)
      .then((bytes) => PdfDocument.load(bytes))
      .then((doc) => {
        if (cancelled) { doc.dispose(); return; }
        loaded = doc;
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
