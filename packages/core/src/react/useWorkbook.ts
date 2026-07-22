import { useEffect, useState } from 'react';
import { loadXlsx, type Workbook } from '../index.js';

export type WorkbookSource = File | ArrayBuffer | Uint8Array | null | undefined;

export interface WorkbookState {
  workbook: Workbook | null;
  loading: boolean;
  error: Error | null;
}

async function toBytes(
  src: Exclude<WorkbookSource, null | undefined>,
): Promise<ArrayBuffer | Uint8Array> {
  if (src instanceof File) return src.arrayBuffer();
  return src;
}

/**
 * Load a .xlsx into a {@link Workbook}, re-loading when `src` changes and
 * disposing the previous workbook on change/unmount.
 */
export function useWorkbook(src: WorkbookSource): WorkbookState {
  const [state, setState] = useState<WorkbookState>({ workbook: null, loading: false, error: null });

  useEffect(() => {
    if (!src) {
      setState({ workbook: null, loading: false, error: null });
      return;
    }
    let cancelled = false;
    let loaded: Workbook | null = null;
    setState({ workbook: null, loading: true, error: null });

    toBytes(src)
      .then((bytes) => {
        if (cancelled) return;
        loaded = loadXlsx(bytes);
        setState({ workbook: loaded, loading: false, error: null });
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        setState({ workbook: null, loading: false, error: err as Error });
      });

    return () => {
      cancelled = true;
      loaded?.dispose();
    };
  }, [src]);

  return state;
}
