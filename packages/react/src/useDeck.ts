import { useEffect, useState } from 'react';
import { loadPptx, type Deck } from '@pptx-viewer/core';

export type DeckSource = File | ArrayBuffer | Uint8Array | null | undefined;

export interface DeckState {
  deck: Deck | null;
  loading: boolean;
  error: Error | null;
}

async function toBytes(src: Exclude<DeckSource, null | undefined>): Promise<ArrayBuffer | Uint8Array> {
  if (src instanceof File) return src.arrayBuffer();
  return src;
}

/**
 * Load a .pptx into a {@link Deck}, re-loading when `src` changes and disposing
 * the previous deck (and its media object URLs) on change/unmount.
 */
export function useDeck(src: DeckSource): DeckState {
  const [state, setState] = useState<DeckState>({ deck: null, loading: false, error: null });

  useEffect(() => {
    if (!src) {
      setState({ deck: null, loading: false, error: null });
      return;
    }
    let cancelled = false;
    let loaded: Deck | null = null;
    setState({ deck: null, loading: true, error: null });

    toBytes(src)
      .then((bytes) => {
        if (cancelled) return;
        loaded = loadPptx(bytes);
        setState({ deck: loaded, loading: false, error: null });
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        setState({ deck: null, loading: false, error: err as Error });
      });

    return () => {
      cancelled = true;
      loaded?.dispose();
    };
  }, [src]);

  return state;
}
