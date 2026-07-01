// ─────────────────────────────────────────────────────────────────────────────
// fonts.ts  –  Font metadata + Google Fonts loader
// No imports from the rest of the codebase.
// ─────────────────────────────────────────────────────────────────────────────

export type PdfFontFamily = 'helvetica' | 'times' | 'courier';
export type FontCategory  = 'sans-serif' | 'serif' | 'monospace' | 'display' | 'handwriting';

export interface FontDefinition {
  name:        string;
  cssStack:    string;
  googleFont?: string;
  pdfFamily:   PdfFontFamily;
  category:    FontCategory;
}

// ─── Font registry ────────────────────────────────────────────────────────────

export const FONTS: FontDefinition[] = [
  // ── Sans-serif ──────────────────────────────────────────────────────────────
  { name: 'Arial',         cssStack: 'Arial,Helvetica,sans-serif',               pdfFamily: 'helvetica', category: 'sans-serif' },
  { name: 'Inter',         cssStack: '"Inter",Arial,sans-serif',                  pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Inter' },
  { name: 'Roboto',        cssStack: '"Roboto",Arial,sans-serif',                 pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Roboto' },
  { name: 'Open Sans',     cssStack: '"Open Sans",Arial,sans-serif',              pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Open Sans' },
  { name: 'Lato',          cssStack: '"Lato",Arial,sans-serif',                   pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Lato' },
  { name: 'Montserrat',    cssStack: '"Montserrat",Arial,sans-serif',             pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Montserrat' },
  { name: 'Raleway',       cssStack: '"Raleway",Arial,sans-serif',                pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Raleway' },
  { name: 'Poppins',       cssStack: '"Poppins",Arial,sans-serif',                pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Poppins' },
  { name: 'Nunito',        cssStack: '"Nunito",Arial,sans-serif',                 pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Nunito' },
  { name: 'Ubuntu',        cssStack: '"Ubuntu",Arial,sans-serif',                 pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Ubuntu' },
  { name: 'Cabin',         cssStack: '"Cabin",Arial,sans-serif',                  pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Cabin' },
  { name: 'DM Sans',       cssStack: '"DM Sans",Arial,sans-serif',                pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'DM Sans' },
  { name: 'Work Sans',     cssStack: '"Work Sans",Arial,sans-serif',              pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Work Sans' },
  { name: 'Jost',          cssStack: '"Jost",Arial,sans-serif',                   pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Jost' },
  { name: 'Noto Sans',     cssStack: '"Noto Sans",Arial,sans-serif',              pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Noto Sans' },
  { name: 'Quicksand',     cssStack: '"Quicksand",Arial,sans-serif',              pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Quicksand' },
  { name: 'Oxygen',        cssStack: '"Oxygen",Arial,sans-serif',                 pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Oxygen' },
  { name: 'PT Sans',       cssStack: '"PT Sans",Arial,sans-serif',                pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'PT Sans' },
  { name: 'Source Sans 3', cssStack: '"Source Sans 3",Arial,sans-serif',          pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Source Sans 3' },
  { name: 'Figtree',       cssStack: '"Figtree",Arial,sans-serif',                pdfFamily: 'helvetica', category: 'sans-serif', googleFont: 'Figtree' },

  // ── Serif ────────────────────────────────────────────────────────────────────
  { name: 'Times New Roman',     cssStack: '"Times New Roman",Georgia,serif',            pdfFamily: 'times', category: 'serif' },
  { name: 'Georgia',             cssStack: 'Georgia,serif',                               pdfFamily: 'times', category: 'serif' },
  { name: 'Playfair Display',    cssStack: '"Playfair Display",Georgia,serif',            pdfFamily: 'times', category: 'serif', googleFont: 'Playfair Display' },
  { name: 'Merriweather',        cssStack: '"Merriweather",Georgia,serif',                pdfFamily: 'times', category: 'serif', googleFont: 'Merriweather' },
  { name: 'Lora',                cssStack: '"Lora",Georgia,serif',                        pdfFamily: 'times', category: 'serif', googleFont: 'Lora' },
  { name: 'PT Serif',            cssStack: '"PT Serif",Georgia,serif',                    pdfFamily: 'times', category: 'serif', googleFont: 'PT Serif' },
  { name: 'Libre Baskerville',   cssStack: '"Libre Baskerville",Georgia,serif',           pdfFamily: 'times', category: 'serif', googleFont: 'Libre Baskerville' },
  { name: 'Cormorant Garamond',  cssStack: '"Cormorant Garamond",Georgia,serif',          pdfFamily: 'times', category: 'serif', googleFont: 'Cormorant Garamond' },
  { name: 'EB Garamond',         cssStack: '"EB Garamond",Georgia,serif',                 pdfFamily: 'times', category: 'serif', googleFont: 'EB Garamond' },
  { name: 'Bitter',              cssStack: '"Bitter",Georgia,serif',                      pdfFamily: 'times', category: 'serif', googleFont: 'Bitter' },
  { name: 'Domine',              cssStack: '"Domine",Georgia,serif',                      pdfFamily: 'times', category: 'serif', googleFont: 'Domine' },
  { name: 'Noto Serif',          cssStack: '"Noto Serif",Georgia,serif',                  pdfFamily: 'times', category: 'serif', googleFont: 'Noto Serif' },
  { name: 'Crimson Text',        cssStack: '"Crimson Text",Georgia,serif',                pdfFamily: 'times', category: 'serif', googleFont: 'Crimson Text' },
  { name: 'Spectral',            cssStack: '"Spectral",Georgia,serif',                    pdfFamily: 'times', category: 'serif', googleFont: 'Spectral' },

  // ── Monospace ────────────────────────────────────────────────────────────────
  { name: 'Courier New',    cssStack: '"Courier New",Courier,monospace',          pdfFamily: 'courier', category: 'monospace' },
  { name: 'Roboto Mono',    cssStack: '"Roboto Mono",Courier,monospace',          pdfFamily: 'courier', category: 'monospace', googleFont: 'Roboto Mono' },
  { name: 'Source Code Pro',cssStack: '"Source Code Pro",Courier,monospace',      pdfFamily: 'courier', category: 'monospace', googleFont: 'Source Code Pro' },
  { name: 'Fira Code',      cssStack: '"Fira Code",Courier,monospace',            pdfFamily: 'courier', category: 'monospace', googleFont: 'Fira Code' },
  { name: 'JetBrains Mono', cssStack: '"JetBrains Mono",Courier,monospace',       pdfFamily: 'courier', category: 'monospace', googleFont: 'JetBrains Mono' },
  { name: 'Space Mono',     cssStack: '"Space Mono",Courier,monospace',           pdfFamily: 'courier', category: 'monospace', googleFont: 'Space Mono' },
  { name: 'Inconsolata',    cssStack: '"Inconsolata",Courier,monospace',          pdfFamily: 'courier', category: 'monospace', googleFont: 'Inconsolata' },
  { name: 'IBM Plex Mono',  cssStack: '"IBM Plex Mono",Courier,monospace',        pdfFamily: 'courier', category: 'monospace', googleFont: 'IBM Plex Mono' },

  // ── Display ──────────────────────────────────────────────────────────────────
  { name: 'Impact',       cssStack: 'Impact,"Arial Black",sans-serif',            pdfFamily: 'helvetica', category: 'display' },
  { name: 'Anton',        cssStack: '"Anton",Impact,sans-serif',                  pdfFamily: 'helvetica', category: 'display', googleFont: 'Anton' },
  { name: 'Bebas Neue',   cssStack: '"Bebas Neue",Impact,sans-serif',             pdfFamily: 'helvetica', category: 'display', googleFont: 'Bebas Neue' },
  { name: 'Oswald',       cssStack: '"Oswald",Arial,sans-serif',                  pdfFamily: 'helvetica', category: 'display', googleFont: 'Oswald' },
  { name: 'Righteous',    cssStack: '"Righteous",Arial,sans-serif',               pdfFamily: 'helvetica', category: 'display', googleFont: 'Righteous' },
  { name: 'Alfa Slab One',cssStack: '"Alfa Slab One",Georgia,serif',              pdfFamily: 'times',     category: 'display', googleFont: 'Alfa Slab One' },

  // ── Handwriting ──────────────────────────────────────────────────────────────
  { name: 'Caveat',          cssStack: '"Caveat",cursive',         pdfFamily: 'courier', category: 'handwriting', googleFont: 'Caveat' },
  { name: 'Dancing Script',  cssStack: '"Dancing Script",cursive', pdfFamily: 'courier', category: 'handwriting', googleFont: 'Dancing Script' },
  { name: 'Pacifico',        cssStack: '"Pacifico",cursive',       pdfFamily: 'courier', category: 'handwriting', googleFont: 'Pacifico' },
  { name: 'Indie Flower',    cssStack: '"Indie Flower",cursive',   pdfFamily: 'courier', category: 'handwriting', googleFont: 'Indie Flower' },
  { name: 'Satisfy',         cssStack: '"Satisfy",cursive',        pdfFamily: 'courier', category: 'handwriting', googleFont: 'Satisfy' },
  { name: 'Kaushan Script',  cssStack: '"Kaushan Script",cursive', pdfFamily: 'courier', category: 'handwriting', googleFont: 'Kaushan Script' },
];

// ─── Lookup helpers ───────────────────────────────────────────────────────────

/** Return the FontDefinition for a given display name, or undefined. */
export function findFont(name: string): FontDefinition | undefined {
  return FONTS.find(f => f.name === name);
}

/** Return the CSS font-family stack for a display name. Falls back to Arial. */
export function resolveCssFontStack(name: string): string {
  const found = findFont(name);
  if (found) return found.cssStack;
  // Legacy key names used in earlier annotation data.
  if (name === 'times')   return '"Times New Roman",Georgia,serif';
  if (name === 'courier') return '"Courier New",Courier,monospace';
  return 'Arial,Helvetica,sans-serif';
}

/** Map a font display name to its nearest PDF Standard14 family. */
export function resolvePdfFamily(name: string): PdfFontFamily {
  const found = findFont(name);
  if (found) return found.pdfFamily;
  // Legacy key names used in earlier annotation data.
  if (name === 'times' || name === 'courier') return name;
  return 'helvetica';
}

// ─── Google Fonts loader ──────────────────────────────────────────────────────

/** Tracks which Google Font family names have already been injected. */
const _loadedGoogleFonts = new Set<string>();

const BATCH_SIZE = 10;
const FONT_WEIGHTS = '0,400;0,700;1,400;1,700';

/**
 * Load Google Fonts via injected <link> tags (batched in groups of 10,
 * deduplicated across calls).
 */
export function ensureGoogleFontsLoaded(fonts: FontDefinition[]): void {
  if (typeof document === 'undefined') return;

  // Collect only the Google Fonts that haven't been loaded yet.
  const pending: string[] = [];
  for (const font of fonts) {
    if (font.googleFont && !_loadedGoogleFonts.has(font.googleFont)) {
      pending.push(font.googleFont);
      _loadedGoogleFonts.add(font.googleFont);
    }
  }

  if (pending.length === 0) return;

  // Inject one <link> per batch of BATCH_SIZE fonts.
  for (let i = 0; i < pending.length; i += BATCH_SIZE) {
    const batch = pending.slice(i, i + BATCH_SIZE);
    const familyParams = batch
      .map(name => `family=${encodeURIComponent(name)}:ital,wght@${FONT_WEIGHTS}`)
      .join('&');
    const href = `https://fonts.googleapis.com/css2?${familyParams}&display=swap`;

    const link = document.createElement('link');
    link.rel  = 'stylesheet';
    link.href = href;
    document.head.appendChild(link);
  }
}

/**
 * Load all fonts that belong to the given categories.
 * Pass no arguments to load every font in the registry.
 */
export function loadFontsByCategory(...categories: FontCategory[]): void {
  const target =
    categories.length === 0
      ? FONTS
      : FONTS.filter(f => categories.includes(f.category));

  ensureGoogleFontsLoaded(target);
}
