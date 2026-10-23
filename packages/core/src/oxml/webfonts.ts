/**
 * Web-font resolution from Google Fonts, shared by every format.
 *
 * A document names fonts (Calibri, Open Sans…) that may not be installed on the
 * viewer's machine. The browser then substitutes a default (Arial/Helvetica)
 * with different glyph widths, so text wraps differently — a slide's text
 * overflows its box, a Word page breaks in a different place. To stay faithful
 * without requiring anything be installed, we fetch the fonts the document uses
 * from Google Fonts and register them under the document's own font names.
 *
 * Microsoft's fonts are proprietary and not on Google Fonts, so we map them to
 * their open, metric-compatible clones — the same substitution LibreOffice and
 * ChromeOS use (Calibri→Carlito, Cambria→Caladea, Arial→Arimo, Times New
 * Roman→Tinos, Courier New→Cousine, Georgia→Gelasio). Because the clones share
 * the originals' metrics, line wrapping matches Office. Real web fonts the
 * document uses (Open Sans, Lato…) are fetched as themselves.
 *
 * Only fonts that are actually used and not already available are fetched, and
 * only the glyphs needed are downloaded (Google serves subset woff2 lazily).
 */

export interface FontInstallation {
  /** Resolves once all loadable fonts are registered. */
  ready: Promise<void>;
  /** Remove the registered fonts from the document. */
  dispose(): void;
}

export const NO_FONTS: FontInstallation = { ready: Promise.resolve(), dispose() {} };

/** Microsoft / system fonts → open metric-compatible clone on Google Fonts. */
const SUBSTITUTES: Record<string, string> = {
  calibri: 'Carlito',
  'calibri light': 'Carlito',
  cambria: 'Caladea',
  'cambria math': 'Caladea',
  arial: 'Arimo',
  helvetica: 'Arimo',
  'times new roman': 'Tinos',
  times: 'Tinos',
  'courier new': 'Cousine',
  courier: 'Cousine',
  georgia: 'Gelasio',
};

/**
 * Common Google Fonts that documents use directly. We only request a
 * passthrough family if it's here, so we never fire (and never log a CORS error
 * for) a request for a font Google doesn't host — e.g. Aptos, Wingdings, or any
 * proprietary face. Map keys are lowercased; values are the canonical name to
 * request. Extend freely; unknown fonts simply fall back to the system font.
 */
const WEB_FONTS: string[] = [
  'Open Sans', 'Lato', 'Montserrat', 'Roboto', 'Roboto Condensed', 'Roboto Slab',
  'Roboto Mono', 'Source Sans Pro', 'Source Sans 3', 'Source Serif Pro', 'Source Code Pro',
  'Poppins', 'Nunito', 'Nunito Sans', 'Raleway', 'Oswald', 'Merriweather', 'Merriweather Sans',
  'PT Sans', 'PT Serif', 'Noto Sans', 'Noto Serif', 'Work Sans', 'Inter', 'Barlow',
  'Barlow Condensed', 'Mulish', 'Rubik', 'Karla', 'DM Sans', 'DM Serif Display', 'Manrope',
  'Quicksand', 'Josefin Sans', 'Playfair Display', 'Libre Franklin', 'Libre Baskerville',
  'Fira Sans', 'Fira Code', 'Cabin', 'Hind', 'Archivo', 'Archivo Narrow', 'Heebo',
  'Titillium Web', 'Exo', 'Exo 2', 'Dosis', 'Bitter', 'Crimson Text', 'Lora', 'Cormorant',
  'Cormorant Garamond', 'EB Garamond', 'Spectral', 'Vollkorn', 'Domine', 'Arvo', 'Fjalla One',
  'Francois One', 'Alfa Slab One', 'Righteous', 'Russo One', 'Sora', 'Space Grotesk',
  'Space Mono', 'Outfit', 'Plus Jakarta Sans', 'Figtree', 'Anton', 'Bebas Neue', 'Comfortaa',
  'Pacifico', 'Dancing Script', 'Lobster', 'Caveat', 'Shadows Into Light', 'Indie Flower',
  'Permanent Marker', 'Amatic SC', 'Teko', 'Saira', 'Saira Condensed', 'Catamaran', 'Maven Pro',
  'Asap', 'Signika', 'Questrial', 'Assistant', 'Varela Round', 'Yanone Kaffeesatz', 'Zilla Slab',
  'Abel', 'Acme', 'Bree Serif', 'Cairo', 'Chivo', 'Crete Round', 'Frank Ruhl Libre', 'Gelasio',
  'IBM Plex Sans', 'IBM Plex Serif', 'IBM Plex Mono', 'Kanit', 'Lexend', 'Marcellus',
  'Noto Sans Display', 'Overpass', 'Prompt', 'PT Mono', 'Public Sans', 'Red Hat Display',
  'Red Hat Text', 'Schibsted Grotesk', 'Slabo 27px', 'Tajawal', 'Urbanist', 'Albert Sans',
  'Be Vietnam Pro', 'Geist', 'Onest', 'Hanken Grotesk', 'Instrument Sans', 'Aleo', 'Karla',
  'Arimo', 'Tinos', 'Cousine', 'Carlito', 'Caladea',
];

/** lowercase family → canonical Google Fonts name, for passthrough requests. */
const WEB_FONT_CANONICAL = new Map(WEB_FONTS.map((n) => [n.toLowerCase(), n]));

/**
 * Resolve a document font to the Google Fonts family to request, or null to
 * skip. Office fonts map to a metric-compatible clone (which must then be
 * renamed to the document's font name); known web fonts pass through as
 * themselves.
 */
function resolveGoogleFamily(family: string): string | null {
  const lower = family.toLowerCase();
  return SUBSTITUTES[lower] ?? WEB_FONT_CANONICAL.get(lower) ?? null;
}

/** CSS generic families and system aliases we never fetch. */
const GENERIC = new Set([
  'serif',
  'sans-serif',
  'monospace',
  'cursive',
  'fantasy',
  'system-ui',
  'ui-sans-serif',
  'ui-serif',
  'ui-monospace',
  'ui-rounded',
  'inherit',
  'initial',
  '-apple-system',
  'blinkmacsystemfont',
]);

export interface WebFontOptions {
  /** Disable Google Fonts fetching entirely (e.g. for offline/privacy). */
  enabled?: boolean;
  /**
   * Builds the Google Fonts CSS URL for a family. Override to point at a
   * self-hosted mirror. Default uses fonts.googleapis.com.
   */
  cssUrl?: (googleFamily: string) => string;
}

/**
 * Fetch and register every family in `families` that isn't already available,
 * resolving once the needed faces have loaded so the caller can re-render with
 * correct metrics. `skip` names families handled elsewhere (embedded fonts).
 */
export function installWebFontFamilies(
  families: Iterable<string>,
  opts: WebFontOptions = {},
  skip: ReadonlySet<string> = new Set(),
): FontInstallation {
  if (
    opts.enabled === false ||
    typeof document === 'undefined' ||
    typeof fetch === 'undefined' ||
    typeof FontFace === 'undefined'
  ) {
    return NO_FONTS;
  }

  const wanted = [...new Set(families)].filter((fam) => {
    const lower = fam.toLowerCase();
    if (GENERIC.has(lower)) return false;
    if (skip.has(lower)) return false;
    return !isFontAvailable(fam); // already installed system-side
  });

  if (wanted.length === 0) return NO_FONTS;

  const cssUrl = opts.cssUrl ?? defaultCssUrl;
  const styles: HTMLStyleElement[] = [];

  const loads = wanted.map(async (family) => {
    const googleFamily = resolveGoogleFamily(family);
    if (!googleFamily) return; // not on Google Fonts — keep the system fallback
    try {
      const css = await fetchFamilyCss(googleFamily, family, cssUrl);
      if (!css) return;
      const el = document.createElement('style');
      el.dataset.cbxWebfont = family;
      el.textContent = css;
      document.head.appendChild(el);
      styles.push(el);
      // Force-load the faces we need now (downloads only the latin subset for
      // this probe), so the re-render reflows with the font present.
      await Promise.all([
        document.fonts.load(`16px "${cssEscape(family)}"`),
        document.fonts.load(`bold 16px "${cssEscape(family)}"`),
      ]).catch(() => {});
    } catch {
      // Unknown family, network/CORS error — fall back to the system font.
    }
  });

  return {
    ready: Promise.all(loads).then(() => undefined),
    dispose() {
      for (const el of styles) el.remove();
    },
  };
}

function defaultCssUrl(googleFamily: string): string {
  const name = googleFamily.trim().replace(/\s+/g, '+');
  // Request the four common variants; Google 400s if a family lacks them, which
  // we treat as "not available" and skip.
  return `https://fonts.googleapis.com/css2?family=${name}:ital,wght@0,400;0,700;1,400;1,700&display=swap`;
}

/**
 * Fetch the family's @font-face CSS from Google and, when we're substituting a
 * clone, rename the CSS family to the document's own name so its `font-family`
 * rules resolve to it.
 */
async function fetchFamilyCss(
  googleFamily: string,
  docFamily: string,
  cssUrl: (f: string) => string,
): Promise<string | null> {
  const res = await fetch(cssUrl(googleFamily));
  if (!res.ok) return null;
  let css = await res.text();
  if (googleFamily.toLowerCase() !== docFamily.toLowerCase()) {
    const re = new RegExp(`font-family:\\s*['"]${escapeRegExp(googleFamily)}['"]`, 'gi');
    css = css.replace(re, `font-family: '${docFamily.replace(/'/g, "\\'")}'`);
  }
  return css;
}

// ─── font availability detection ─────────────────────────────────────────────

let probeCtx: CanvasRenderingContext2D | null | undefined;
const PROBE_TEXT = 'mmmmmmmmmmlli wwWWgq019';

function getProbeCtx(): CanvasRenderingContext2D | null {
  if (probeCtx !== undefined) return probeCtx;
  try {
    probeCtx = document.createElement('canvas').getContext('2d');
  } catch {
    probeCtx = null;
  }
  return probeCtx;
}

/**
 * Whether `family` renders with its own glyphs (installed system-side or already
 * registered) rather than falling back. Compares the probe string's width in the
 * family against three generic baselines; any difference means it's available.
 */
export function isFontAvailable(family: string): boolean {
  const ctx = getProbeCtx();
  if (!ctx) return false;
  const name = `"${cssEscape(family)}"`;
  for (const base of ['monospace', 'sans-serif', 'serif']) {
    ctx.font = `16px ${base}`;
    const baseWidth = ctx.measureText(PROBE_TEXT).width;
    ctx.font = `16px ${name}, ${base}`;
    if (ctx.measureText(PROBE_TEXT).width !== baseWidth) return true;
  }
  return false;
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function cssEscape(s: string): string {
  return s.replace(/["\\]/g, '\\$&');
}
