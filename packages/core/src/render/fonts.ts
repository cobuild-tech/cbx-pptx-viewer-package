/**
 * Embedded-font installation.
 *
 * This mirrors how PowerPoint for the web achieves font accuracy: it makes the
 * deck's actual fonts available to the browser. We read the fonts embedded in
 * the package and register each variant with the FontFace API under its real
 * typeface name, so text renders in the genuine font (correct metrics, wrapping,
 * and weights) instead of a substitute.
 *
 * PowerPoint embeds fonts as `.fntdata`. Mac/modern files store raw OpenType
 * (TTF/OTF) which the browser loads directly; Windows often wraps them in an
 * (uncompressed) EOT header, which we unwrap to the trailing sfnt data. Fonts in
 * a format the browser can't decode (e.g. MicroType-compressed EOT) are skipped
 * and fall back to the system/substitute font.
 */
import type { Deck } from '../parse/deck.js';

export interface FontInstallation {
  /** Resolves once all loadable embedded fonts are registered. */
  ready: Promise<void>;
  /** Remove the registered fonts from the document. */
  dispose(): void;
}

const NOOP: FontInstallation = { ready: Promise.resolve(), dispose() {} };

export function installDeckFonts(deck: Deck): FontInstallation {
  if (
    typeof FontFace === 'undefined' ||
    typeof document === 'undefined' ||
    deck.embeddedFonts.length === 0
  ) {
    return NOOP;
  }

  const added: FontFace[] = [];
  const loads: Promise<void>[] = [];

  for (const font of deck.embeddedFonts) {
    for (const face of font.faces) {
      const bytes = deck.fontBytes(face.part);
      if (!bytes) continue;
      try {
        const ff = new FontFace(font.typeface, toSfnt(bytes), {
          weight: String(face.weight),
          style: face.style,
        });
        loads.push(
          ff.load().then(
            () => {
              document.fonts.add(ff);
              added.push(ff);
            },
            () => {
              // Unsupported format — silently fall back to a substitute font.
            },
          ),
        );
      } catch {
        // Constructor can throw on a malformed source; ignore and fall back.
      }
    }
  }

  return {
    ready: Promise.all(loads).then(() => undefined),
    dispose() {
      for (const ff of added) document.fonts.delete(ff);
    },
  };
}

/** sfnt magic numbers for TrueType/OpenType. */
function looksLikeSfnt(b: Uint8Array): boolean {
  if (b.length < 4) return false;
  const tag = (b[0]! << 24) | (b[1]! << 16) | (b[2]! << 8) | b[3]!;
  return (
    tag === 0x00010000 || // TrueType
    tag === 0x4f54544f || // 'OTTO' (CFF)
    tag === 0x74727565 || // 'true'
    tag === 0x74746366 // 'ttcf' (collection)
  );
}

/**
 * Return raw sfnt bytes for FontFace as a clean ArrayBuffer. If the data is
 * already TTF/OTF, copy it through; if it's an uncompressed EOT (font data
 * appended after the header), extract the trailing sfnt; otherwise return as-is
 * and let FontFace decide.
 */
function toSfnt(bytes: Uint8Array): ArrayBuffer {
  if (looksLikeSfnt(bytes)) return copy(bytes);
  if (bytes.length > 16) {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const fontDataSize = dv.getUint32(4, true); // EOT FontDataSize, little-endian
    if (fontDataSize > 0 && fontDataSize <= bytes.byteLength) {
      const candidate = bytes.subarray(bytes.byteLength - fontDataSize);
      if (looksLikeSfnt(candidate)) return copy(candidate);
    }
  }
  return copy(bytes);
}

/** Copy into a fresh ArrayBuffer that FontFace can own. */
function copy(b: Uint8Array): ArrayBuffer {
  const out = new ArrayBuffer(b.byteLength);
  new Uint8Array(out).set(b);
  return out;
}
