/**
 * The format-agnostic view of run formatting.
 *
 * PowerPoint and Word describe the same handful of character properties but
 * encode them completely differently — DrawingML uses attributes on `<a:rPr>`,
 * WordprocessingML uses child elements of `<w:rPr>` with half-point sizes. The
 * *value* is identical, so it lives here and each format slice owns only its
 * own encoding.
 */
export interface RunFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Font size in points. */
  sizePt?: number;
  /** sRGB hex, no leading '#'. */
  colorHex?: string;
  /** Typeface name. */
  font?: string;
}

/** True if the format carries no actual overrides. */
export function isEmptyFormat(f: RunFormat | undefined): boolean {
  return !f || Object.values(f).every((v) => v === undefined);
}

/** Merge `over` on top of `base` (later wins; undefined does not clear). */
export function mergeFormat(base: RunFormat | undefined, over: RunFormat | undefined): RunFormat {
  const out: RunFormat = { ...base };
  if (over) {
    for (const [k, v] of Object.entries(over)) {
      if (v !== undefined) (out as Record<string, unknown>)[k] = v;
    }
  }
  return out;
}
