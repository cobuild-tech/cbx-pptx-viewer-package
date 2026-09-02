/**
 * The format-agnostic view of run and paragraph formatting.
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

/**
 * The format-agnostic view of *paragraph* formatting — the properties that
 * belong to a whole paragraph rather than to a stretch of characters.
 *
 * Names are deliberately neutral rather than DrawingML's: `align: 'justify'` is
 * `algn="just"` in PowerPoint and `<w:jc w:val="both"/>` in Word, and only the
 * format slice should have to know that. Values are **absolute, never deltas** —
 * a format is stamped on the DOM and read back after arbitrary editing, so
 * "one level deeper" is meaningless by the time it is applied; the caller
 * resolves a nudge against the paragraph's own level first.
 */
export interface ParaFormat {
  /** Bulleted, auto-numbered, or explicitly neither. */
  list?: 'bullet' | 'number' | 'none';
  /** Indent level, 0-based (PowerPoint allows nine). */
  level?: number;
  align?: 'left' | 'center' | 'right' | 'justify';
  /** Line spacing as a multiple of single: 1 = single, 1.5 = one and a half. */
  lineSpacingPct?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
}

/** True if the format carries no actual overrides. */
export function isEmptyFormat(f: RunFormat | ParaFormat | undefined): boolean {
  return !f || Object.values(f).every((v) => v === undefined);
}

/** Merge `over` on top of `base` (later wins; undefined does not clear). */
export function mergeFormat<T extends RunFormat | ParaFormat>(
  base: T | undefined,
  over: T | undefined,
): T {
  const out = { ...base } as T;
  if (over) {
    for (const [k, v] of Object.entries(over)) {
      if (v !== undefined) (out as Record<string, unknown>)[k] = v;
    }
  }
  return out;
}
