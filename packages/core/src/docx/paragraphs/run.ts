/**
 * Run parser for WordprocessingML.
 *
 * Parses <w:r> elements (and the <w:rPr> run-properties block) into TextRun IR
 * nodes. Inherits base properties from the paragraph's resolved style, then
 * applies any run-level overrides.
 */
import { child, children, attr, attrNum, localName, type XmlNode } from '../../oxml/xml.js';
import { halfPtToPt } from '../units.js';
import type { TextRun, Color } from '../model.js';
import type { ResolvedRunStyle } from '../styles/styles.js';
import { mergeRunProps, StyleMap } from '../styles/styles.js';

/**
 * Parse a single <w:r> into a TextRun, inheriting from the resolved style.
 * Returns null if the run has no printable text (e.g. only a page break).
 */
export function parseRun(
  rEl: XmlNode,
  baseStyle: ResolvedRunStyle,
  resolveImage?: (relId: string) => string | undefined,
  styles?: StyleMap,
): TextRun | null {
  // Collect text from <w:t> children (may be multiple for preserved spaces).
  const textParts: string[] = [];
  for (const child_ of rEl.children) {
    const name = localName(child_.name);
    if (name === 't') {
      textParts.push(child_.text ?? '');
    } else if (name === 'br') {
      const type_ = attr(child_, 'w:type') ?? attr(child_, 'type');
      // Page/column breaks are handled at paragraph level; line breaks become \n.
      if (!type_ || type_ === 'textWrapping') textParts.push('\n');
    }
  }
  if (textParts.length === 0) return null;
  const text = textParts.join('');

  // Merge run properties on top of the inherited style.
  const style: ResolvedRunStyle = { ...baseStyle };
  const rPr = child(rEl, 'rPr');
  if (rPr && styles) {
    const rStyleEl = child(rPr, 'rStyle');
    if (rStyleEl) {
      const rStyleId = attr(rStyleEl, 'w:val') ?? attr(rStyleEl, 'val');
      if (rStyleId) {
        const charStyle = styles.get(rStyleId);
        // Merge character style run props (lower priority than explicit rPr)
        const cr = charStyle.run;
        if (cr.bold !== undefined) style.bold = cr.bold;
        if (cr.italic !== undefined) style.italic = cr.italic;
        if (cr.underline !== undefined) style.underline = cr.underline;
        if (cr.strike !== undefined) style.strike = cr.strike;
        if (cr.sizePt !== undefined) style.sizePt = cr.sizePt;
        if (cr.colorHex !== undefined) style.colorHex = cr.colorHex;
        if (cr.fontAscii !== undefined) style.fontAscii = cr.fontAscii;
        if (cr.fontHAnsi !== undefined) style.fontHAnsi = cr.fontHAnsi;
        if (cr.caps !== undefined) style.caps = cr.caps;
      }
    }
  }
  if (rPr) mergeRunProps(style, rPr);

  const run: TextRun = { text };
  if (style.bold) run.bold = true;
  if (style.italic) run.italic = true;
  if (style.underline) run.underline = true;
  if (style.strike) run.strike = true;
  if (style.sizePt !== undefined) run.sizePt = style.sizePt;
  if (style.colorHex) run.color = { hex: style.colorHex };
  if (style.fontAscii || style.fontHAnsi) run.font = style.fontAscii ?? style.fontHAnsi;
  if (style.caps) run.caps = style.caps;
  if (style.highlight) run.highlight = highlightNameToColor(style.highlight);
  if (style.letterSpacingPt !== undefined && style.letterSpacingPt !== 0) {
    run.letterSpacingPt = style.letterSpacingPt;
  }

  // Vertical alignment (super/subscript).
  if (rPr) {
    const vertEl = child(rPr, 'vertAlign');
    if (vertEl) {
      const val = attr(vertEl, 'w:val') ?? attr(vertEl, 'val');
      if (val === 'superscript') run.baseline = 30;
      if (val === 'subscript') run.baseline = -25;
    }
  }

  return run;
}

/**
 * Check if a run contains a page break (<w:br w:type="page">).
 * Used by the paragraph parser to split pages.
 */
export function runHasPageBreak(rEl: XmlNode): boolean {
  for (const child_ of rEl.children) {
    if (localName(child_.name) === 'br') {
      const type_ = attr(child_, 'w:type') ?? attr(child_, 'type');
      if (type_ === 'page') return true;
    }
  }
  return false;
}

const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: 'FFFF00',
  green: '00FF00',
  cyan: '00FFFF',
  magenta: 'FF00FF',
  blue: '0000FF',
  red: 'FF0000',
  darkBlue: '00008B',
  darkCyan: '008B8B',
  darkGreen: '006400',
  darkMagenta: '8B008B',
  darkRed: '8B0000',
  darkYellow: '808000',
  darkGray: 'A9A9A9',
  lightGray: 'D3D3D3',
  black: '000000',
  white: 'FFFFFF',
};

function highlightNameToColor(name: string): Color | undefined {
  const hex = HIGHLIGHT_COLORS[name];
  return hex ? { hex } : undefined;
}
