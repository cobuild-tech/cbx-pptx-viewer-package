/**
 * Unit conversion helpers specific to WordprocessingML.
 *
 * Word uses three unit systems that differ from DrawingML:
 *  - Twips  (twentieths of a point, 1/1440 inch) for page sizes, margins, indent
 *  - Half-points for font sizes  (<w:sz w:val="24"> = 12 pt)
 *  - Eighths of a point for border widths  (<w:sz w:val="4"> = 0.5 pt border)
 *
 * EMU conversions (used for DrawingML inline images) come from oxml/units.ts.
 */

/** 1 twip = 1/20 pt = 1/1440 inch = 96/1440 px at 96 DPI. */
export function twipsToPx(twips: number): number {
  return twips / 15;
}

/** Word font size attribute (<w:sz>) stores half-points. */
export function halfPtToPt(halfPt: number): number {
  return halfPt / 2;
}

/** Word border width (<w:sz>) stores eighths of a point. */
export function borderSzToPx(sz: number): number {
  return (sz / 8) * (96 / 72);
}
