/**
 * Unit conversions for the OOXML/DrawingML coordinate system.
 *
 * PowerPoint stores geometry in EMU (English Metric Units) and font sizes in
 * hundredths of a point. We render in a CSS-pixel base coordinate space at
 * 96 DPI and let the viewer apply `transform: scale()` to fit the viewport,
 * so absolute positions stay exact and everything scales proportionally.
 */

/** EMU per inch. */
export const EMU_PER_INCH = 914400;
/** EMU per point (1/72 inch). */
export const EMU_PER_POINT = 12700;
/** CSS pixels per inch (the canonical screen DPI). */
export const PX_PER_INCH = 96;
/** EMU per CSS pixel at 96 DPI. */
export const EMU_PER_PX = EMU_PER_INCH / PX_PER_INCH; // 9525

/** Convert EMU to CSS pixels (96 DPI base space). */
export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

/** Convert CSS pixels (96 DPI base space) back to EMU, rounded to a whole unit.
 * EMU is an integer type in OOXML, so this is where an edited geometry lands on
 * the grid; going back through {@link emuToPx} is exact to within half an EMU
 * (about 5e-5 px), so repeated edits cannot drift. */
export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX);
}

/** Convert degrees to the `rot` attribute's 60000ths of a degree. */
export function degToAngle(deg: number): number {
  return Math.round(deg * 60000);
}

/** Convert EMU to points. */
export function emuToPt(emu: number): number {
  return emu / EMU_PER_POINT;
}

/** Convert points to CSS pixels (96 DPI). */
export function ptToPx(pt: number): number {
  return (pt / 72) * PX_PER_INCH;
}

/**
 * Convert a DrawingML font size (`sz`, in hundredths of a point) to points.
 * e.g. sz="1800" -> 18pt.
 */
export function fontSizeToPt(sz: number): number {
  return sz / 100;
}

/**
 * Convert a 60000ths-of-a-degree angle (`rot` attribute) to degrees.
 * e.g. rot="5400000" -> 90deg.
 */
export function angleToDeg(rot: number): number {
  return rot / 60000;
}

/**
 * Convert a DrawingML percentage value to a 0..1 fraction.
 * DrawingML stores percentages as 1000ths of a percent, so 100% = 100000.
 */
export function pctToFraction(pct: number): number {
  return pct / 100000;
}
