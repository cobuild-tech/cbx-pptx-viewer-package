/**
 * Render-agnostic PDF model.
 *
 * All dimensions are in CSS pixels at scale 1.0 (72 DPI / 1 CSS px per PDF pt).
 * Actual canvas render resolution is managed by the viewer and document renderer.
 */

export interface PdfPage {
  /** 0-based page index. */
  index: number;
  /** Page width in CSS pixels at scale 1.0. */
  widthPx: number;
  /** Page height in CSS pixels at scale 1.0. */
  heightPx: number;
}
