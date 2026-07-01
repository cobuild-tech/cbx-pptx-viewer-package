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

/** A single text item extracted from a PDF page via pdf.js getTextContent(). */
export interface PdfTextItem {
  /** Unique stable ID: `"{pageIndex}-{itemIndex}"`. */
  id: string;
  pageIndex: number;
  /** Raw text string. */
  str: string;
  /** CSS x position (px) at scale 1.0 — top-left origin. */
  cssX: number;
  /** CSS y position (px) at scale 1.0 — top-left origin. */
  cssY: number;
  cssWidth: number;
  cssHeight: number;
  /** PDF user-space x — bottom-left origin, matches pdf-lib coordinates. */
  pdfX: number;
  /** PDF user-space y (baseline) — bottom-left origin, matches pdf-lib coordinates. */
  pdfY: number;
  pdfWidth: number;
  pdfHeight: number;
  /** Approximate font size in PDF points. */
  fontSize: number;
  fontName: string;
  /** The raw 6-element PDF transform matrix [a,b,c,d,e,f]. */
  transform: number[];
}

/**
 * A group of nearby text items that form a logical editable line/block.
 * Each block maps to one contenteditable div in edit mode.
 */
export interface PdfTextBlock {
  id: string;
  pageIndex: number;
  items: PdfTextItem[];
  /** Concatenated text of all items. */
  text: string;
  cssX: number;
  cssY: number;
  cssWidth: number;
  cssHeight: number;
  /** PDF-space bounding box (for export). */
  pdfX: number;
  pdfY: number;
  pdfWidth: number;
  pdfHeight: number;
  fontSize: number;
}

/**
 * A user-added free-form text annotation placed over the PDF.
 * Annotations are new content — they do not replace existing PDF text.
 */
export interface PdfAnnotation {
  id: string;
  pageIndex: number;
  /** CSS x position of the annotation's top-left corner (scale 1.0). */
  cssX: number;
  /** CSS y position of the annotation's top-left corner (scale 1.0). */
  cssY: number;
  /** Annotation box width in CSS px (also acts as min-width when editing). */
  width: number;
  /** Annotation box min-height in CSS px. Undefined = auto (content-driven). */
  height?: number;
  /** Font size in CSS px / PDF pt. */
  fontSize: number;
  /** Text content — may contain '\n' for multiline. */
  text: string;
  /** CSS color string for the rendered text. */
  color: string;
  fontWeight: 'normal' | 'bold';
  fontStyle:  'normal' | 'italic';
  /** Font display name (e.g. 'Inter', 'Playfair Display', 'Arial'). Legacy keys 'helvetica'/'times'/'courier' are also accepted. */
  fontFamily: string;
  textAlign:  'left' | 'center' | 'right';
}

/**
 * Style overrides for an existing PDF text block.
 * Only the fields present are applied; absent fields fall back to the original block values.
 */
export interface PdfBlockStyle {
  /** Override x position in CSS px. */
  cssX?:       number;
  /** Override y position in CSS px. */
  cssY?:       number;
  /** Override width in CSS px (also acts as min-width when editing). */
  cssWidth?:   number;
  /** Override min-height in CSS px. Undefined = auto (content-driven). */
  cssHeight?:  number;
  fontSize?:   number;
  color?:      string;
  fontWeight?: 'normal' | 'bold';
  fontStyle?:  'normal' | 'italic';
  /** Font display name (same pool as PdfAnnotation.fontFamily). */
  fontFamily?: string;
  textAlign?:  'left' | 'center' | 'right';
}

/** Invertible edit operations applied to the PDF model. */
export type PdfEditOp =
  | {
      kind: 'replaceText';
      blockId: string;
      pageIndex: number;
      oldText: string;
      newText: string;
    }
  | {
      kind: 'addAnnotation';
      annotation: PdfAnnotation;
    }
  | {
      kind: 'removeAnnotation';
      annotationId: string;
      pageIndex: number;
      /** Snapshot of the annotation at removal time (used to reconstruct on redo). */
      annotation: PdfAnnotation;
    }
  | {
      kind: 'updateAnnotation';
      annotationId: string;
      pageIndex: number;
      oldAnnotation: PdfAnnotation;
      newAnnotation: PdfAnnotation;
    }
  | {
      kind:      'styleBlock';
      blockId:   string;
      pageIndex: number;
      /** Previous complete style (empty object = no prior overrides). */
      oldStyle:  PdfBlockStyle;
      /** New complete style to apply. */
      newStyle:  PdfBlockStyle;
    };
