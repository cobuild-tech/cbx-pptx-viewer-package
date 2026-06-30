/**
 * PdfDocument loader — the top-level entry point for PDF files.
 *
 * Uses pdf.js to load and render PDF pages. The returned PdfDocument
 * exposes page metadata and on-demand canvas rendering, and must be
 * dispose()d to free pdf.js internal resources.
 *
 * Usage:
 *   const doc = await PdfDocument.load(arrayBuffer);
 *   doc.pages;               // PdfPage[] — natural dimensions
 *   await doc.renderPage(0, 2.0);  // HTMLCanvasElement at 2× render scale
 *   doc.dispose();
 */
import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy } from 'pdfjs-dist';
// Vite resolves this ?url import to the bundled worker asset URL.
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import type { PdfPage } from '../model.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

export class PdfDocument {
  /** Page metadata at scale 1.0 — available synchronously after load. */
  readonly pages: PdfPage[];
  private readonly proxy: PDFDocumentProxy;

  private constructor(proxy: PDFDocumentProxy, pages: PdfPage[]) {
    this.proxy = proxy;
    this.pages = pages;
  }

  /**
   * Load a PDF from raw bytes. Resolves with an immutable PdfDocument
   * containing page metadata. Call dispose() when done.
   */
  static async load(data: ArrayBuffer | Uint8Array): Promise<PdfDocument> {
    const typedData = data instanceof Uint8Array ? data : new Uint8Array(data);
    const proxy = await pdfjsLib.getDocument({ data: typedData }).promise;

    const pages: PdfPage[] = [];
    for (let i = 0; i < proxy.numPages; i++) {
      const pageProxy = await proxy.getPage(i + 1);
      const viewport = pageProxy.getViewport({ scale: 1.0 });
      pages.push({ index: i, widthPx: viewport.width, heightPx: viewport.height });
      pageProxy.cleanup();
    }

    return new PdfDocument(proxy, pages);
  }

  /**
   * Render a page to a canvas element at the given scale factor.
   *
   * The returned canvas has pixel dimensions: widthPx×scale × heightPx×scale.
   * The caller is responsible for setting CSS display dimensions as needed.
   *
   * @param index    0-based page index
   * @param scale    Render scale relative to natural size (1.0 = 72 DPI).
   *                 Use 2.0 for retina sharpness on standard displays.
   */
  async renderPage(index: number, scale: number): Promise<HTMLCanvasElement> {
    const pageProxy = await this.proxy.getPage(index + 1);
    const viewport = pageProxy.getViewport({ scale });

    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    canvas.style.display = 'block';

    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('Failed to acquire 2D canvas context.');

    await pageProxy.render({ canvasContext: ctx, viewport }).promise;
    pageProxy.cleanup();
    return canvas;
  }

  /** Release all pdf.js internal resources. Call when the viewer is destroyed. */
  dispose(): void {
    this.proxy.destroy();
  }
}
