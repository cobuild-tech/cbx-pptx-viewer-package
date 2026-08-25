/**
 * Workbook loader — the top-level entry point for SpreadsheetML.
 *
 * XLSX pipeline: read (OPC) -> parse (XML -> model) -> render (DOM)
 *
 * Sheets are parsed lazily and cached, because a workbook is usually opened on
 * one sheet and the others may be large. Editing follows the same architecture
 * as `pptx/edit` and `docx/edit`: every parsed cell records the `<c>` node it
 * came from (into a `WeakMap`, so the model itself stays a pure value tree), and
 * a commit mutates that XML, marks the part dirty and re-parses the sheet.
 */
import { OpcPackage } from '../../oxml/package.js';
import { XlsxRelType } from '../relTypes.js';
import { XlsxStyles } from '../styles/styles.js';
import { parseSharedStrings, parseSheetXml } from '../sheets/sheet.js';
import { children, attr, child, serializeXml, createElement, type XmlNode } from '../../oxml/xml.js';
import type { XlsxSheet, XlsxSheetSummary, XlsxCell } from '../model.js';

const XLSX_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const STYLES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
const SML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** Where a parsed model object came from in the package. */
export interface XlsxSource {
  /** The `<c>` element the cell was parsed from. */
  node: XmlNode;
  /** Worksheet part path, e.g. "xl/worksheets/sheet1.xml". */
  part: string;
}

export class Workbook {
  readonly sheetSummaries: XlsxSheetSummary[];
  private readonly sheetsCache = new Map<string, XlsxSheet>();
  private readonly pkg: OpcPackage;
  private readonly wbPart: string;
  private sharedStrings: string[];
  private styles: XlsxStyles;
  private stylesPartPath: string | undefined;
  /**
   * Model object -> the XML node it was parsed from. Off the model itself so
   * the render-agnostic types stay pure; a WeakMap means entries die with the
   * sheet when it is re-parsed.
   */
  private sources = new WeakMap<object, XlsxSource>();

  private constructor(
    pkg: OpcPackage,
    wbPart: string,
    sheetSummaries: XlsxSheetSummary[],
    sharedStrings: string[],
    styles: XlsxStyles,
    stylesPartPath: string | undefined,
  ) {
    this.pkg = pkg;
    this.wbPart = wbPart;
    this.sheetSummaries = sheetSummaries;
    this.sharedStrings = sharedStrings;
    this.styles = styles;
    this.stylesPartPath = stylesPartPath;
  }

  static load(data: ArrayBuffer | Uint8Array): Workbook {
    const pkg = OpcPackage.load(data);
    const wbPart =
      pkg.relByType('', XlsxRelType.OfficeDocument)?.target ??
      (pkg.has('xl/workbook.xml') ? 'xl/workbook.xml' : undefined);

    if (!wbPart) {
      throw new Error('Not a SpreadsheetML package: no workbook part found.');
    }

    const wbXml = pkg.getXml(wbPart);
    if (!wbXml) {
      throw new Error('xl/workbook.xml is missing or empty.');
    }

    // Styles & Shared strings
    const stylesPath =
      pkg.relByType(wbPart, XlsxRelType.Styles)?.target ??
      (pkg.has('xl/styles.xml') ? 'xl/styles.xml' : undefined);
    const styles = new XlsxStyles(stylesPath ? pkg.getXml(stylesPath) : undefined);

    const stringsRel = pkg.relByType(wbPart, XlsxRelType.SharedStrings);
    const stringsXml = stringsRel ? pkg.getXml(stringsRel.target) : pkg.getXml('xl/sharedStrings.xml');
    const sharedStrings = parseSharedStrings(stringsXml);

    // Sheets list
    const sheetSummaries: XlsxSheetSummary[] = [];
    const sheetsNode = child(wbXml, 'sheets');
    if (sheetsNode) {
      for (const sNode of children(sheetsNode, 'sheet')) {
        const name = attr(sNode, 'name') ?? 'Sheet';
        const id = attr(sNode, 'sheetId') ?? `${sheetSummaries.length + 1}`;
        const rId = attr(sNode, 'id') ?? '';
        const rel = pkg.resolveRel(wbPart, rId);

        let targetPath = rel?.target;
        if (!targetPath) {
          const fallback = `xl/worksheets/sheet${id}.xml`;
          if (pkg.has(fallback)) targetPath = fallback;
        }

        sheetSummaries.push({ id, name, rId, targetPath });
      }
    }

    return new Workbook(pkg, wbPart, sheetSummaries, sharedStrings, styles, stylesPath);
  }

  /** Resolve a sheet key (0-indexed position, name, or sheetId) to its summary. */
  summaryOf(key: number | string = 0): XlsxSheetSummary | undefined {
    if (typeof key === 'number') return this.sheetSummaries[key];
    return (
      this.sheetSummaries.find((s) => s.name === key || s.id === key) ?? this.sheetSummaries[0]
    );
  }

  /**
   * Get a parsed XlsxSheet by index or sheet name/id (0-indexed or string).
   */
  getSheet(key: number | string = 0): XlsxSheet | undefined {
    const summary = this.summaryOf(key);
    if (!summary || !summary.targetPath) return undefined;
    const cached = this.sheetsCache.get(summary.id);
    if (cached) return cached;
    return this.parseSheet(summary.id);
  }

  private parseSheet(sheetId: string): XlsxSheet | undefined {
    const summary = this.sheetSummaries.find((s) => s.id === sheetId);
    if (!summary?.targetPath) return undefined;

    const sheetXml = this.pkg.getXml(summary.targetPath);
    if (!sheetXml) return undefined;

    const part = summary.targetPath;
    const parsed = parseSheetXml(
      summary.id,
      summary.name,
      sheetXml,
      this.sharedStrings,
      this.styles,
      (cell: XlsxCell, node: XmlNode) => this.sources.set(cell, { node, part }),
    );
    this.sheetsCache.set(summary.id, parsed);
    return parsed;
  }

  // ─── Editing ───────────────────────────────────────────────────────────────

  /** The `<c>` element a parsed cell came from, with its worksheet part path. */
  sourceOf(model: object): XlsxSource | undefined {
    return this.sources.get(model);
  }

  /** Worksheet part path for a sheet id, if the sheet resolves to a part. */
  sheetPart(sheetId: string): string | undefined {
    return this.sheetSummaries.find((s) => s.id === sheetId)?.targetPath;
  }

  /** The workbook part path (`xl/workbook.xml`). */
  get workbookPart(): string {
    return this.wbPart;
  }

  /** Parsed XML root of the workbook part (the live, mutable node). */
  workbookXml(): XmlNode | undefined {
    return this.pkg.getXml(this.wbPart);
  }

  /** Parsed XML root of a worksheet part (the live, mutable node). */
  sheetXml(sheetId: string): XmlNode | undefined {
    const part = this.sheetPart(sheetId);
    return part ? this.pkg.getXml(part) : undefined;
  }

  /** The styles part path, creating a minimal `xl/styles.xml` if absent. */
  stylesPart(): string {
    if (this.stylesPartPath) return this.stylesPartPath;
    const path = 'xl/styles.xml';
    this.pkg.setPart(path, serializeXml(minimalStyleSheet()));
    this.addRelationship(this.wbPart, XlsxRelType.Styles, 'styles.xml');
    this.addContentTypeOverride(`/${path}`, STYLES_CONTENT_TYPE);
    this.stylesPartPath = path;
    return path;
  }

  /** Parsed XML root of the styles part (the live, mutable node). */
  stylesXml(): XmlNode | undefined {
    return this.pkg.getXml(this.stylesPart());
  }

  /** Re-parse the style table — call after mutating `xl/styles.xml`. */
  reloadStyles(): void {
    this.styles = new XlsxStyles(
      this.stylesPartPath ? this.pkg.getXml(this.stylesPartPath) : undefined,
    );
  }

  /**
   * Re-parse one sheet from its (possibly mutated) XML, replacing the cached
   * model. `pkg.getXml` returns the cached node the edit layer mutated, so this
   * picks up edits without re-reading the zip.
   *
   * **Every model reference into that sheet goes stale.** Fresh cells are built
   * and the source map entries for the old ones are dropped, so callers must
   * re-read from the returned sheet after each commit.
   */
  rebuildSheet(sheetId: string): XlsxSheet | undefined {
    this.reloadStyles();
    this.sheetsCache.delete(sheetId);
    return this.parseSheet(sheetId);
  }

  /** Drop every cached sheet so the next read re-parses (used after undo). */
  invalidate(): void {
    this.reloadStyles();
    this.sheetsCache.clear();
  }

  /** Mark a part as mutated in place so export re-serializes it. */
  markDirty(part: string): void {
    this.pkg.markDirty(part);
  }

  /** True if any part has been edited. */
  get hasEdits(): boolean {
    return this.pkg.hasEdits;
  }

  /** Current XML text of a part — used to snapshot for undo. */
  snapshotPart(part: string): string | undefined {
    return this.pkg.serializePart(part);
  }

  /** Restore a part from a snapshot taken by {@link snapshotPart}. */
  restorePart(part: string, xml: string): void {
    this.pkg.setPart(part, xml);
  }

  /** Re-zip the workbook, edits included, as .xlsx bytes. */
  toBytes(): Uint8Array {
    return this.pkg.toBytes();
  }

  /** Re-zip the workbook, edits included, as a .xlsx Blob. */
  exportBlob(): Blob {
    return this.pkg.toBlob(XLSX_CONTENT_TYPE);
  }

  /** Append a relationship to a source part's .rels, returning its new id. */
  private addRelationship(sourcePart: string, type: string, target: string): string {
    const dir = sourcePart.includes('/') ? sourcePart.slice(0, sourcePart.lastIndexOf('/')) : '';
    const base = dir ? sourcePart.slice(dir.length + 1) : sourcePart;
    const relsPath = `${dir ? dir + '/' : ''}_rels/${base}.rels`;
    const relsXml =
      this.pkg.getXml(relsPath) ??
      createElement('Relationships', {
        xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships',
      });

    let n = relsXml.children.length + 1;
    const used = new Set(relsXml.children.map((c) => attr(c, 'Id')));
    while (used.has(`rId${n}`)) n++;
    const id = `rId${n}`;
    relsXml.children.push(createElement('Relationship', { Id: id, Type: type, Target: target }));
    // setPart (rather than markDirty) so the relationship cache is dropped.
    this.pkg.setPart(relsPath, serializeXml(relsXml));
    return id;
  }

  /** Declare a part's content type, if `[Content_Types].xml` is present. */
  private addContentTypeOverride(partName: string, contentType: string): void {
    const ct = this.pkg.getXml('[Content_Types].xml');
    if (!ct) return;
    if (ct.children.some((c) => attr(c, 'PartName') === partName)) return;
    ct.children.push(
      createElement('Override', { PartName: partName, ContentType: contentType }),
    );
    this.pkg.setPart('[Content_Types].xml', serializeXml(ct));
  }

  dispose(): void {
    this.sheetsCache.clear();
  }
}

/** The smallest valid style sheet: one default font, fill, border and xf. */
function minimalStyleSheet(): XmlNode {
  return createElement('styleSheet', { xmlns: SML_NS }, [
    createElement('fonts', { count: '1' }, [
      createElement('font', {}, [
        createElement('sz', { val: '11' }),
        createElement('name', { val: 'Calibri' }),
      ]),
    ]),
    createElement('fills', { count: '2' }, [
      createElement('fill', {}, [createElement('patternFill', { patternType: 'none' })]),
      createElement('fill', {}, [createElement('patternFill', { patternType: 'gray125' })]),
    ]),
    createElement('borders', { count: '1' }, [createElement('border')]),
    createElement('cellStyleXfs', { count: '1' }, [
      createElement('xf', { numFmtId: '0', fontId: '0', fillId: '0', borderId: '0' }),
    ]),
    createElement('cellXfs', { count: '1' }, [
      createElement('xf', { numFmtId: '0', fontId: '0', fillId: '0', borderId: '0', xfId: '0' }),
    ]),
  ]);
}
