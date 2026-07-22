import { OpcPackage } from '../../oxml/package.js';
import { XlsxRelType } from '../relTypes.js';
import { XlsxStyles } from '../styles/styles.js';
import { parseSharedStrings, parseSheetXml } from '../sheets/sheet.js';
import { children, attr, child } from '../../oxml/xml.js';
import type { XlsxSheet, XlsxSheetSummary } from '../model.js';

export class Workbook {
  readonly sheetSummaries: XlsxSheetSummary[];
  private readonly sheetsCache = new Map<string, XlsxSheet>();
  private readonly pkg: OpcPackage;
  private readonly sharedStrings: string[];
  private readonly styles: XlsxStyles;

  private constructor(
    pkg: OpcPackage,
    sheetSummaries: XlsxSheetSummary[],
    sharedStrings: string[],
    styles: XlsxStyles,
  ) {
    this.pkg = pkg;
    this.sheetSummaries = sheetSummaries;
    this.sharedStrings = sharedStrings;
    this.styles = styles;
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
    const stylesRel = pkg.relByType(wbPart, XlsxRelType.Styles);
    const stylesXml = stylesRel ? pkg.getXml(stylesRel.target) : pkg.getXml('xl/styles.xml');
    const styles = new XlsxStyles(stylesXml);

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

    return new Workbook(pkg, sheetSummaries, sharedStrings, styles);
  }

  /**
   * Get a parsed XlsxSheet by index or sheet name/id (0-indexed or string).
   */
  getSheet(key: number | string = 0): XlsxSheet | undefined {
    let summary: XlsxSheetSummary | undefined;
    if (typeof key === 'number') {
      summary = this.sheetSummaries[key];
    } else {
      summary =
        this.sheetSummaries.find((s) => s.name === key || s.id === key) ??
        this.sheetSummaries[0];
    }

    if (!summary || !summary.targetPath) return undefined;
    if (this.sheetsCache.has(summary.id)) {
      return this.sheetsCache.get(summary.id);
    }

    const sheetXml = this.pkg.getXml(summary.targetPath);
    if (!sheetXml) return undefined;

    const parsed = parseSheetXml(summary.id, summary.name, sheetXml, this.sharedStrings, this.styles);
    this.sheetsCache.set(summary.id, parsed);
    return parsed;
  }

  dispose(): void {
    this.sheetsCache.clear();
  }
}
