import { OpcPackage } from '../oxml/package.js';
import { child, children, attr, attrNum, path, type XmlNode } from '../oxml/xml.js';
import { emuToPx } from '../oxml/units.js';
import type {
  DocxDocumentModel,
  DocxElement,
  DocxParagraphElement,
  DocxRunElement,
  DocxTableElement,
  DocxTableRow,
  DocxTableCell,
  DocxImageElement,
} from './model.js';

const RelType_OfficeDocument =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

export class DocxDocument {
  readonly model: DocxDocumentModel;
  private readonly pkg: OpcPackage;
  private readonly documentPart: string;
  private readonly urlCache = new Map<string, string>();

  private constructor(pkg: OpcPackage, documentPart: string, model: DocxDocumentModel) {
    this.pkg = pkg;
    this.documentPart = documentPart;
    this.model = model;
  }

  static load(data: ArrayBuffer | Uint8Array): DocxDocument {
    const pkg = OpcPackage.load(data);
    const docPart = pkg.relByType('', RelType_OfficeDocument)?.target;
    if (!docPart) {
      throw new Error('Not a WordprocessingML package: no document part.');
    }

    const docXml = pkg.getXml(docPart);
    if (!docXml) {
      throw new Error('document.xml is missing or empty.');
    }

    const bodyNode = child(docXml, 'body');
    const bodyElements: DocxElement[] = [];

    if (bodyNode) {
      for (const node of bodyNode.children) {
        const local = node.name.split(':').pop();
        if (local === 'p') {
          bodyElements.push(parseParagraph(pkg, docPart, node));
        } else if (local === 'tbl') {
          bodyElements.push(parseTable(pkg, docPart, node));
        }
      }
    }

    const doc = new DocxDocument(pkg, docPart, { body: bodyElements });
    doc.resolveImages();
    return doc;
  }

  /** Object URL for an embedded media part (cached). Browser only. */
  imageUrl(part: string): string | undefined {
    const cached = this.urlCache.get(part);
    if (cached) return cached;
    const bytes = this.pkg.getBytes(part);
    if (!bytes || typeof URL === 'undefined' || typeof Blob === 'undefined') return undefined;
    const type = this.pkg.contentType(part) ?? 'application/octet-stream';
    const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
    this.urlCache.set(part, url);
    return url;
  }

  private resolveImages() {
    const traverse = (elements: DocxElement[]) => {
      for (const el of elements) {
        if (el.type === 'paragraph') {
          for (const run of el.runs) {
            if (run.image) {
              const rel = this.pkg.resolveRel(this.documentPart, run.image.relId);
              if (rel && rel.mode === 'Internal') {
                run.image.blobUrl = this.imageUrl(rel.target);
              }
            }
          }
        } else if (el.type === 'table') {
          for (const row of el.rows) {
            for (const cell of row.cells) {
              traverse(cell.content);
            }
          }
        }
      }
    };
    traverse(this.model.body);
  }

  dispose(): void {
    if (typeof URL !== 'undefined') {
      for (const url of this.urlCache.values()) {
        URL.revokeObjectURL(url);
      }
    }
    this.urlCache.clear();
  }
}

function parseParagraph(pkg: OpcPackage, docPart: string, node: XmlNode): DocxParagraphElement {
  const pPr = child(node, 'pPr');
  const alignmentVal = attr(child(pPr, 'jc'), 'val');
  let alignment: DocxParagraphElement['alignment'] = undefined;
  if (alignmentVal === 'left' || alignmentVal === 'center' || alignmentVal === 'right' || alignmentVal === 'both') {
    alignment = alignmentVal === 'both' ? 'justify' : alignmentVal;
  }

  const styleId = attr(child(pPr, 'pStyle'), 'val');
  const numPr = child(pPr, 'numPr');
  const isListItem = !!numPr;
  const listLevel = numPr ? (attrNum(child(numPr, 'ilvl'), 'val') ?? 0) : undefined;
  const listNumId = numPr ? attr(child(numPr, 'numId'), 'val') : undefined;

  const runs: DocxRunElement[] = [];

  const walkChildren = (parent: XmlNode, isHyperlink = false, hyperlinkUrl?: string) => {
    for (const c of parent.children) {
      const cLocal = c.name.split(':').pop();
      if (cLocal === 'r') {
        const run = parseRun(c);
        if (run) {
          if (isHyperlink) {
            run.isHyperlink = true;
            run.hyperlinkUrl = hyperlinkUrl;
          }
          runs.push(run);
        }
      } else if (cLocal === 'hyperlink') {
        const rId = attr(c, 'id') || attr(c, 'rId');
        let linkUrl = '';
        if (rId) {
          const rel = pkg.resolveRel(docPart, rId);
          if (rel) {
            linkUrl = rel.target;
          }
        }
        walkChildren(c, true, linkUrl);
      }
    }
  };

  walkChildren(node);

  return {
    type: 'paragraph',
    alignment,
    styleId,
    isListItem,
    listLevel,
    listNumId,
    runs,
  };
}

function parseRun(node: XmlNode): DocxRunElement | null {
  const rPr = child(node, 'rPr');
  const bold = child(rPr, 'b') !== undefined;
  const italic = child(rPr, 'i') !== undefined;
  const underline = child(rPr, 'u') !== undefined;
  const strike = child(rPr, 'strike') !== undefined;
  const color = attr(child(rPr, 'color'), 'val');
  const szVal = attrNum(child(rPr, 'sz'), 'val');
  const fontSize = szVal !== undefined ? szVal / 2 : undefined; // half-points -> pt
  const fontFamily = attr(child(rPr, 'rFonts'), 'ascii');

  // Check for text
  let text = '';
  for (const c of node.children) {
    const local = c.name.split(':').pop();
    if (local === 't') {
      text += c.text;
    } else if (local === 'tab') {
      text += '\t';
    } else if (local === 'br') {
      text += '\n';
    }
  }

  // Check for drawing (image)
  let image: DocxImageElement | undefined = undefined;
  const drawing = child(node, 'drawing');
  if (drawing) {
    const inline = child(drawing, 'inline') || child(drawing, 'anchor');
    if (inline) {
      const extent = child(inline, 'extent');
      const cx = attrNum(extent, 'cx');
      const cy = attrNum(extent, 'cy');
      const width = cx !== undefined ? emuToPx(cx) : undefined;
      const height = cy !== undefined ? emuToPx(cy) : undefined;

      const blip = path(inline, 'graphic/graphicData/pic/blipFill/blip');
      const relId = attr(blip, 'embed') || attr(blip, 'r:embed');

      const docPr = child(inline, 'docPr');
      const altText = attr(docPr, 'descr') || attr(docPr, 'title') || '';

      if (relId) {
        image = {
          relId,
          width,
          height,
          altText,
        };
      }
    }
  }

  if (text === '' && !image) {
    return null;
  }

  return {
    type: 'run',
    text,
    bold,
    italic,
    underline,
    strike,
    color,
    fontSize,
    fontFamily,
    image,
  };
}

function parseTable(pkg: OpcPackage, docPart: string, node: XmlNode): DocxTableElement {
  const rows: DocxTableRow[] = [];
  for (const rowNode of children(node, 'tr')) {
    const cells: DocxTableCell[] = [];
    for (const cellNode of children(rowNode, 'tc')) {
      const tcPr = child(cellNode, 'tcPr');
      const colSpan = attrNum(child(tcPr, 'gridSpan'), 'val') ?? 1;
      const shadingColor = attr(child(tcPr, 'shd'), 'fill');

      const content: DocxElement[] = [];
      for (const nestedNode of cellNode.children) {
        const nestedLocal = nestedNode.name.split(':').pop();
        if (nestedLocal === 'p') {
          content.push(parseParagraph(pkg, docPart, nestedNode));
        } else if (nestedLocal === 'tbl') {
          content.push(parseTable(pkg, docPart, nestedNode));
        }
      }

      cells.push({
        colSpan,
        shadingColor: shadingColor && shadingColor !== 'auto' ? `#${shadingColor}` : undefined,
        content,
      });
    }
    rows.push({ cells });
  }

  return {
    type: 'table',
    rows,
  };
}
