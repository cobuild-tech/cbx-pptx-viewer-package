import { type XmlNode, children, child, attr, attrNum, attrBool } from '../../oxml/xml.js';
import type {
  XlsxCellStyle,
  XlsxFont,
  XlsxFill,
  XlsxBorder,
  XlsxBorderSide,
  XlsxAlignment,
  BorderStyle,
} from '../model.js';

const BUILTIN_NUM_FMTS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  14: 'yyyy-mm-dd',
  49: '@',
};

function parseColor(node?: XmlNode): string | undefined {
  if (!node) return undefined;
  const rgb = attr(node, 'rgb');
  if (rgb) {
    // Standard Excel ARGB (8 chars) or RGB (6 chars)
    const clean = rgb.length === 8 ? rgb.slice(2) : rgb;
    return `#${clean.toLowerCase()}`;
  }
  return undefined;
}

export class XlsxStyles {
  private readonly numFmts = new Map<number, string>();
  private readonly fonts: XlsxFont[] = [];
  private readonly fills: XlsxFill[] = [];
  private readonly borders: XlsxBorder[] = [];
  private readonly cellStyles: XlsxCellStyle[] = [];

  constructor(stylesXml?: XmlNode) {
    if (!stylesXml) return;
    this.parseNumFmts(stylesXml);
    this.parseFonts(stylesXml);
    this.parseFills(stylesXml);
    this.parseBorders(stylesXml);
    this.parseCellXfs(stylesXml);
  }

  getStyle(styleId?: number): XlsxCellStyle | undefined {
    if (styleId === undefined) return undefined;
    return this.cellStyles[styleId];
  }

  private parseNumFmts(xml: XmlNode): void {
    const numFmtsNode = child(xml, 'numFmts');
    if (!numFmtsNode) return;
    for (const nf of children(numFmtsNode, 'numFmt')) {
      const id = attrNum(nf, 'numFmtId');
      const code = attr(nf, 'formatCode');
      if (id !== undefined && code) {
        this.numFmts.set(id, code);
      }
    }
  }

  private parseFonts(xml: XmlNode): void {
    const fontsNode = child(xml, 'fonts');
    if (!fontsNode) return;
    for (const f of children(fontsNode, 'font')) {
      const font: XlsxFont = {};
      const nameNode = child(f, 'name');
      if (nameNode) font.name = attr(nameNode, 'val');
      const szNode = child(f, 'sz');
      if (szNode) font.sizePt = attrNum(szNode, 'val');
      const colorNode = child(f, 'color');
      if (colorNode) font.colorHex = parseColor(colorNode);
      if (child(f, 'b')) font.bold = true;
      if (child(f, 'i')) font.italic = true;
      if (child(f, 'u')) font.underline = true;
      if (child(f, 'strike')) font.strike = true;
      this.fonts.push(font);
    }
  }

  private parseFills(xml: XmlNode): void {
    const fillsNode = child(xml, 'fills');
    if (!fillsNode) return;
    for (const fillNode of children(fillsNode, 'fill')) {
      const fill: XlsxFill = {};
      const pattern = child(fillNode, 'patternFill');
      if (pattern) {
        fill.patternType = attr(pattern, 'patternType');
        const fg = child(pattern, 'fgColor');
        if (fg) fill.fgColorHex = parseColor(fg);
        const bg = child(pattern, 'bgColor');
        if (bg) fill.bgColorHex = parseColor(bg);
      }
      this.fills.push(fill);
    }
  }

  private parseBorders(xml: XmlNode): void {
    const bordersNode = child(xml, 'borders');
    if (!bordersNode) return;
    for (const bNode of children(bordersNode, 'border')) {
      const border: XlsxBorder = {};
      const parseSide = (sideName: string): XlsxBorderSide | undefined => {
        const sideNode = child(bNode, sideName);
        if (!sideNode) return undefined;
        const style = attr(sideNode, 'style') as BorderStyle | undefined;
        if (!style || style === 'none') return undefined;
        const colorNode = child(sideNode, 'color');
        return { style, colorHex: parseColor(colorNode) ?? '#d4d4d4' };
      };

      border.top = parseSide('top');
      border.bottom = parseSide('bottom');
      border.left = parseSide('left');
      border.right = parseSide('right');
      this.borders.push(border);
    }
  }

  private parseCellXfs(xml: XmlNode): void {
    const xfsNode = child(xml, 'cellXfs');
    if (!xfsNode) return;
    for (const xf of children(xfsNode, 'xf')) {
      const style: XlsxCellStyle = {};
      const fontId = attrNum(xf, 'fontId');
      if (fontId !== undefined && this.fonts[fontId]) {
        style.font = this.fonts[fontId];
      }
      const fillId = attrNum(xf, 'fillId');
      if (fillId !== undefined && this.fills[fillId]) {
        style.fill = this.fills[fillId];
      }
      const borderId = attrNum(xf, 'borderId');
      if (borderId !== undefined && this.borders[borderId]) {
        style.border = this.borders[borderId];
      }
      const numFmtId = attrNum(xf, 'numFmtId');
      if (numFmtId !== undefined) {
        style.numFmtId = numFmtId;
        style.numFmtCode = this.numFmts.get(numFmtId) ?? BUILTIN_NUM_FMTS[numFmtId];
      }

      const alignNode = child(xf, 'alignment');
      if (alignNode) {
        const align: XlsxAlignment = {};
        const h = attr(alignNode, 'horizontal');
        if (h === 'left' || h === 'center' || h === 'right' || h === 'justify' || h === 'fill') {
          align.horizontal = h;
        }
        const v = attr(alignNode, 'vertical');
        if (v === 'top' || v === 'center' || v === 'bottom') {
          align.vertical = v;
        }
        if (attrBool(alignNode, 'wrapText') || attr(alignNode, 'wrapText') === '1') {
          align.wrapText = true;
        }
        style.alignment = align;
      }

      this.cellStyles.push(style);
    }
  }
}
