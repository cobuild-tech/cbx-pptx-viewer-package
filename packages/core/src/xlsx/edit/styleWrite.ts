/**
 * Write-back for `xl/styles.xml`.
 *
 * A spreadsheet cell has no inline formatting: `<c s="3">` points at the third
 * entry of `<cellXfs>`, which in turn points at a font, a fill, a border and a
 * number format. Formatting a cell therefore means *interning* — derive the new
 * font (or fill, or xf) from the one the cell already uses, look for an
 * identical entry that already exists, and only append when there is none.
 * Without that, every keystroke on the toolbar would grow the style table, and
 * Excel caps it.
 */
import {
  attr,
  child,
  children,
  cloneNode,
  createElement,
  localName,
  serializeNode,
  type XmlNode,
} from '../../oxml/xml.js';
import type { RunFormat } from '../../oxml/edit/format.js';
import type { XlsxAlignment } from '../model.js';

/** CT_Stylesheet's element order — new containers must respect it. */
const STYLESHEET_ORDER = [
  'numFmts',
  'fonts',
  'fills',
  'borders',
  'cellStyleXfs',
  'cellXfs',
  'cellStyles',
  'dxfs',
  'tableStyles',
  'colors',
  'extLst',
];

/** CT_Font's element order. Excel is lenient about it; Excel Online is not. */
const FONT_ORDER = [
  'b',
  'i',
  'strike',
  'condense',
  'extend',
  'outline',
  'shadow',
  'u',
  'vertAlign',
  'sz',
  'color',
  'name',
  'family',
  'charset',
  'scheme',
];

/** The first number format id free for custom codes; 0..163 are built in. */
const FIRST_CUSTOM_NUM_FMT = 164;

/** Everything a toolbar can change about one cell. */
export interface CellFormatPatch extends RunFormat {
  /** Solid fill colour as 'RRGGBB', or null to clear the fill. */
  fillHex?: string | null;
  /** Alignment overrides; only the keys present are changed. */
  alignment?: XlsxAlignment;
  /** Number format code, e.g. '#,##0.00', or null for General. */
  numFmtCode?: string | null;
}

/** True if the patch asks for nothing. */
export function isEmptyPatch(patch: CellFormatPatch): boolean {
  return Object.values(patch).every((v) => v === undefined);
}

function prefixOf(node: XmlNode): string {
  const i = node.name.indexOf(':');
  return i === -1 ? '' : node.name.slice(0, i + 1);
}

function argb(hex: string): string {
  const clean = hex.replace(/^#/, '').toUpperCase();
  return clean.length === 8 ? clean : `FF${clean}`;
}

export class StyleWriter {
  private readonly styles: XmlNode;
  private readonly prefix: string;

  constructor(stylesXml: XmlNode) {
    this.styles = stylesXml;
    this.prefix = prefixOf(stylesXml);
  }

  /**
   * The style index a cell should point at once `patch` is applied on top of
   * the style it currently uses. Appends to the style table only when the
   * resulting entry is new.
   */
  styleIdFor(baseStyleId: number | undefined, patch: CellFormatPatch): number {
    const cellXfs = this.container('cellXfs');
    const base = baseStyleId !== undefined ? children(cellXfs, 'xf')[baseStyleId] : undefined;
    const xf = base
      ? { name: base.name, attrs: { ...base.attrs }, children: base.children.map(cloneNode), text: '' }
      : createElement(`${this.prefix}xf`, {
          numFmtId: '0',
          fontId: '0',
          fillId: '0',
          borderId: '0',
          xfId: '0',
        });

    if (hasFontChange(patch)) {
      xf.attrs['fontId'] = String(this.internFont(Number(attr(xf, 'fontId') ?? 0), patch));
      xf.attrs['applyFont'] = '1';
    }

    if (patch.fillHex !== undefined) {
      xf.attrs['fillId'] = String(this.internFill(patch.fillHex));
      xf.attrs['applyFill'] = '1';
    }

    if (patch.numFmtCode !== undefined) {
      xf.attrs['numFmtId'] = String(
        patch.numFmtCode === null ? 0 : this.internNumFmt(patch.numFmtCode),
      );
      xf.attrs['applyNumberFormat'] = '1';
    }

    if (patch.alignment) {
      this.applyAlignment(xf, patch.alignment);
      xf.attrs['applyAlignment'] = '1';
    }

    return this.intern(cellXfs, 'xf', xf);
  }

  /** Derive a font from the cell's current one and intern it. */
  private internFont(baseFontId: number, patch: CellFormatPatch): number {
    const fonts = this.container('fonts');
    const base = children(fonts, 'font')[baseFontId];
    const font = base
      ? { name: base.name, attrs: { ...base.attrs }, children: base.children.map(cloneNode), text: '' }
      : createElement(`${this.prefix}font`, {}, [
          createElement(`${this.prefix}sz`, { val: '11' }),
          createElement(`${this.prefix}name`, { val: 'Calibri' }),
        ]);

    const toggle = (name: string, on: boolean | undefined) => {
      if (on === undefined) return;
      const at = font.children.findIndex((c) => localName(c.name) === name);
      if (on && at === -1) font.children.push(createElement(`${this.prefix}${name}`));
      if (!on && at !== -1) font.children.splice(at, 1);
    };
    toggle('b', patch.bold);
    toggle('i', patch.italic);
    toggle('u', patch.underline);
    toggle('strike', patch.strike);

    const setVal = (name: string, val: string) => {
      const existing = font.children.find((c) => localName(c.name) === name);
      if (existing) existing.attrs['val'] = val;
      else font.children.push(createElement(`${this.prefix}${name}`, { val }));
    };
    if (patch.sizePt !== undefined) setVal('sz', String(patch.sizePt));
    if (patch.font !== undefined) {
      setVal('name', patch.font);
      // A theme font would otherwise keep overriding the explicit name.
      font.children = font.children.filter((c) => localName(c.name) !== 'scheme');
    }
    if (patch.colorHex !== undefined) {
      const color = font.children.find((c) => localName(c.name) === 'color');
      if (color) {
        // An indexed/theme colour must go, or it wins over the rgb we set.
        color.attrs = { rgb: argb(patch.colorHex) };
      } else {
        font.children.push(
          createElement(`${this.prefix}color`, { rgb: argb(patch.colorHex) }),
        );
      }
    }

    font.children.sort((a, b) => rank(a, FONT_ORDER) - rank(b, FONT_ORDER));
    return this.intern(fonts, 'font', font);
  }

  /** Intern a solid fill, or return the "no fill" index for null. */
  private internFill(hex: string | null): number {
    if (hex === null) return 0;
    const fills = this.container('fills');
    const fill = createElement(`${this.prefix}fill`, {}, [
      createElement(`${this.prefix}patternFill`, { patternType: 'solid' }, [
        createElement(`${this.prefix}fgColor`, { rgb: argb(hex) }),
        createElement(`${this.prefix}bgColor`, { indexed: '64' }),
      ]),
    ]);
    return this.intern(fills, 'fill', fill);
  }

  /** Intern a custom number format code, returning its numFmtId. */
  private internNumFmt(code: string): number {
    const numFmts = this.container('numFmts');
    for (const nf of children(numFmts, 'numFmt')) {
      if (attr(nf, 'formatCode') === code) return Number(attr(nf, 'numFmtId'));
    }
    let id = FIRST_CUSTOM_NUM_FMT;
    const used = new Set(children(numFmts, 'numFmt').map((nf) => Number(attr(nf, 'numFmtId'))));
    while (used.has(id)) id++;
    numFmts.children.push(
      createElement(`${this.prefix}numFmt`, { numFmtId: String(id), formatCode: code }),
    );
    numFmts.attrs['count'] = String(children(numFmts, 'numFmt').length);
    return id;
  }

  private applyAlignment(xf: XmlNode, align: XlsxAlignment): void {
    let node = xf.children.find((c) => localName(c.name) === 'alignment');
    if (!node) {
      node = createElement(`${this.prefix}alignment`);
      xf.children.unshift(node);
    }
    if (align.horizontal !== undefined) node.attrs['horizontal'] = align.horizontal;
    if (align.vertical !== undefined) node.attrs['vertical'] = align.vertical;
    if (align.wrapText !== undefined) node.attrs['wrapText'] = align.wrapText ? '1' : '0';
  }

  /** Index of an identical existing entry, or the index of the appended one. */
  private intern(container: XmlNode, name: string, node: XmlNode): number {
    const existing = children(container, name);
    const key = serializeNode(node);
    for (let i = 0; i < existing.length; i++) {
      if (serializeNode(existing[i]!) === key) return i;
    }
    container.children.push(node);
    container.attrs['count'] = String(existing.length + 1);
    return existing.length;
  }

  /** A top-level styles container, created in schema order if absent. */
  private container(name: string): XmlNode {
    const existing = child(this.styles, name);
    if (existing) return existing;
    const created = createElement(`${this.prefix}${name}`, { count: '0' });
    const rankNew = STYLESHEET_ORDER.indexOf(name);
    const at = this.styles.children.findIndex((c) => {
      const r = STYLESHEET_ORDER.indexOf(localName(c.name));
      return r !== -1 && r > rankNew;
    });
    if (at === -1) this.styles.children.push(created);
    else this.styles.children.splice(at, 0, created);
    return created;
  }
}

function hasFontChange(patch: CellFormatPatch): boolean {
  return (
    patch.bold !== undefined ||
    patch.italic !== undefined ||
    patch.underline !== undefined ||
    patch.strike !== undefined ||
    patch.sizePt !== undefined ||
    patch.colorHex !== undefined ||
    patch.font !== undefined
  );
}

function rank(node: XmlNode, order: string[]): number {
  const r = order.indexOf(localName(node.name));
  return r === -1 ? order.length : r;
}
