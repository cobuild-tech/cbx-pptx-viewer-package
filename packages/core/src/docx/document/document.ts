/**
 * DocxDocument loader — the top-level entry point.
 *
 * Reads the OPC package, resolves styles + numbering, parses word/document.xml
 * into a paginated model, and returns a {@link DocxDocument} the renderer/viewer
 * consume. Mirrors the PPTX {@link Deck}: it vends object URLs for embedded
 * media and must be `dispose()`d to free them.
 *
 * DOCX pipeline: read (OPC) -> parse (XML -> model) -> paginate (sections) -> render (DOM)
 */
import { OpcPackage } from '../../oxml/package.js';
import { DocxRelType } from '../relTypes.js';
import { StyleTable } from '../styles/styles.js';
import { Numbering } from '../numbering/numbering.js';
import { parseBody, parseBlocks } from './body.js';
import { child, children, path, attr, attrNum, type XmlNode } from '../../oxml/xml.js';
import { twipToPx } from '../units.js';
import type { ThemeFonts } from '../styles/styles.js';
import { parseTheme, type ColorContext } from '../../oxml/drawingml/color.js';
import { formatPageNumber } from '../numbering/numbering.js';
import type { DocxBlock, DocxComment } from '../model.js';
import type { DocxSection, EmbeddedFont } from '../model.js';
import type { DocxSettings, DocxSource, ParseContext } from './context.js';

export type { DocxSource } from './context.js';

const DOCX_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

export class DocxDocument {
  /** Parsed sections; the viewer's paginator flows these into fixed-size pages. */
  sections: DocxSection[];
  /** Fonts embedded in the package (empty in v1 — reserved for parity with Deck). */
  readonly embeddedFonts: EmbeddedFont[];
  private readonly pkg: OpcPackage;
  private readonly docPart: string;
  private readonly urlCache = new Map<string, string>();
  /**
   * Model object -> the XML node it was parsed from. Off the model itself so the
   * render-agnostic types stay pure; a WeakMap means entries die with the
   * sections when the document is rebuilt.
   */
  private sources = new WeakMap<object, DocxSource>();

  private constructor(pkg: OpcPackage, docPart: string, embeddedFonts: EmbeddedFont[]) {
    this.pkg = pkg;
    this.docPart = docPart;
    this.embeddedFonts = embeddedFonts;
    this.sections = [];
  }

  static load(data: ArrayBuffer | Uint8Array): DocxDocument {
    const pkg = OpcPackage.load(data);
    const docPart = pkg.relByType('', DocxRelType.OfficeDocument)?.target;
    if (!docPart) throw new Error('Not a WordprocessingML package: no officeDocument part.');
    if (!pkg.getXml(docPart)) throw new Error('word/document.xml is missing or empty.');

    const doc = new DocxDocument(pkg, docPart, []);
    doc.sections = doc.parse();
    return doc;
  }

  /**
   * Parse (or re-parse) the document part into sections, recording every
   * paragraph's and run's XML source as it goes.
   *
   * Styles and numbering are rebuilt on every call *by design*: `Numbering`
   * carries live counters that `marker()` advances as it walks the body, so
   * reusing one across parses would render an ordered list as "5. 6. 7." the
   * second time round.
   */
  private parse(): DocxSection[] {
    const { pkg, docPart } = this;
    const docXml = pkg.getXml(docPart);
    if (!docXml) return [];

    const stylesPart = pkg.relByType(docPart, DocxRelType.Styles)?.target;
    const numberingPart = pkg.relByType(docPart, DocxRelType.Numbering)?.target;
    const themePart = pkg.relByType(docPart, DocxRelType.Theme)?.target;
    const themeXml = themePart ? pkg.getXml(themePart) : undefined;
    const stylesXml = stylesPart ? pkg.getXml(stylesPart) : undefined;
    const eaLang = attr(path(stylesXml, 'docDefaults/rPrDefault/rPr/lang'), 'eastAsia');
    const styles = StyleTable.parse(stylesXml, readThemeFonts(themeXml, eaLang));
    const numbering = Numbering.parse(numberingPart ? pkg.getXml(numberingPart) : undefined);

    const settingsPart = pkg.relByType(docPart, DocxRelType.Settings)?.target;
    const settingsXml = settingsPart ? pkg.getXml(settingsPart) : undefined;
    const settings = readSettings(settingsXml);
    const colors: ColorContext = { theme: parseTheme(themeXml), clrMap: readColorMapping(settingsXml) };

    const notes = new NoteStore(pkg, docPart, settingsXml);
    const comments = readComments(pkg, docPart);
    const parsedComments = new Map<string, DocxComment>();
    const openComments = new Set<string>();

    const makeCtx = (partPath: string, noteNumber?: string): ParseContext => {
      const c: ParseContext = {
        styles,
        numbering,
        settings,
        colors,
        partPath,
        note: (kind, id) => notes.get(kind, id, (part, num) => makeCtx(part, num)),
        comment: (id) => {
          const cached = parsedComments.get(id);
          if (cached || !comments) return cached;
          const node = comments.byId.get(id);
          if (!node) return undefined;
          const author = attr(node, 'author') ?? '';
          const c: DocxComment = { id, author, blocks: [] };
          const initials = attr(node, 'initials');
          const date = attr(node, 'date');
          if (initials) c.initials = initials;
          if (date) c.date = date;
          parsedComments.set(id, c);
          // Parsed after caching: a comment's own content never references itself.
          c.blocks = makeCtx(comments.part).parseBlocks(node);
          return c;
        },
        openComments,
        ...(noteNumber !== undefined ? { noteNumber } : {}),
        rel: (relId) => (relId ? pkg.resolveRel(partPath, relId) : undefined),
        relsByType: (type) => pkg.relsByType(partPath, type),
        getPartXml: (part) => pkg.getXml(part),
        forPart: (part) => makeCtx(part),
        parseBlocks: (container) => parseBlocks(container, c),
        // The context knows its own part, so header/footer content is
        // automatically distinguishable from the main document body.
        recordSource: (model, node, owner) => {
          this.sources.set(model, { node, part: partPath, ...(owner ? { owner } : {}) });
        },
      };
      return c;
    };

    const body = child(docXml, 'body');
    if (!body) return [];
    const sections = parseBody(body, makeCtx(docPart));

    // Footnotes sit at the foot of their page, under the separator rule.
    const separator = notes.separator('footnote', (part) => makeCtx(part));
    if (separator.length) for (const sec of sections) sec.footnoteSeparator = separator;
    // Endnotes collect after the document's last paragraph, in reference order.
    const endnotes = notes.referenced('endnote');
    const last = sections[sections.length - 1];
    if (last && endnotes.length) {
      last.blocks.push(...notes.separator('endnote', (part) => makeCtx(part)), ...endnotes);
    }
    return sections;
  }

  // ─── Editing ───────────────────────────────────────────────────────────────

  /**
   * The XML node a parsed paragraph or run came from. For a run this is its
   * `<w:t>`, with the owning `<w:r>` on `.owner`.
   */
  sourceOf(model: object): DocxSource | undefined {
    return this.sources.get(model);
  }

  /**
   * True if `model` belongs to the main document part. Header and footer
   * content is parsed from its own part and stays read-only — editing it would
   * change every page that shares it.
   */
  isEditable(model: object): boolean {
    return this.sources.get(model)?.part === this.docPart;
  }

  /** The main document part path (`word/document.xml`). */
  get mainPart(): string {
    return this.docPart;
  }

  /** Parsed XML root of the main document part (the live, mutable node). */
  documentXml(): XmlNode | undefined {
    return this.pkg.getXml(this.docPart);
  }

  /**
   * Re-parse the document from its (possibly mutated) XML, replacing
   * `sections`. `pkg.getXml` returns the cached node the edit layer mutated, so
   * this picks up edits without re-reading the zip.
   */
  rebuild(): DocxSection[] {
    this.sources = new WeakMap();
    this.sections = this.parse();
    return this.sections;
  }

  /** Mark the document part as mutated so export re-serializes it. */
  markDirty(): void {
    this.pkg.markDirty(this.docPart);
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

  /** Re-zip the document, edits included, as .docx bytes. */
  toBytes(): Uint8Array {
    return this.pkg.toBytes();
  }

  /** Re-zip the document, edits included, as a .docx Blob. */
  exportBlob(): Blob {
    return this.pkg.toBlob(DOCX_CONTENT_TYPE);
  }

  /** Raw bytes of a part (e.g. an embedded font), for FontFace registration. */
  fontBytes(part: string): Uint8Array | undefined {
    return this.pkg.getBytes(part);
  }

  /** Object URL for an embedded media part (cached). Browser only. */
  imageUrl(part: string): string | undefined {
    const cached = this.urlCache.get(part);
    if (cached) return cached;
    const bytes = this.pkg.getBytes(part);
    if (!bytes || typeof URL === 'undefined' || typeof Blob === 'undefined') return undefined;
    const type = this.pkg.contentType(part) ?? 'application/octet-stream';
    const url = URL.createObjectURL(new Blob([bytes as unknown as BlobPart], { type }));
    this.urlCache.set(part, url);
    return url;
  }

  /** Revoke all media object URLs. Call when the document is no longer displayed. */
  dispose(): void {
    if (typeof URL !== 'undefined') {
      for (const url of this.urlCache.values()) URL.revokeObjectURL(url);
    }
    this.urlCache.clear();
  }
}

/** Theme script tag for an East Asian language, which picks the scheme's script font. */
const EA_SCRIPT: Record<string, string> = { zh: 'Hans', 'zh-tw': 'Hant', 'zh-hk': 'Hant', ja: 'Jpan', ko: 'Hang' };

/**
 * The major/minor typefaces of a theme part's font scheme: latin, East Asian
 * (the scheme's <a:ea>, or — left empty, as usual — its <a:font script> for the
 * document's East Asian language) and complex script.
 */
function readThemeFonts(theme: XmlNode | undefined, eaLang: string | undefined): ThemeFonts {
  const scheme = path(theme, 'themeElements/fontScheme');
  const lang = (eaLang ?? '').toLowerCase();
  const script = EA_SCRIPT[lang] ?? EA_SCRIPT[lang.split('-')[0] ?? ''];
  const face = (which: string, slot: string) => attr(path(scheme, `${which}/${slot}`), 'typeface') || undefined;
  const ea = (which: string) => {
    const direct = face(which, 'ea');
    if (direct) return direct;
    const sf = children(child(scheme, which), 'font').find((f) => attr(f, 'script') === script);
    return attr(sf, 'typeface') || undefined;
  };
  const out: ThemeFonts = {};
  const set = (key: keyof ThemeFonts, v: string | undefined) => {
    if (v) out[key] = v;
  };
  set('major', face('majorFont', 'latin'));
  set('minor', face('minorFont', 'latin'));
  set('majorEa', ea('majorFont'));
  set('minorEa', ea('minorFont'));
  set('majorCs', face('majorFont', 'cs'));
  set('minorCs', face('minorFont', 'cs'));
  return out;
}

/** Layout settings from word/settings.xml, with Word's defaults. */
function readSettings(settings: XmlNode | undefined): DocxSettings {
  const tab = attrNum(child(settings, 'defaultTabStop'), 'val');
  const mode = children(child(settings, 'compat'), 'compatSetting').find(
    (c) => attr(c, 'name') === 'compatibilityMode',
  );
  const compatMode = Number(attr(mode, 'val'));
  return {
    defaultTabPx: twipToPx(tab && tab > 0 ? tab : 720),
    // No compatibilityMode means a pre-2010 document: Word 2007's rules (12).
    compatMode: Number.isFinite(compatMode) && compatMode > 0 ? compatMode : 12,
    evenAndOddHeaders: (() => {
      const el = child(settings, 'evenAndOddHeaders');
      return el !== undefined && attr(el, 'val') !== '0' && attr(el, 'val') !== 'false';
    })(),
  };
}

/** Theme-scheme names Word's <w:clrSchemeMapping> uses, as DrawingML scheme keys. */
const SCHEME_KEY: Record<string, string> = {
  light1: 'lt1',
  dark1: 'dk1',
  light2: 'lt2',
  dark2: 'dk2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hyperlink: 'hlink',
  followedHyperlink: 'folHlink',
};

/**
 * Where a shape's bg1/tx1/bg2/tx2 point (<w:clrSchemeMapping>); Word's
 * default is the light-background mapping.
 */
function readColorMapping(settings: XmlNode | undefined): Record<string, string> {
  const map: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
  const el = child(settings, 'clrSchemeMapping');
  for (const [attrName, key] of [['bg1', 'bg1'], ['t1', 'tx1'], ['bg2', 'bg2'], ['t2', 'tx2']] as const) {
    const v = attr(el, attrName);
    if (v && SCHEME_KEY[v]) map[key] = SCHEME_KEY[v]!;
  }
  return map;
}

/**
 * Footnotes and endnotes: parsed from their own parts when first referenced,
 * and numbered in the order the body references them — which is how Word
 * numbers them, whatever ids the file uses.
 */
class NoteStore {
  private readonly parts: Record<'footnote' | 'endnote', { path?: string; byId: Map<string, XmlNode> }>;
  private readonly fmt: Record<'footnote' | 'endnote', string>;
  private readonly start: Record<'footnote' | 'endnote', number>;
  private readonly done = new Map<string, { number: string; blocks: DocxBlock[] }>();
  private readonly order: Record<'footnote' | 'endnote', DocxBlock[][]> = { footnote: [], endnote: [] };
  private count = { footnote: 0, endnote: 0 };

  constructor(pkg: OpcPackage, docPart: string, settings: XmlNode | undefined) {
    const load = (type: string, root: string) => {
      const path = pkg.relByType(docPart, type)?.target;
      const byId = new Map<string, XmlNode>();
      const xml = path ? pkg.getXml(path) : undefined;
      for (const n of children(xml, root)) {
        const id = attr(n, 'id');
        if (id !== undefined) byId.set(id, n);
      }
      return { ...(path ? { path } : {}), byId };
    };
    this.parts = { footnote: load(DocxRelType.Footnotes, 'footnote'), endnote: load(DocxRelType.Endnotes, 'endnote') };
    const pr = (name: string) => child(settings, name);
    // Word's defaults: footnotes 1, 2, 3…; endnotes i, ii, iii…
    this.fmt = {
      footnote: attr(child(pr('footnotePr'), 'numFmt'), 'val') ?? 'decimal',
      endnote: attr(child(pr('endnotePr'), 'numFmt'), 'val') ?? 'lowerRoman',
    };
    this.start = {
      footnote: attrNum(child(pr('footnotePr'), 'numStart'), 'val') ?? 1,
      endnote: attrNum(child(pr('endnotePr'), 'numStart'), 'val') ?? 1,
    };
  }

  get(
    kind: 'footnote' | 'endnote',
    id: string,
    ctxFor: (part: string, noteNumber: string) => ParseContext,
  ): { number: string; blocks: DocxBlock[] } | undefined {
    const key = `${kind}:${id}`;
    const cached = this.done.get(key);
    if (cached) return cached;
    const part = this.parts[kind];
    const node = part.byId.get(id);
    if (!node || !part.path) return undefined;
    const number = formatPageNumber(this.start[kind] + this.count[kind]++, this.fmt[kind]);
    const note = { number, blocks: ctxFor(part.path, number).parseBlocks(node) };
    this.done.set(key, note);
    this.order[kind].push(note.blocks);
    return note;
  }

  /** Every note of a kind that was referenced, in reference order. */
  referenced(kind: 'footnote' | 'endnote'): DocxBlock[] {
    return this.order[kind].flat();
  }

  /** The separator rule's content (the note of type "separator"). */
  separator(kind: 'footnote' | 'endnote', ctxFor: (part: string) => ParseContext): DocxBlock[] {
    const part = this.parts[kind];
    if (!part.path) return [];
    for (const n of part.byId.values()) {
      if (attr(n, 'type') === 'separator') return ctxFor(part.path).parseBlocks(n);
    }
    return [];
  }
}

/** comments.xml, indexed by comment id. */
function readComments(pkg: OpcPackage, docPart: string): { part: string; byId: Map<string, XmlNode> } | undefined {
  const part = pkg.relByType(docPart, DocxRelType.Comments)?.target;
  const xml = part ? pkg.getXml(part) : undefined;
  if (!part || !xml) return undefined;
  const byId = new Map<string, XmlNode>();
  for (const c of children(xml, 'comment')) {
    const id = attr(c, 'id');
    if (id !== undefined) byId.set(id, c);
  }
  return { part, byId };
}
