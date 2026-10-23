/**
 * Shared parsing context threaded through the DOCX parse slices. Holds the
 * resolved style table, numbering state, and a relationship resolver for the
 * main document part (images, hyperlinks). Lives in its own module so the
 * parse slices can import it without a cycle through body.ts.
 */
import type { Relationship } from '../../oxml/package.js';
import type { XmlNode } from '../../oxml/xml.js';
import type { StyleTable } from '../styles/styles.js';
import type { Numbering } from '../numbering/numbering.js';
import type { DocxBlock, DocxComment } from '../model.js';
import type { ColorContext } from '../../oxml/drawingml/color.js';

/**
 * Where a model object came from in the XML.
 *
 * A DocxRun records its `<w:t>` node, not its `<w:r>` — one `<w:r>` emits one
 * run per `<w:t>` child, so the run element alone cannot identify a run. The
 * owning `<w:r>` is kept alongside because that is where `<w:rPr>` lives.
 */
export interface DocxSource {
  node: XmlNode;
  /** The `<w:r>` a `<w:t>` belongs to, for rPr cloning. */
  owner?: XmlNode;
  /** The `<w:br>`/`<w:cr>` behind a run's {@link DocxRun.breakBefore}. */
  br?: XmlNode;
  /** The `<w:tab>`/`<w:ptab>` behind a run's {@link DocxRun.tabBefore}. */
  tab?: XmlNode;
  /** Package part the node lives in (document, header or footer). */
  part: string;
}

/**
 * Records which XML node a model object was parsed from. The edit layer needs
 * this because model indices do not match XML indices — `logicalChildren`
 * flattens `<w:sdt>`/`<w:smartTag>`/`<w:fldSimple>`, and one `<w:r>` can emit
 * several runs. Capturing at parse time is the only exact mapping.
 */
export type DocxSourceSink = (model: object, src: DocxSource) => void;

/** Document-wide layout settings from word/settings.xml. */
export interface DocxSettings {
  /** Interval of the default tab stops (<w:defaultTabStop>), in px. */
  defaultTabPx: number;
  /**
   * <w:compatSetting w:name="compatibilityMode">: 15 for Word 2013+, lower for
   * documents laid out with an older Word's rules.
   */
  compatMode: number;
  /** <w:evenAndOddHeaders>: even-numbered pages use the "even" header/footer. */
  evenAndOddHeaders: boolean;
}

export interface ParseContext {
  styles: StyleTable;
  numbering: Numbering;
  settings: DocxSettings;
  /** Theme colours and the document's bg/tx mapping, for DrawingML shapes. */
  colors: ColorContext;
  /**
   * A footnote or endnote by id: its display number (assigned in reference
   * order, as Word numbers them) and its parsed content. Undefined for a
   * missing note.
   */
  note?(kind: 'footnote' | 'endnote', id: string): { number: string; blocks: DocxBlock[] } | undefined;
  /** Inside a note's own content: the number its <w:footnoteRef/> repeats. */
  readonly noteNumber?: string;
  /** A comment by id, parsed from comments.xml. */
  comment?(id: string): DocxComment | undefined;
  /**
   * Comments whose range is open at this point of the walk. Ranges cross
   * paragraph boundaries, so this lives on the context, shared by the walk.
   */
  openComments?: Set<string>;
  /** The part these relationships resolve against (document / header / footer). */
  readonly partPath: string;
  /**
   * Parse the block flow (paragraphs/tables) inside a container, against this
   * context's part. Used to render text-box (<w:txbxContent>) content nested in
   * a shape without an import cycle through body.ts.
   */
  parseBlocks(container: XmlNode | undefined): DocxBlock[];
  /**
   * Resolve an r:id / r:embed relationship against {@link partPath}. Header and
   * footer parts have their OWN rels, so image/hyperlink ids there must resolve
   * against the header/footer part, not word/document.xml.
   */
  rel(relId: string | undefined): Relationship | undefined;
  /** Every relationship of a type from {@link partPath}. */
  relsByType(type: string): Relationship[];
  /** Parsed XML of a package part by path (for header/footer parts). */
  getPartXml(part: string): XmlNode | undefined;
  /** A context whose relationships resolve against a different part. */
  forPart(partPath: string): ParseContext;
  /**
   * Edit support: record the XML node a model object came from. The context
   * knows its own {@link partPath}, so headers/footers are distinguishable from
   * the main document — which is what makes them read-only in the editor.
   */
  recordSource?(model: object, node: XmlNode, owner?: XmlNode, pre?: { br?: XmlNode; tab?: XmlNode }): void;
}
