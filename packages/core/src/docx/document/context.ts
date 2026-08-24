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
import type { DocxBlock } from '../model.js';

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

export interface ParseContext {
  styles: StyleTable;
  numbering: Numbering;
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
  /** Parsed XML of a package part by path (for header/footer parts). */
  getPartXml(part: string): XmlNode | undefined;
  /** A context whose relationships resolve against a different part. */
  forPart(partPath: string): ParseContext;
  /**
   * Edit support: record the XML node a model object came from. The context
   * knows its own {@link partPath}, so headers/footers are distinguishable from
   * the main document — which is what makes them read-only in the editor.
   */
  recordSource?(model: object, node: XmlNode, owner?: XmlNode): void;
}
