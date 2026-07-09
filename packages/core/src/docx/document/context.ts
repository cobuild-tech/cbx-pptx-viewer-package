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

export interface ParseContext {
  styles: StyleTable;
  numbering: Numbering;
  /** The part these relationships resolve against (document / header / footer). */
  readonly partPath: string;
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
   * When true, anchored (floating) images are hoisted to page-level floats
   * rather than emitted inline. Set while parsing header/footer parts, whose
   * banners are positioned absolutely on the page.
   */
  hoistAnchors?: boolean;
}
