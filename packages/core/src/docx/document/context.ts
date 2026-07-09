/**
 * Shared parsing context threaded through the DOCX parse slices. Holds the
 * resolved style table, numbering state, and a relationship resolver for the
 * main document part (images, hyperlinks). Lives in its own module so the
 * parse slices can import it without a cycle through body.ts.
 */
import type { Relationship } from '../../oxml/package.js';
import type { StyleTable } from '../styles/styles.js';
import type { Numbering } from '../numbering/numbering.js';

export interface ParseContext {
  styles: StyleTable;
  numbering: Numbering;
  /** Resolve an r:id / r:embed relationship on the main document part. */
  rel(relId: string | undefined): Relationship | undefined;
}
