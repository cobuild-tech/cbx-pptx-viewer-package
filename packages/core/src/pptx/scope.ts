/**
 * Resolution scope: the per-part context used while turning XML into the model.
 *
 * Carries the color context plus relationship resolvers (image / hyperlink),
 * because a slide, its layout, and its master each resolve `r:embed`/`r:id`
 * against their own relationships file. Parsers/resolvers take a scope so they
 * never need to know about the OPC package directly.
 */
import type { ColorContext } from './color.js';
import type { XmlNode } from '../oxml/xml.js';

/**
 * Records which XML node a model object was parsed from, plus the part that
 * node lives in. The edit layer needs this because model indices do not
 * correspond to XML indices (empty runs are dropped, `<a:br>`/`<a:fld>` become
 * runs), so the only exact mapping is the one captured at parse time.
 *
 * Because each of a slide's master/layout/slide scopes carries its own part
 * path, this also tells the editor whether a text body is the slide's own
 * (editable) or inherited (read-only).
 */
export type SourceSink = (model: object, node: XmlNode, part: string) => void;

export interface ParseScope {
  colorCtx: ColorContext;
  /** Resolve an `r:embed`/`r:link` id to a media part path. */
  resolveImage(relId: string): string | undefined;
  /** Resolve an `r:id` on a hyperlink to its (usually external) target URL. */
  resolveHyperlink?(relId: string): string | undefined;
  /** Edit support: record the XML node a model object came from. */
  recordSource?(model: object, node: XmlNode): void;
}
