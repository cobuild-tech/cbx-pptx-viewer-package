/**
 * Resolution scope: the per-part context used while turning XML into the model.
 *
 * Carries the color context plus relationship resolvers (image / hyperlink),
 * because a slide, its layout, and its master each resolve `r:embed`/`r:id`
 * against their own relationships file. Parsers/resolvers take a scope so they
 * never need to know about the OPC package directly.
 */
import type { ColorContext } from './color.js';

export interface ParseScope {
  colorCtx: ColorContext;
  /** Resolve an `r:embed`/`r:link` id to a media part path. */
  resolveImage(relId: string): string | undefined;
  /** Resolve an `r:id` on a hyperlink to its (usually external) target URL. */
  resolveHyperlink?(relId: string): string | undefined;
}
