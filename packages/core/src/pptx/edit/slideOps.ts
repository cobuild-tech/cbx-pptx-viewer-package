/**
 * Structural slide edits — operations on the *deck* rather than on the text
 * inside one slide.
 *
 * Deleting a slide is not a single-part change the way typing is. A slide is
 * referenced from four places at once:
 *
 *   1. `<p:sldIdLst>` in presentation.xml — the running order;
 *   2. a relationship in ppt/_rels/presentation.xml.rels;
 *   3. an `Override` in [Content_Types].xml;
 *   4. optionally `<p:custShow>` entries and the `<p14:section>` list, both of
 *      which address slides by the same `sldId`/`r:id` pair.
 *
 * Anything left pointing at a part that is gone makes PowerPoint declare the
 * file corrupt, so all four are cleaned up together. Media the slide referenced
 * is deliberately *not* removed: other slides commonly share the same image
 * part, and an unreferenced media part is harmless (its content type comes from
 * a `Default` extension rule, not an override).
 */
import type { OpcPackage } from '../../oxml/package.js';
import { child, children, attr, localName, type XmlNode } from '../../oxml/xml.js';
import { RelType } from '../relTypes.js';

/** The parts one slide deletion touches, in the order they should be snapshotted. */
export interface SlideDeletionPlan {
  /** Relationship id of the slide within presentation.xml.rels. */
  relId: string;
  /** The slide's `<p:sldId id>` value, used to purge section/custom-show refs. */
  slideId: string | undefined;
  /** Parts rewritten in place (presentation.xml, its rels, content types). */
  changed: string[];
  /** Parts dropped from the package entirely (slide, its rels, notes slide). */
  removed: string[];
}

/**
 * Work out everything a deletion of `slidePart` would touch, without mutating
 * anything. Returns undefined if the slide is not referenced by the
 * presentation (in which case there is nothing coherent to delete).
 */
export function planSlideDeletion(
  pkg: OpcPackage,
  presPart: string,
  slidePart: string,
): SlideDeletionPlan | undefined {
  const rel = pkg
    .relsByType(presPart, RelType.Slide)
    .find((r) => r.mode === 'Internal' && r.target === slidePart);
  if (!rel) return undefined;

  const presXml = pkg.getXml(presPart);
  const sldId = children(child(presXml, 'sldIdLst'), 'sldId').find(
    (n) => attr(n, 'r:id') === rel.id,
  );

  // Each removed part takes its own .rels with it, and those have to be named
  // here too: the caller snapshots this list, and a deletion it cannot put back
  // in full is not undoable.
  const removed: string[] = [];
  for (const part of [slidePart, ...relatedParts(pkg, slidePart)]) {
    removed.push(part);
    const partRels = relsPathOf(part);
    if (pkg.has(partRels)) removed.push(partRels);
  }
  const changed = [presPart, relsPathOf(presPart), CONTENT_TYPES];

  const plan: SlideDeletionPlan = {
    relId: rel.id,
    slideId: attr(sldId, 'id'),
    changed,
    removed,
  };
  return plan;
}

/**
 * Delete a slide. Returns the plan that was carried out, or undefined if the
 * slide is not part of the presentation (nothing is mutated in that case).
 *
 * Callers are responsible for snapshotting {@link SlideDeletionPlan.changed}
 * and `removed` first if they want the edit to be undoable, and for rebuilding
 * the deck afterwards — every model object for the deck's slides is stale once
 * the running order changes.
 */
export function deleteSlide(
  pkg: OpcPackage,
  presPart: string,
  slidePart: string,
): SlideDeletionPlan | undefined {
  const plan = planSlideDeletion(pkg, presPart, slidePart);
  if (!plan) return undefined;

  const presXml = pkg.getXml(presPart);
  if (presXml) {
    purgeSlideRefs(presXml, plan.slideId, plan.relId);
    pkg.markDirty(presPart);
  }
  pkg.removeRelationship(presPart, plan.relId);
  for (const part of plan.removed) pkg.deletePart(part);
  return plan;
}

const CONTENT_TYPES = '[Content_Types].xml';

/** Parts that exist only to serve this slide and die with it. */
function relatedParts(pkg: OpcPackage, slidePart: string): string[] {
  // A notes slide relates *back* to its slide, so leaving it behind would leave
  // a dangling relationship. Nothing else in the package points at it.
  return pkg
    .relsByType(slidePart, RelType.NotesSlide)
    .filter((r) => r.mode === 'Internal')
    .map((r) => r.target);
}

function relsPathOf(partPath: string): string {
  const i = partPath.lastIndexOf('/');
  const dir = i === -1 ? '' : partPath.slice(0, i);
  const base = i === -1 ? partPath : partPath.slice(i + 1);
  return `${dir ? dir + '/' : ''}_rels/${base}.rels`;
}

/**
 * Drop every reference to the slide from presentation.xml, wherever it appears.
 *
 * `<p:sldIdLst>`, `<p:custShow>`'s slide list and the `<p14:section>` list all
 * address a slide the same way — by its numeric `sldId` or by the relationship
 * id — but they sit at different depths and (for sections) under a different
 * namespace inside `<p:extLst>`. Matching on local name plus the id makes this
 * one recursive pass instead of three hand-written traversals, and it keeps
 * working for any future list that follows the same convention.
 */
function purgeSlideRefs(node: XmlNode, slideId: string | undefined, relId: string): void {
  node.children = node.children.filter((c) => {
    const name = localName(c.name);
    if (name === 'sldId' && slideId !== undefined && attr(c, 'id') === slideId) return false;
    // Custom shows list slides as <p:sld r:id="…"> with no numeric id.
    if (name === 'sld' && attr(c, 'r:id') === relId) return false;
    return true;
  });
  for (const c of node.children) purgeSlideRefs(c, slideId, relId);
}
