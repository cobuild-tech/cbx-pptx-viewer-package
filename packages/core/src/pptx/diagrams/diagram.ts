/**
 * SmartArt diagram resolver.
 *
 * A diagram `graphicFrame` references a data model (`dgm:relIds/@r:dm`); the
 * actual shapes are PowerPoint's pre-laid-out drawing (`diagrams/drawingN.xml`,
 * root `<dsp:drawing>`), linked from the slide by a `diagramDrawing` rel. That
 * drawing is a normal shape tree (`dsp:sp` etc., which our local-name lookups
 * read just like `p:sp`), authored in the frame's own coordinate space (canvas
 * origin 0,0, extent = the frame's ext). So we hand its `spTree` straight to the
 * shape pipeline and render the result inside the frame box — no scaling needed.
 *
 * This resolver only locates the drawing's `spTree` + a scope for its media; the
 * caller (buildFrame) runs `buildShapes`, keeping the dependency graph acyclic.
 */
import { child, attr, type XmlNode } from '../../oxml/xml.js';
import type { ParseScope } from '../scope.js';
import type { SlideBuildCtx } from '../shapes/props.js';
import { RelType } from '../relTypes.js';

export function resolveDiagramDrawing(
  graphicData: XmlNode | undefined,
  ctx: SlideBuildCtx,
): { spTree: XmlNode; scope: ParseScope } | undefined {
  const drawings = ctx.parts.relTargetsByType(RelType.DiagramDrawing);
  if (drawings.length === 0) return undefined;

  // Correlate the drawing to this frame's data model when a slide carries more
  // than one diagram (data1.xml <-> drawing1.xml, by trailing file number).
  const relIds = child(graphicData, 'relIds');
  const dataPart = relIds ? ctx.parts.partForRel(attr(relIds, 'r:dm') ?? '') : undefined;
  const drawingPart = pickDrawing(drawings, dataPart);

  const spTree = child(ctx.parts.xml(drawingPart), 'spTree');
  if (!spTree) return undefined;
  return { spTree, scope: ctx.parts.scopeFor(drawingPart) };
}

function pickDrawing(drawings: string[], dataPart: string | undefined): string {
  if (drawings.length === 1 || !dataPart) return drawings[0]!;
  const n = numStem(dataPart);
  return drawings.find((d) => numStem(d) === n) ?? drawings[0]!;
}

function numStem(part: string): string {
  const m = part.match(/(\d+)\.xml$/i);
  return m ? m[1]! : '';
}
