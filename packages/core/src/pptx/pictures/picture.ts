/**
 * Picture resolver: `<p:pic>` -> {@link PictureShape}.
 *
 * Resolves the image relationship to a media part, inherits geometry from the
 * matching layout/master placeholder (picture placeholders behave like text
 * ones), and captures any source-rectangle crop. A non-rectangular preset on
 * the picture clips the image (e.g. a photo cropped into a circle).
 */
import { child, attr, attrNum, type XmlNode } from '../../oxml/xml.js';
import type { PictureShape } from '../model.js';
import type { ParseScope } from '../scope.js';
import {
  type SlideBuildCtx,
  resolveTransform,
  resolveGeometry,
  resolveStroke,
} from '../shapes/props.js';
import { placeholderOf, matchPlaceholder } from '../shapes/placeholders.js';

export function buildPic(pic: XmlNode, ctx: SlideBuildCtx, scope: ParseScope): PictureShape | null {
  const spPr = child(pic, 'spPr');
  const blipFill = child(pic, 'blipFill');
  const blip = child(blipFill, 'blip');
  const rId = attr(blip, 'embed') ?? attr(blip, 'link');
  const part = rId ? scope.resolveImage(rId) : undefined;
  if (!part) return null;

  const shape: PictureShape = { kind: 'picture', part, fill: { type: 'none' } };

  // Picture placeholders inherit geometry from the layout/master, like text ones.
  const ph = placeholderOf(pic);
  const layoutPh = ph ? matchPlaceholder(ph, ctx.layoutPhs) : undefined;
  const masterPh = ph ? matchPlaceholder(ph, ctx.masterPhs) : undefined;

  const layoutSpPr = child(layoutPh?.sp, 'spPr');
  const masterSpPr = child(masterPh?.sp, 'spPr');
  const layoutStyle = child(layoutPh?.sp, 'style');
  const masterStyle = child(masterPh?.sp, 'style');

  const transform = resolveTransform(spPr, layoutSpPr, masterSpPr);
  if (transform) shape.transform = transform;
  if (ph) shape.placeholder = ph;

  // A non-rectangular preset clips the image (e.g. picture cropped to a circle).
  if (child(spPr, 'prstGeom') || child(spPr, 'custGeom') ||
      child(layoutSpPr, 'prstGeom') || child(layoutSpPr, 'custGeom') ||
      child(masterSpPr, 'prstGeom') || child(masterSpPr, 'custGeom')) {
    shape.geom = resolveGeometry(spPr, layoutSpPr, masterSpPr);
  }

  const stroke = resolveStroke(spPr, child(pic, 'style'), layoutSpPr, layoutStyle, masterSpPr, masterStyle, ctx);
  if (stroke) shape.stroke = stroke;

  const srcRect = child(blipFill, 'srcRect');
  if (srcRect) {
    shape.crop = {
      l: (attrNum(srcRect, 'l') ?? 0) / 100000,
      t: (attrNum(srcRect, 't') ?? 0) / 100000,
      r: (attrNum(srcRect, 'r') ?? 0) / 100000,
      b: (attrNum(srcRect, 'b') ?? 0) / 100000,
    };
  }
  return shape;
}
