/**
 * Chart parts on slides. The parsing itself is shared with Word documents and
 * lives in oxml/drawingml/chart; slides resolve its colours in their scope.
 */
import type { XmlNode } from '../../oxml/xml.js';
import { parseChart as parseChartXml } from '../../oxml/drawingml/chart/parse.js';
import type { Chart } from '../model.js';
import type { ParseScope } from '../scope.js';

export function parseChart(chartSpace: XmlNode, scope: ParseScope): Chart | undefined {
  return parseChartXml(chartSpace, scope.colorCtx);
}
