/**
 * Render-agnostic chart model, shared by every format that embeds DrawingML
 * charts (slides and Word documents alike).
 */
import type { Color } from '../color.js';

export type ChartKind = 'bar' | 'line' | 'pie' | 'doughnut' | 'area' | 'scatter';

/** How multi-series plots stack/cluster (`<c:grouping>`). */
export type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';

export interface ChartSeries {
  /** Series display name (legend / tooltip). */
  name?: string;
  /** Y values (bar height, line/area point, pie slice), ordered by category. */
  values: number[];
  /** Scatter/bubble X values, parallel to {@link values}. */
  xValues?: number[];
  /** Series-level fill color (bars/area/line), resolved against the theme. */
  color?: Color;
  /** Per-point colors (pie/doughnut slices, or individual bar `<c:dPt>`). */
  pointColors?: (Color | undefined)[];
  /** Scatter: the points are joined by a line (`scatterStyle` line/lineMarker/smooth*). */
  line?: boolean;
  /** The series draws no point markers (`<c:marker><c:symbol val="none"/>`). */
  noMarkers?: boolean;
}

/**
 * A chart parsed from a `chartN.xml` part. Only the cached values
 * (`numCache`/`strCache`) are read — we render the snapshot, not the live
 * spreadsheet link. Lengths/positions are computed by the renderer in the
 * frame's pixel box.
 */
export interface Chart {
  kind: ChartKind;
  /** Bar charts: true = horizontal bars (`barDir=bar`), false = columns (`col`). */
  barHorizontal?: boolean;
  grouping?: ChartGrouping;
  /** Shared category-axis labels. */
  categories: string[];
  series: ChartSeries[];
  title?: string;
  /** Legend position, or omitted when there is no legend. */
  legend?: 'r' | 'l' | 't' | 'b' | 'tr';
  /** Doughnut hole radius as a fraction 0..1 of the outer radius. */
  holeSize?: number;
  /** Whether the chart shows numeric data labels on points. */
  showValueLabels?: boolean;
  /** Scatter: the X value axis has major gridlines (vertical lines). */
  xGridlines?: boolean;
}
