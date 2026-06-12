/**
 * DOM renderer: {@link Slide} model -> HTML elements.
 *
 * Every slide renders into a div sized to the slide's base pixel space; the
 * viewer scales that div to fit the viewport. Shapes are absolutely positioned.
 * Fills use native CSS (solid / linear-gradient / background-image) with a
 * `clip-path` for non-rectangular geometry; outlines are drawn as an SVG path
 * overlay so stroke width hugs the shape edge.
 */
import type {
  Slide,
  SlideSize,
  Shape,
  Fill,
  Stroke,
  Transform,
  TextBody,
  Paragraph,
  TextRun,
  Bullet,
  PresetShape,
  PictureShape,
  GroupShape,
  ConnectorShape,
  FrameShape,
  Table,
} from '../model.js';
import { ptToPx } from '../units.js';
import { colorToCss } from '../resolve/color.js';
import { presetPath, OPEN_PRESETS } from '../geometry/presets.js';

const SVG_NS = 'http://www.w3.org/2000/svg';

export interface RenderDeps {
  /** Resolve a media part path to a displayable URL. */
  imageUrl(part: string): string | undefined;
}

/** Render one slide into a positioned div sized in the slide's base px space. */
export function renderSlide(slide: Slide, size: SlideSize, deps: RenderDeps): HTMLDivElement {
  const root = document.createElement('div');
  root.className = 'pptx-slide';
  root.style.position = 'relative';
  root.style.width = `${size.wPx}px`;
  root.style.height = `${size.hPx}px`;
  root.style.overflow = 'hidden';
  applyFillBackground(root, slide.background, deps);

  for (const shape of slide.shapes) {
    const el = renderShape(shape, deps);
    if (el) root.appendChild(el);
  }
  return root;
}

function renderShape(shape: Shape, deps: RenderDeps): HTMLElement | null {
  switch (shape.kind) {
    case 'shape':
      return renderPreset(shape, deps);
    case 'picture':
      return renderPicture(shape, deps);
    case 'group':
      return renderGroup(shape, deps);
    case 'connector':
      return renderConnector(shape, deps);
    case 'frame':
      return renderFrame(shape, deps);
  }
}

/** Position + rotate/flip a shape container per its transform. */
function positioned(transform: Transform | undefined): HTMLDivElement {
  const el = document.createElement('div');
  el.style.position = 'absolute';
  if (transform) {
    el.style.left = `${transform.x}px`;
    el.style.top = `${transform.y}px`;
    el.style.width = `${transform.w}px`;
    el.style.height = `${transform.h}px`;
    const parts: string[] = [];
    if (transform.rot) parts.push(`rotate(${transform.rot}deg)`);
    if (transform.flipH || transform.flipV) {
      parts.push(`scale(${transform.flipH ? -1 : 1}, ${transform.flipV ? -1 : 1})`);
    }
    if (parts.length) {
      el.style.transform = parts.join(' ');
      el.style.transformOrigin = 'center';
    }
  }
  return el;
}

function renderPreset(shape: PresetShape, deps: RenderDeps): HTMLElement {
  const el = positioned(shape.transform);
  const w = shape.transform?.w ?? 0;
  const h = shape.transform?.h ?? 0;

  if (shape.geom.type === 'preset') {
    const open = OPEN_PRESETS.has(shape.geom.preset);
    const d = presetPath(shape.geom.preset, w, h, shape.geom.adjust);
    if (!open) {
      const fillLayer = document.createElement('div');
      fillLayer.style.position = 'absolute';
      fillLayer.style.inset = '0';
      applyFillBackground(fillLayer, shape.fill, deps);
      if (shape.geom.preset !== 'rect') fillLayer.style.clipPath = `path('${d}')`;
      el.appendChild(fillLayer);
    }
    if (shape.stroke) el.appendChild(strokeOverlay(d, w, h, shape.stroke, false));
  } else {
    // Custom geometry: render each subpath in a scaled SVG.
    for (const p of shape.geom.paths) {
      el.appendChild(customGeomSvg(p.d, p.w, p.h, w, h, shape.fill, shape.stroke));
    }
  }

  if (shape.text) el.appendChild(renderTextBody(shape.text, deps));
  return el;
}

function strokeOverlay(
  d: string,
  w: number,
  h: number,
  stroke: Stroke,
  fill: boolean,
): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('width', `${w}`);
  svg.setAttribute('height', `${h}`);
  svg.style.position = 'absolute';
  svg.style.inset = '0';
  svg.style.overflow = 'visible';
  const path = document.createElementNS(SVG_NS, 'path');
  path.setAttribute('d', d);
  path.setAttribute('fill', fill ? colorToCss(stroke.color) : 'none');
  if (!fill) {
    path.setAttribute('stroke', colorToCss(stroke.color));
    path.setAttribute('stroke-width', `${stroke.width}`);
    if (stroke.dash) path.setAttribute('stroke-dasharray', stroke.dash.join(','));
    if (stroke.cap) path.setAttribute('stroke-linecap', stroke.cap);
  }
  svg.appendChild(path);
  return svg;
}

function customGeomSvg(
  d: string,
  pw: number,
  ph: number,
  w: number,
  h: number,
  fill: Fill,
  stroke: Stroke | undefined,
): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('width', `${w}`);
  svg.setAttribute('height', `${h}`);
  svg.setAttribute('viewBox', `0 0 ${pw || w} ${ph || h}`);
  svg.setAttribute('preserveAspectRatio', 'none');
  svg.style.position = 'absolute';
  svg.style.inset = '0';
  // Geometry must stay within the shape box; clipping prevents an imperfect
  // path (e.g. a spiral with an off arc) from drawing lines across the slide.
  svg.style.overflow = 'hidden';
  const path = document.createElementNS(SVG_NS, 'path');
  path.setAttribute('d', d);
  path.setAttribute('fill', fill.type === 'solid' ? colorToCss(fill.color) : 'none');
  if (stroke) {
    path.setAttribute('stroke', colorToCss(stroke.color));
    path.setAttribute('stroke-width', `${stroke.width}`);
    path.setAttribute('vector-effect', 'non-scaling-stroke');
    if (stroke.dash) path.setAttribute('stroke-dasharray', stroke.dash.join(','));
  }
  svg.appendChild(path);
  return svg;
}

function renderConnector(shape: ConnectorShape, _deps: RenderDeps): HTMLElement {
  const el = positioned(shape.transform);
  const w = shape.transform?.w ?? 0;
  const h = shape.transform?.h ?? 0;
  if (shape.stroke) {
    const d =
      shape.geom.type === 'preset'
        ? presetPath(shape.geom.preset, w, h, shape.geom.adjust)
        : shape.geom.paths[0]?.d ?? `M0,0 L${w},${h}`;
    el.appendChild(strokeOverlay(d, w, h, shape.stroke, false));
  }
  return el;
}

function renderPicture(shape: PictureShape, deps: RenderDeps): HTMLElement | null {
  const url = deps.imageUrl(shape.part);
  if (!url) return null;
  const el = positioned(shape.transform);
  el.style.overflow = 'hidden';

  // Clip the image to a non-rectangular preset (e.g. cropped into a circle).
  if (shape.geom?.type === 'preset' && shape.geom.preset !== 'rect') {
    const w = shape.transform?.w ?? 0;
    const h = shape.transform?.h ?? 0;
    el.style.clipPath = `path('${presetPath(shape.geom.preset, w, h, shape.geom.adjust)}')`;
  }

  const img = document.createElement('img');
  img.src = url;
  img.alt = '';
  img.draggable = false;
  img.style.position = 'absolute';
  img.style.display = 'block';

  const crop = shape.crop;
  if (crop && (crop.l || crop.t || crop.r || crop.b)) {
    const visW = Math.max(0.001, 1 - crop.l - crop.r);
    const visH = Math.max(0.001, 1 - crop.t - crop.b);
    img.style.width = `${100 / visW}%`;
    img.style.height = `${100 / visH}%`;
    img.style.left = `${(-crop.l / visW) * 100}%`;
    img.style.top = `${(-crop.t / visH) * 100}%`;
  } else {
    img.style.width = '100%';
    img.style.height = '100%';
    img.style.objectFit = 'fill';
  }
  el.appendChild(img);
  if (shape.stroke) {
    el.style.outline = `${shape.stroke.width}px solid ${colorToCss(shape.stroke.color)}`;
    el.style.outlineOffset = `-${shape.stroke.width}px`;
  }
  return el;
}

function renderGroup(shape: GroupShape, deps: RenderDeps): HTMLElement {
  const el = positioned(shape.transform);
  const inner = document.createElement('div');
  inner.style.position = 'absolute';
  inner.style.left = '0';
  inner.style.top = '0';
  inner.style.transformOrigin = '0 0';

  // Map the child coordinate space (chOff/chExt) onto the group box (off/ext).
  // Guard against a missing/degenerate child extent producing an enormous scale.
  const co = shape.childOffset;
  const sx = co.w > 1 && shape.transform ? shape.transform.w / co.w : 1;
  const sy = co.h > 1 && shape.transform ? shape.transform.h / co.h : 1;
  inner.style.transform = `scale(${sx}, ${sy}) translate(${-co.x}px, ${-co.y}px)`;

  for (const childShape of shape.children) {
    const childEl = renderShape(childShape, deps);
    if (childEl) inner.appendChild(childEl);
  }
  el.appendChild(inner);
  return el;
}

function renderFrame(shape: FrameShape, deps: RenderDeps): HTMLElement {
  const el = positioned(shape.transform);
  if (shape.frameType === 'table' && shape.table) {
    el.appendChild(renderTable(shape.table, deps));
  } else {
    // Charts / diagrams / unknown: a labeled placeholder so it isn't invisible.
    el.style.display = 'flex';
    el.style.alignItems = 'center';
    el.style.justifyContent = 'center';
    el.style.border = '1px dashed #bbb';
    el.style.color = '#999';
    el.style.font = '12px system-ui, sans-serif';
    el.textContent = `[${shape.frameType}]`;
  }
  return el;
}

function renderTable(table: Table, deps: RenderDeps): HTMLTableElement {
  const tbl = document.createElement('table');
  tbl.style.borderCollapse = 'collapse';
  tbl.style.tableLayout = 'fixed';
  tbl.style.width = '100%';

  const colgroup = document.createElement('colgroup');
  for (const w of table.colWidths) {
    const col = document.createElement('col');
    col.style.width = `${w}px`;
    colgroup.appendChild(col);
  }
  tbl.appendChild(colgroup);

  table.rows.forEach((row, r) => {
    const tr = document.createElement('tr');
    // The XML row height is a minimum; the row grows to fit its content,
    // matching PowerPoint (cell text is rendered in normal flow below).
    tr.style.height = `${table.rowHeights[r] ?? 0}px`;
    for (const cell of row) {
      if (cell === null) continue;
      const td = document.createElement('td');
      td.style.padding = '0';
      td.style.verticalAlign = cell.text ? anchorToValign(cell.text.anchor) : 'top';
      if (cell.colSpan > 1) td.colSpan = cell.colSpan;
      if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
      applyFillBackground(td, cell.fill, deps);
      const b = cell.borders;
      if (b.l) td.style.borderLeft = cssBorder(b.l);
      if (b.r) td.style.borderRight = cssBorder(b.r);
      if (b.t) td.style.borderTop = cssBorder(b.t);
      if (b.b) td.style.borderBottom = cssBorder(b.b);
      if (cell.text) td.appendChild(renderTextBody(cell.text, deps, true));
      tr.appendChild(td);
    }
    tbl.appendChild(tr);
  });
  return tbl;
}

function cssBorder(s: Stroke): string {
  return `${s.width}px solid ${colorToCss(s.color)}`;
}

/**
 * Render a text body. In the default (shape) mode the box is absolutely
 * positioned and uses flex for vertical anchoring. In `flow` mode (table cells)
 * it's a normal-flow block with the insets as padding, so the cell — and its
 * table row — grows to fit the content, just like PowerPoint.
 */
function renderTextBody(body: TextBody, _deps: RenderDeps, flow = false): HTMLDivElement {
  const box = document.createElement('div');
  box.style.boxSizing = 'border-box';
  if (flow) {
    box.style.padding = `${body.insets.t}px ${body.insets.r}px ${body.insets.b}px ${body.insets.l}px`;
  } else {
    box.style.position = 'absolute';
    box.style.left = `${body.insets.l}px`;
    box.style.top = `${body.insets.t}px`;
    box.style.right = `${body.insets.r}px`;
    box.style.bottom = `${body.insets.b}px`;
    box.style.display = 'flex';
    box.style.flexDirection = 'column';
    box.style.justifyContent =
      body.anchor === 'ctr' ? 'center' : body.anchor === 'bottom' ? 'flex-end' : 'flex-start';
    // PowerPoint shows text that overflows its box (the "do not autofit" default)
    // rather than clipping it; only the slide edge clips.
    box.style.overflow = 'visible';
  }

  // Track auto-number counters per level.
  const counters: number[] = [];
  for (const para of body.paragraphs) {
    box.appendChild(renderParagraph(para, body, counters));
  }
  return box;
}

function anchorToValign(anchor: TextBody['anchor']): string {
  return anchor === 'ctr' ? 'middle' : anchor === 'bottom' ? 'bottom' : 'top';
}

function renderParagraph(para: Paragraph, body: TextBody, counters: number[]): HTMLDivElement {
  const p = document.createElement('div');
  p.style.whiteSpace = 'pre-wrap';
  p.style.margin = '0';
  // Keep natural height when text exceeds the box; PowerPoint clips/overflows
  // rather than letting flexbox compress paragraphs on top of each other.
  p.style.flexShrink = '0';
  p.style.boxSizing = 'border-box';
  if (para.align) {
    p.style.textAlign =
      para.align === 'ctr' ? 'center' : para.align === 'r' ? 'right' : para.align === 'just' ? 'justify' : 'left';
  }
  if (para.marginLeftPx !== undefined) p.style.marginLeft = `${para.marginLeftPx}px`;
  if (para.indentPx !== undefined) p.style.textIndent = `${para.indentPx}px`;
  if (para.spaceBeforePt !== undefined) p.style.marginTop = `${ptToPx(para.spaceBeforePt)}px`;
  if (para.spaceAfterPt !== undefined) p.style.marginBottom = `${ptToPx(para.spaceAfterPt)}px`;
  if (para.lineSpacingPct !== undefined) p.style.lineHeight = `${para.lineSpacingPct}`;
  else if (para.lineSpacingPt !== undefined) p.style.lineHeight = `${ptToPx(para.lineSpacingPt)}px`;

  const bulletStr = bulletText(para.bullet, para.level, counters);
  if (bulletStr) {
    const b = document.createElement('span');
    b.textContent = bulletStr + ' ';
    if (para.bullet && 'color' in para.bullet && para.bullet.color) {
      b.style.color = colorToCss(para.bullet.color);
    }
    if (para.bullet && para.bullet.type === 'char' && para.bullet.font) {
      b.style.fontFamily = para.bullet.font;
    }
    p.appendChild(b);
  }

  if (para.runs.length === 0) {
    // Preserve an empty line's height with a zero-width space.
    p.appendChild(document.createTextNode('​'));
  }
  for (const run of para.runs) {
    p.appendChild(renderRun(run, body.fontScale));
  }
  return p;
}

function bulletText(bullet: Bullet | undefined, level: number, counters: number[]): string {
  if (!bullet || bullet.type === 'none') {
    // No explicit bullet: reset deeper counters but render nothing.
    return '';
  }
  if (bullet.type === 'char') return bullet.char;
  // Auto-numbered: maintain a per-level counter.
  counters[level] = (counters[level] ?? (bullet.startAt ?? 1) - 1) + 1;
  for (let i = level + 1; i < counters.length; i++) counters[i] = 0;
  const n = counters[level];
  if (bullet.scheme.includes('Paren')) return `${n})`;
  if (bullet.scheme.includes('alphaUc')) return `${String.fromCharCode(64 + n)}.`;
  if (bullet.scheme.includes('alphaLc')) return `${String.fromCharCode(96 + n)}.`;
  return `${n}.`;
}

function renderRun(run: TextRun, fontScale: number | undefined): HTMLElement {
  const span = document.createElement('span');
  span.textContent = run.text;
  if (run.bold) span.style.fontWeight = 'bold';
  if (run.italic) span.style.fontStyle = 'italic';
  const decorations: string[] = [];
  if (run.underline) decorations.push('underline');
  if (run.strike) decorations.push('line-through');
  if (decorations.length) span.style.textDecoration = decorations.join(' ');
  if (run.sizePt !== undefined) {
    const pt = fontScale ? run.sizePt * fontScale : run.sizePt;
    span.style.fontSize = `${ptToPx(pt)}px`;
  }
  if (run.color) span.style.color = colorToCss(run.color);
  if (run.font) span.style.fontFamily = `"${run.font}", Arial, Helvetica, sans-serif`;
  if (run.highlight) span.style.backgroundColor = colorToCss(run.highlight);
  if (run.letterSpacingPt) span.style.letterSpacing = `${ptToPx(run.letterSpacingPt)}px`;
  if (run.caps === 'all') span.style.textTransform = 'uppercase';
  else if (run.caps === 'small') span.style.fontVariant = 'small-caps';
  if (run.baseline) {
    span.style.verticalAlign = run.baseline > 0 ? 'super' : 'sub';
    span.style.fontSize = span.style.fontSize || 'smaller';
  }

  if (run.hyperlink) {
    const a = document.createElement('a');
    a.href = run.hyperlink;
    a.target = '_blank';
    a.rel = 'noopener noreferrer';
    a.style.color = 'inherit';
    a.appendChild(span);
    return a;
  }
  return span;
}

/** Apply a {@link Fill} to an element's CSS background. */
function applyFillBackground(el: HTMLElement, fill: Fill, deps: RenderDeps): void {
  switch (fill.type) {
    case 'solid':
      el.style.background = colorToCss(fill.color);
      break;
    case 'gradient': {
      const stops = fill.stops.map((s) => `${colorToCss(s.color)} ${(s.pos * 100).toFixed(1)}%`);
      el.style.background = fill.radial
        ? `radial-gradient(${stops.join(',')})`
        : `linear-gradient(${fill.angle ?? 0}deg, ${stops.join(',')})`;
      break;
    }
    case 'image': {
      const url = deps.imageUrl(fill.part);
      if (url) {
        el.style.backgroundImage = `url("${url}")`;
        el.style.backgroundSize = 'cover';
        el.style.backgroundPosition = 'center';
      }
      break;
    }
    case 'none':
      break;
  }
}
